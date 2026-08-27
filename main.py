import os
import asyncio
import json
from datetime import datetime
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile, ReplyKeyboardMarkup, KeyboardButton
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from openai import AsyncOpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import httpx

BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
REPLICATE_API_TOKEN = os.getenv("REPLICATE_API_TOKEN")
ADMIN_IDS = [909828109]

client = AsyncOpenAI(api_key=XAI_API_KEY, base_url="https://api.x.ai/v1")
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
users_db = {}

PLAN_LIMITS = {"premium": 15}

class Form(StatesGroup):
    waiting_topic = State()
    waiting_slides = State()
    waiting_style = State()
    waiting_confirm = State()
    waiting_extra = State()

def get_user(uid):
    if uid not in users_db:
        users_db[uid] = {
            "name": "",
            "plan": "premium",
            "generations": 0,
            "history": []
        }
    return users_db[uid]

def can_generate(uid):
    u = get_user(uid)
    return u["generations"] < PLAN_LIMITS.get(u["plan"], 15)

async def ask_grok(prompt: str) -> str:
    try:
        r = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный создатель презентаций. Отвечай только на русском."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=4000
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"Ошибка: {e}"

async def generate_image(prompt: str, path: str) -> bool:
    try:
        headers = {
            "Authorization": f"Bearer {REPLICATE_API_TOKEN}",
            "Content-Type": "application/json"
        }
        async with httpx.AsyncClient(timeout=120) as http:
            create = await http.post(
                "https://api.replicate.com/v1/models/black-forest-labs/flux-schnell/predictions",
                headers=headers,
                json={"input": {"prompt": prompt, "aspect_ratio": "16:9"}}
            )
            print("Replicate create:", create.status_code)
            create.raise_for_status()
            get_url = create.json()["urls"]["get"]

            for _ in range(30):
                await asyncio.sleep(2)
                check = await http.get(get_url, headers=headers)
                info = check.json()
                status = info.get("status")
                print("Replicate status:", status)
                if status == "succeeded":
                    out = info.get("output")
                    img_url = out[0] if isinstance(out, list) else out
                    img = await http.get(str(img_url))
                    img.raise_for_status()
                    with open(path, "wb") as f:
                        f.write(img.content)
                    print("Image saved:", path)
                    return True
                if status in ("failed", "canceled"):
                    print("Replicate failed:", info)
                    return False
        print("Image generation timeout")
        return False
    except Exception as e:
        print("Image generation error:", e)
        return False

def add_placeholder(slide, left, top, width, height, text):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(60, 60, 70)
    shape.line.color.rgb = RGBColor(120, 120, 130)
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor(200, 200, 210)
    p.alignment = PP_ALIGN.CENTER

def main_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="📊 Сделать презентацию")],
        [KeyboardButton(text="📁 Моя история"), KeyboardButton(text="ℹ️ Мой тариф")]
    ], resize_keyboard=True)

def slides_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="5 слайдов"), KeyboardButton(text="8 слайдов")],
        [KeyboardButton(text="10 слайдов")]
    ], resize_keyboard=True)

def style_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="Тёмный")],
        [KeyboardButton(text="Светлый")],
        [KeyboardButton(text="Синий")]
    ], resize_keyboard=True)

def confirm_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="✅ Делать полную версию")],
        [KeyboardButton(text="➕ Добавить информацию")],
        [KeyboardButton(text="✏️ Изменить тему")]
    ], resize_keyboard=True)

@dp.message(Command("start"))
async def cmd_start(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    u["name"] = m.from_user.first_name or "друг"
    await m.answer(
        f"Привет, {u['name']}!\n\nЯ делаю презентации с текстом и картинками.",
        reply_markup=main_kb()
    )
    await state.clear()

@dp.message(F.text == "📊 Сделать презентацию")
async def start_pres(m: Message, state: FSMContext):
    if not can_generate(m.from_user.id):
        await m.answer("Лимит генераций закончился.")
        return
    await m.answer("Напиши тему презентации:")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_topic)
async def process_topic(m: Message, state: FSMContext):
    text = m.text or ""
    if len(text) > 500:
        await m.answer("Слишком длинно. Максимум 500 символов.")
        return
    await state.update_data(topic=text, extra="", extra_used=0)
    await m.answer("Сколько слайдов?", reply_markup=slides_kb())
    await state.set_state(Form.waiting_slides)

@dp.message(Form.waiting_slides)
async def process_slides(m: Message, state: FSMContext):
    slides = 8
    t = m.text or ""
    if "5" in t:
        slides = 5
    elif "10" in t:
        slides = 10
    await state.update_data(slides=slides)
    await m.answer("Выбери стиль:", reply_markup=style_kb())
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_style)
async def process_style(m: Message, state: FSMContext):
    data = await state.get_data()
    await state.update_data(style=m.text or "Тёмный")
    await m.answer("Делаю пробный вариант...")
    prompt = f"""Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Стиль: {m.text}
Дополнительно: {data.get('extra', '')}

Сделай короткий образец структуры обычным текстом:
Название: ...
1. ...
2. ...
3. ...
Без JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)
    await m.answer(
        f"Пробный вариант:\n\n{sample}\n\nВыбери действие:",
        reply_markup=confirm_kb()
    )
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text == "✏️ Изменить тему")
async def change_topic(m: Message, state: FSMContext):
    await m.answer("Напиши новую тему:")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_confirm, F.text == "➕ Добавить информацию")
async def add_extra(m: Message, state: FSMContext):
    data = await state.get_data()
    if data.get("extra_used", 0) >= 3:
        await m.answer("Лимит добавлений исчерпан.")
        return
    await m.answer("Напиши дополнительную информацию (до 800 символов):")
    await state.set_state(Form.waiting_extra)

@dp.message(Form.waiting_extra)
async def process_extra(m: Message, state: FSMContext):
    data = await state.get_data()
    text = m.text or ""
    if len(text) > 800:
        await m.answer("Слишком длинно. Максимум 800 символов.")
        return
    old = data.get("extra", "")
    new_extra = (old + "\n" + text).strip() if old else text
    extra_used = data.get("extra_used", 0) + 1
    await state.update_data(extra=new_extra, extra_used=extra_used)
    await m.answer("Обновляю пробный вариант...")
    prompt = f"""Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Стиль: {data.get('style')}
Дополнительно: {new_extra}

Сделай короткий образец структуры обычным текстом:
Название: ...
1. ...
2. ...
3. ...
Без JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)
    await m.answer(
        f"Обновлённый пробный вариант:\n\n{sample}\n\nВыбери действие:",
        reply_markup=confirm_kb()
    )
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text.in_(["✅ Делать полную версию", "делай", "да", "ок"]))
async def confirm_generate(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Делаю финальную версию с картинками. Это может занять 1–2 минуты.")

    prompt = f"""Сделай сильную подробную презентацию.
Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Стиль: {data.get('style')}
Дополнительно: {data.get('extra', '')}

На каждом слайде 4–7 предложений полезного текста.
Верни ТОЛЬКО JSON:
{{
  "title": "Название",
  "slides": [
    {{"title": "Заголовок", "content": "Текст", "image_prompt": "English prompt for a realistic photo about this slide"}}
  ]
}}"""
    raw = await ask_grok(prompt)
    try:
        content = json.loads(raw[raw.find("{"):raw.rfind("}") + 1])
    except:
        await m.answer("Ошибка генерации текста. Попробуй ещё раз.")
        await state.clear()
        return

    style = data.get("style", "Тёмный")
    if style == "Светлый":
        bg, tc = (245, 245, 245), (30, 30, 30)
    elif style == "Синий":
        bg, tc = (15, 30, 60), (230, 235, 245)
    else:
        bg, tc = (25, 25, 30), (230, 230, 235)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*bg)
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.8), Inches(11.7), Inches(2))
    p = tb.text_frame.paragraphs[0]
    p.text = content.get("title", "Презентация")
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*tc)
    p.alignment = PP_ALIGN.CENTER

    slides_data = content.get("slides", [])
    max_images = min(5, len(slides_data))

    for idx, s in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg_shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = RGBColor(*bg)

        tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.8))
        p = tb.text_frame.paragraphs[0]
        p.text = s.get("title", "")
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*tc)

        cb = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(6.3), Inches(5.5))
        tf = cb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(*tc)

        img_path = f"/tmp/img_{uid}_{idx}.png"
        img_ok = False
        if idx < max_images:
            img_prompt = s.get("image_prompt") or f"High quality photo about: {s.get('title', data.get('topic'))}"
            img_ok = await generate_image(img_prompt, img_path)

        if img_ok:
            try:
                slide.shapes.add_picture(img_path, Inches(7.2), Inches(1.4), width=Inches(5.5))
            except Exception as e:
                print("Insert image error:", e)
                add_placeholder(slide, 7.2, 1.4, 5.5, 5.2, "Вставь сюда фото")
        else:
            add_placeholder(slide, 7.2, 1.4, 5.5, 5.2, "Вставь сюда фото")

    pptx_path = f"pres_{uid}.pptx"
    prs.save(pptx_path)

    pdf_path = f"pres_{uid}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    font_name, font_bold = "Helvetica", "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
        font_name, font_bold = "DejaVu", "DejaVuBold"
    except:
        pass

    c.setFont(font_bold, 14)
    c.drawString(40, h - 50, content.get("title", "")[:65])
    y = h - 90
    for i, s in enumerate(slides_data, 1):
        if y < 70:
            c.showPage()
            y = h - 50
        c.setFont(font_bold, 11)
        c.drawString(40, y, f"{i}. {s.get('title', '')[:70]}")
        y -= 16
        c.setFont(font_name, 9)
        text = (s.get("content", "") or "")[:200]
        while text:
            c.drawString(40, y, text[:90])
            text = text[90:]
            y -= 13
            if y < 50:
                c.showPage()
                y = h - 50
        y -= 10
    c.save()

    await m.answer_document(FSInputFile(pptx_path), caption="PPTX")
    await m.answer_document(FSInputFile(pdf_path), caption="PDF")
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await m.answer("Готово!", reply_markup=main_kb())
    await state.clear()

@dp.message(F.text == "📁 Моя история")
async def history(m: Message):
    u = get_user(m.from_user.id)
    if not u["history"]:
        await m.answer("История пустая.")
        return
    await m.answer("История:\n\n" + "\n".join(u["history"][-10:]))

@dp.message(F.text == "ℹ️ Мой тариф")
async def my_plan(m: Message):
    u = get_user(m.from_user.id)
    limit = PLAN_LIMITS.get(u["plan"], 15)
    left = max(0, limit - u["generations"])
    await m.answer(f"Генераций: {u['generations']} из {limit}\nОсталось: {left}")

@dp.message(Command("grant"))
async def grant(m: Message):
    if m.from_user.id not in ADMIN_IDS:
        return
    try:
        uid = int(m.text.split()[1])
        get_user(uid)["plan"] = "premium"
        await m.answer(f"Доступ выдан пользователю {uid}")
    except:
        await m.answer("Формат: /grant user_id")

async def main():
    print("Бот запущен")
    print("REPLICATE TOKEN:", "YES" if REPLICATE_API_TOKEN else "NO")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
