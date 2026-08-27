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
from PIL import Image
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
    waiting_confirm = State()
    waiting_extra = State()

THEMES = {
    "nature": {
        "bg": (248, 246, 241), "ink": (22, 22, 24), "mid": (70, 70, 74), "mute": (130, 128, 124), "line": (22, 22, 24),
        "photo": "photorealistic nature photography, cinematic sunlight, no text, no watermark"
    },
    "business": {
        "bg": (16, 16, 18), "ink": (245, 245, 247), "mid": (196, 196, 200), "mute": (120, 120, 126), "line": (212, 175, 90),
        "photo": "premium business photography, architecture, office, cinematic, no text, no watermark"
    },
    "tech": {
        "bg": (8, 16, 28), "ink": (240, 246, 255), "mid": (176, 196, 220), "mute": (110, 130, 155), "line": (90, 170, 230),
        "photo": "futuristic technology photography, neon ambient light, cinematic, no text, no watermark"
    },
    "school": {
        "bg": (250, 249, 246), "ink": (28, 32, 40), "mid": (60, 64, 72), "mute": (120, 124, 132), "line": (40, 90, 180),
        "photo": "clear educational photo, bright, simple subject, no text, no watermark"
    },
    "fashion": {
        "bg": (252, 250, 247), "ink": (18, 18, 18), "mid": (70, 66, 62), "mute": (140, 134, 128), "line": (18, 18, 18),
        "photo": "editorial fashion photography, magazine look, cinematic, no text, no watermark"
    },
    "default": {
        "bg": (248, 246, 241), "ink": (22, 22, 24), "mid": (70, 70, 74), "mute": (130, 128, 124), "line": (22, 22, 24),
        "photo": "cinematic photorealistic photo, no text, no watermark"
    }
}

def pick_theme(topic: str):
    t = (topic or "").lower()
    if any(x in t for x in ["животн", "природ", "океан", "кит", "лес", "эколог", "моря", "цвет"]):
        return "nature", THEMES["nature"]
    if any(x in t for x in ["бизнес", "компани", "продаж", "финанс", "инвест", "стартап", "рынок"]):
        return "business", THEMES["business"]
    if any(x in t for x in ["крипт", "блокчейн", "нейро", "техно", "ai", "ии", "код", "робот", "софт"]):
        return "tech", THEMES["tech"]
    if any(x in t for x in ["школ", "универ", "урок", "студент", "доклад", "история", "биолог"]):
        return "school", THEMES["school"]
    if any(x in t for x in ["мод", "стиль", "бренд", "дизайн", "искусств", "фото"]):
        return "fashion", THEMES["fashion"]
    return "default", THEMES["default"]

def get_user(uid):
    if uid not in users_db:
        users_db[uid] = {"name": "", "plan": "premium", "generations": 0, "history": []}
    return users_db[uid]

def can_generate(uid):
    u = get_user(uid)
    return u["generations"] < PLAN_LIMITS.get(u["plan"], 15)

async def ask_grok(prompt: str) -> str:
    try:
        r = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты арт-директор презентаций. Пиши коротко, сильно, по-русски."},
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
        headers = {"Authorization": f"Bearer {REPLICATE_API_TOKEN}", "Content-Type": "application/json"}
        async with httpx.AsyncClient(timeout=120) as http:
            create = await http.post(
                "https://api.replicate.com/v1/models/black-forest-labs/flux-schnell/predictions",
                headers=headers,
                json={"input": {"prompt": prompt, "aspect_ratio": "16:9", "output_format": "png"}}
            )
            print("Replicate create:", create.status_code)
            create.raise_for_status()
            get_url = create.json()["urls"]["get"]
            for _ in range(30):
                await asyncio.sleep(2)
                info = (await http.get(get_url, headers=headers)).json()
                if info.get("status") == "succeeded":
                    out = info.get("output")
                    url = out[0] if isinstance(out, list) else out
                    raw = await http.get(str(url))
                    with open(path, "wb") as f:
                        f.write(raw.content)
                    Image.open(path).convert("RGB").save(path, "PNG")
                    return True
                if info.get("status") in ("failed", "canceled"):
                    return False
        return False
    except Exception as e:
        print("Image generation error:", e)
        return False

def cover(src, dest, w, h):
    im = Image.open(src).convert("RGB")
    t = w / h
    iw, ih = im.size
    if iw / ih > t:
        nw = int(ih * t)
        x = (iw - nw) // 2
        im = im.crop((x, 0, x + nw, ih))
    else:
        nh = int(iw / t)
        y = (ih - nh) // 2
        im = im.crop((0, y, iw, y + nh))
    im.resize((w, h), Image.LANCZOS).save(dest, "PNG")

def rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = RGBColor(*color)
    s.line.fill.background()

def txt(slide, l, t, w, h, text, size, color, bold=False):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text or ""
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = RGBColor(*color)
    p.font.name = "Calibri"
    return box

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
    await m.answer(f"Привет, {u['name']}!\n\nЯ делаю презентации под твою тему.", reply_markup=main_kb())
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
    name, _ = pick_theme(text)
    await state.update_data(topic=text, extra="", extra_used=0, theme_name=name)
    await m.answer("Сколько слайдов?", reply_markup=slides_kb())
    await state.set_state(Form.waiting_slides)

@dp.message(Form.waiting_slides)
async def process_slides(m: Message, state: FSMContext):
    slides = 8
    t = m.text or ""
    if "5" in t: slides = 5
    elif "10" in t: slides = 10
    await state.update_data(slides=slides)
    data = await state.get_data()
    await m.answer("Делаю пробный вариант...")
    sample = await ask_grok(f"Тема: {data.get('topic')}\nСлайдов: {slides}\nДоп: {data.get('extra')}\nКороткий план: название и 3 пункта. Без JSON.")
    await state.update_data(sample=sample)
    await m.answer(f"Пробный вариант:\n\n{sample}\n\nВыбери действие:", reply_markup=confirm_kb())
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text == "✏️ Изменить тему")
async def change_topic(m: Message, state: FSMContext):
    await m.answer("Напиши новую тему:")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_confirm, F.text == "➕ Добавить информацию")
async def add_extra(m: Message, state: FSMContext):
    if (await state.get_data()).get("extra_used", 0) >= 3:
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
    extra = ((data.get("extra") or "") + "\n" + text).strip()
    await state.update_data(extra=extra, extra_used=data.get("extra_used", 0) + 1)
    await m.answer("Обновляю пробный вариант...")
    sample = await ask_grok(f"Тема: {data.get('topic')}\nДоп: {extra}\nКороткий план, 3 пункта.")
    await state.update_data(sample=sample)
    await m.answer(f"Обновлённый пробный вариант:\n\n{sample}\n\nВыбери действие:", reply_markup=confirm_kb())
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text.in_(["✅ Делать полную версию", "делай", "да", "ок"]))
async def confirm_generate(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Собираю презентацию. Это займёт 1–2 минуты.")

    theme_name, colors = pick_theme(data.get("topic", ""))
    raw = await ask_grok(f"""Собери презентацию как дорогой журнал.
Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Доп: {data.get('extra')}
Стиль оформления: {theme_name}
Заголовок слайда 3–6 слов.
Текст: 2 коротких абзаца.
Только JSON:
{{"title":"...","slides":[{{"title":"...","content":"абзац1\\n\\nабзац2","image_prompt":"..."}}]}}""")
    try:
        content = json.loads(raw[raw.find("{"):raw.rfind("}") + 1])
    except:
        await m.answer("Ошибка генерации текста. Попробуй ещё раз.")
        await state.clear()
        return

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slides_data = content.get("slides", [])
    n = len(slides_data)
    images = []
    for i, s in enumerate(slides_data):
        if i >= 5:
            images.append(None)
            continue
        src = f"/tmp/{uid}_{i}.png"
        prompt = f"{s.get('image_prompt') or s.get('title')}, {colors['photo']}"
        ok = await generate_image(prompt, src)
        if ok:
            wide = f"/tmp/{uid}_{i}_w.png"
            tall = f"/tmp/{uid}_{i}_t.png"
            cover(src, wide, 1920, 1080)
            cover(src, tall, 1260, 1500)
            images.append((wide, tall))
        else:
            images.append(None)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, colors["bg"])
    if images and images[0]:
        slide.shapes.add_picture(images[0][0], Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
        rect(slide, 0, 4.7, 13.333, 2.8, colors["bg"])
    txt(slide, 0.7, 5.0, 12, 1.5, content.get("title", "Презентация"), 40, colors["ink"], True)
    txt(slide, 0.7, 6.6, 12, 0.4, "01  /  введение", 13, colors["mute"])

    for idx, s in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        rect(slide, 0, 0, 13.333, 7.5, colors["bg"])
        layout = idx % 3
        img = images[idx] if idx < len(images) else None

        if layout == 0:
            if img:
                slide.shapes.add_picture(img[1], Inches(0), Inches(0), width=Inches(6.4), height=Inches(7.5))
            txt(slide, 7.05, 1.5, 5.5, 1.6, s.get("title", ""), 30, colors["ink"], True)
            rect(slide, 7.05, 3.25, 0.85, 0.05, colors["line"])
            box = slide.shapes.add_textbox(Inches(7.05), Inches(3.5), Inches(5.5), Inches(3.2))
            tf = box.text_frame
            tf.word_wrap = True
            blocks = [x.strip() for x in (s.get("content") or "").split("\n") if x.strip()][:2]
            for i, b in enumerate(blocks or [""]):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = b
                p.font.size = Pt(16)
                p.font.color.rgb = RGBColor(*colors["mid"])
                p.font.name = "Calibri"
                p.space_after = Pt(16)
            txt(slide, 7.05, 6.95, 5.5, 0.3, f"{idx+2:02}  /  {n+1:02}", 12, colors["mute"])
        elif layout == 1:
            if img:
                slide.shapes.add_picture(img[0], Inches(0), Inches(0), width=Inches(13.333), height=Inches(4.55))
            txt(slide, 0.7, 4.85, 12, 1.0, s.get("title", ""), 28, colors["ink"], True)
            box = slide.shapes.add_textbox(Inches(0.7), Inches(5.85), Inches(12), Inches(1.2))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = " ".join((s.get("content") or "").split())
            p.font.size = Pt(15)
            p.font.color.rgb = RGBColor(*colors["mid"])
            p.font.name = "Calibri"
        else:
            txt(slide, 0.7, 1.3, 8.2, 2.2, s.get("title", ""), 36, colors["ink"], True)
            rect(slide, 0.7, 3.6, 1.1, 0.06, colors["line"])
            box = slide.shapes.add_textbox(Inches(0.7), Inches(3.9), Inches(7.4), Inches(2.6))
            tf = box.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = " ".join((s.get("content") or "").split())
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(*colors["mid"])
            p.font.name = "Calibri"
            if img:
                slide.shapes.add_picture(img[1], Inches(8.7), Inches(1.3), width=Inches(3.9), height=Inches(4.7))
            txt(slide, 0.7, 6.95, 5.5, 0.3, f"{idx+2:02}  /  {n+1:02}", 12, colors["mute"])

    pptx_path = f"pres_{uid}.pptx"
    prs.save(pptx_path)

    pdf_path = f"pres_{uid}.pdf"
    pdf = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    fn, fb = "Helvetica", "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
        fn, fb = "DejaVu", "DejaVuBold"
    except:
        pass
    pdf.setFont(fb, 14)
    pdf.drawString(40, h - 50, content.get("title", "")[:65])
    y = h - 90
    for i, s in enumerate(slides_data, 1):
        if y < 70:
            pdf.showPage()
            y = h - 50
        pdf.setFont(fb, 11)
        pdf.drawString(40, y, f"{i}. {s.get('title', '')[:70]}")
        y -= 20
    pdf.save()

    await m.answer_document(FSInputFile(pptx_path), caption="PPTX")
    await m.answer_document(FSInputFile(pdf_path), caption="PDF")
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await m.answer("Готово!", reply_markup=main_kb())
    await state.clear()

@dp.message(F.text == "📁 Моя история")
async def history(m: Message):
    u = get_user(m.from_user.id)
    await m.answer("История пустая." if not u["history"] else "История:\n\n" + "\n".join(u["history"][-10:]))

@dp.message(F.text == "ℹ️ Мой тариф")
async def my_plan(m: Message):
    u = get_user(m.from_user.id)
    limit = PLAN_LIMITS.get(u["plan"], 15)
    await m.answer(f"Генераций: {u['generations']} из {limit}\nОсталось: {max(0, limit - u['generations'])}")

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
