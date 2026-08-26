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
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
ADMIN_IDS = [909828109]

client = AsyncOpenAI(api_key=XAI_API_KEY, base_url="https://api.x.ai/v1")
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
users_db = {}

PLAN_LIMITS = {
    "free": 3,
    "basic": 10,
    "standard": 40,
    "premium": 100
}

STYLES = {
    "free": ["Графит", "Снег"],
    "basic": ["Графит", "Снег", "Океан"],
    "standard": ["Графит", "Снег", "Океан", "Ночь", "Мята", "Коралл"],
    "premium": ["Графит", "Снег", "Океан", "Ночь", "Мята", "Коралл", "Уголь", "Лаванда", "Сканди"]
}

STYLE_COLORS = {
    "Графит": (40, 40, 40),
    "Снег": (245, 245, 245),
    "Океан": (20, 40, 80),
    "Ночь": (15, 15, 15),
    "Мята": (220, 240, 230),
    "Коралл": (255, 230, 220),
    "Уголь": (30, 30, 30),
    "Лаванда": (230, 220, 240),
    "Сканди": (240, 235, 225)
}

class Form(StatesGroup):
    waiting_category = State()
    waiting_topic = State()
    waiting_slides = State()
    waiting_style = State()
    waiting_confirm = State()

def get_user(uid):
    if uid not in users_db:
        users_db[uid] = {
            "name": "",
            "plan": "free",
            "generations": 0,
            "history": []
        }
    return users_db[uid]

def can_generate(uid):
    u = get_user(uid)
    return u["generations"] < PLAN_LIMITS.get(u["plan"], 3)

async def ask_grok(prompt: str) -> str:
    try:
        r = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный создатель презентаций. Отвечай только на русском."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=3500
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"Ошибка: {e}"

def get_prompt(plan: str, topic: str, slides: int, style: str) -> str:
    if plan == "premium":
        level = (
            "Сделай максимально сильную и подробную презентацию. "
            "На каждом слайде 5–8 предложений насыщенного полезного текста. "
            "Заголовки сильные, структура профессиональная."
        )
    elif plan == "standard":
        level = (
            "Сделай качественную полноценную презентацию. "
            "На каждом слайде 3–5 предложений хорошего текста. "
            "Заголовки понятные и сильные."
        )
    else:
        level = (
            "Сделай нормальную аккуратную презентацию. "
            "На каждом слайде 2–3 предложения краткого полезного текста."
        )

    return f"""{level}

Тема: {topic}
Количество слайдов: {slides}
Стиль: {style}

Верни ТОЛЬКО валидный JSON:
{{
  "title": "Название презентации",
  "slides": [
    {{"title": "Заголовок слайда", "content": "Текст слайда"}}
  ]
}}"""

def category_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="👨‍🎓 Студент")],
        [KeyboardButton(text="💼 Предприниматель")],
        [KeyboardButton(text="🏠 Для себя")]
    ], resize_keyboard=True)

def main_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="📊 Сделать презентацию")],
        [KeyboardButton(text="📁 Моя история"), KeyboardButton(text="ℹ️ Мой тариф")]
    ], resize_keyboard=True)

def slides_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="5 слайдов"), KeyboardButton(text="8 слайдов")],
        [KeyboardButton(text="10 слайдов"), KeyboardButton(text="12 слайдов")]
    ], resize_keyboard=True)

def style_kb(plan):
    styles = STYLES.get(plan, STYLES["free"])
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=s)] for s in styles],
        resize_keyboard=True
    )

@dp.message(Command("start"))
async def cmd_start(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    u["name"] = m.from_user.first_name or "друг"
    await m.answer(
        f"Привет, {u['name']}!\n\nЯ делаю презентации.\nВыбери категорию:",
        reply_markup=category_kb()
    )
    await state.set_state(Form.waiting_category)

@dp.message(Form.waiting_category)
async def process_category(m: Message, state: FSMContext):
    await m.answer("Что будем делать?", reply_markup=main_kb())
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
    await state.update_data(topic=m.text)
    await m.answer("Сколько слайдов?", reply_markup=slides_kb())
    await state.set_state(Form.waiting_slides)

@dp.message(Form.waiting_slides)
async def process_slides(m: Message, state: FSMContext):
    slides = 8
    text = m.text or ""
    if "5" in text:
        slides = 5
    elif "10" in text:
        slides = 10
    elif "12" in text:
        slides = 12
    await state.update_data(slides=slides)
    u = get_user(m.from_user.id)
    await m.answer("Выбери стиль:", reply_markup=style_kb(u["plan"]))
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_style)
async def process_style(m: Message, state: FSMContext):
    data = await state.get_data()
    await state.update_data(style=m.text)
    await m.answer("Делаю пробный вариант...")

    prompt = f"""Тема: {data['topic']}
Слайдов: {data['slides']}
Стиль: {m.text}

Сделай короткий образец структуры презентации обычным текстом:
Название: ...
1. ...
2. ...
3. ...
Не используй JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)
    await m.answer(f"Пробный вариант:\n\n{sample}\n\nЕсли нравится — напиши «делай»")
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text.lower().in_(["делай", "да", "ок", "хорошо", "подтверждаю"]))
async def confirm_generate(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Делаю финальную версию...")

    prompt = get_prompt(u["plan"], data["topic"], data["slides"], data["style"])
    raw = await ask_grok(prompt)

    try:
        content = json.loads(raw[raw.find("{"):raw.rfind("}") + 1])
    except:
        await m.answer("Ошибка генерации. Попробуй ещё раз.")
        await state.clear()
        return

    bg = STYLE_COLORS.get(data.get("style"), (40, 40, 40))
    tc = (230, 230, 230) if sum(bg) < 300 else (30, 30, 30)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Титульный
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*bg)
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(11.7), Inches(2))
    p = tb.text_frame.paragraphs[0]
    p.text = content.get("title", "Презентация")
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*tc)
    p.alignment = PP_ALIGN.CENTER

    for s in content.get("slides", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg_shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = RGBColor(*bg)

        tb = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(12), Inches(1))
        p = tb.text_frame.paragraphs[0]
        p.text = s.get("title", "")
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*tc)

        cb = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(12), Inches(5.2))
        tf = cb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(*tc)

    pptx_path = f"pres_{uid}.pptx"
    prs.save(pptx_path)

    # PDF
    pdf_path = f"pres_{uid}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, h - 50, content.get("title", "")[:60])
    y = h - 90
    for i, s in enumerate(content.get("slides", []), 1):
        if y < 70:
            c.showPage()
            y = h - 50
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y, f"{i}. {s.get('title', '')[:70]}")
        y -= 16
        c.setFont("Helvetica", 9)
        c.drawString(40, y, (s.get("content", "") or "")[:90])
        y -= 20
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
    names = {
        "free": "Бесплатный",
        "basic": "Базовый (199 ₽)",
        "standard": "Стандарт (499 ₽)",
        "premium": "Премиум (999 ₽)"
    }
    limit = PLAN_LIMITS.get(u["plan"], 3)
    left = max(0, limit - u["generations"])
    await m.answer(
        f"Тариф: {names.get(u['plan'])}\n"
        f"Использовано: {u['generations']} из {limit}\n"
        f"Осталось: {left}"
    )

@dp.message(Command("grant"))
async def grant(m: Message):
    if m.from_user.id not in ADMIN_IDS:
        return
    try:
        parts = m.text.split()
        uid = int(parts[1])
        plan = parts[2] if len(parts) > 2 else "premium"
        get_user(uid)["plan"] = plan
        await m.answer(f"Выдан тариф {plan} пользователю {uid}")
    except:
        await m.answer("Формат: /grant user_id [basic/standard/premium]")

@dp.message(Command("revoke"))
async def revoke(m: Message):
    if m.from_user.id not in ADMIN_IDS:
        return
    try:
        uid = int(m.text.split()[1])
        get_user(uid)["plan"] = "free"
        await m.answer(f"Сброшен тариф у {uid}")
    except:
        await m.answer("Формат: /revoke user_id")

@dp.message(Command("users"))
async def users_list(m: Message):
    if m.from_user.id not in ADMIN_IDS:
        return
    text = "Пользователи:\n"
    for uid, d in list(users_db.items())[:20]:
        text += f"{uid} | {d.get('name')} | {d.get('plan')} | {d.get('generations')}\n"
    await m.answer(text or "Пусто")

async def main():
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
