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

# ====================== НАСТРОЙКИ ======================
BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
ADMIN_IDS = [909828109]  # ← СВОЙ TELEGRAM ID

client = AsyncOpenAI(api_key=XAI_API_KEY, base_url="https://api.x.ai/v1")
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

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
    "standard": ["Графит", "Снег", "Океан", "Ночь", "Мята", "Коралл", "Уголь"],
    "premium": ["Графит", "Снег", "Океан", "Ночь", "Мята", "Коралл", "Уголь",
                "Эспрессо", "Лаванда", "Неон", "Сканди", "Изумруд"]
}

STYLE_COLORS = {
    "Графит": (40, 40, 40),
    "Снег": (245, 245, 245),
    "Океан": (20, 40, 80),
    "Ночь": (15, 15, 15),
    "Мята": (220, 240, 230),
    "Коралл": (255, 230, 220),
    "Уголь": (30, 30, 30),
    "Эспрессо": (50, 30, 20),
    "Лаванда": (230, 220, 240),
    "Неон": (10, 10, 20),
    "Сканди": (240, 235, 225),
    "Изумруд": (20, 50, 40)
}

# ====================== СОСТОЯНИЯ ======================
class Form(StatesGroup):
    waiting_category = State()
    waiting_topic = State()
    waiting_slides = State()
    waiting_style = State()
    waiting_confirm = State()

# ====================== ФУНКЦИИ ======================
def get_quality_prompt(plan: str, topic: str, slides: int, style: str) -> str:
    if plan == "premium":
        level = "Ты топовый презентационный дизайнер. Сделай презентацию профессионального уровня: сильные заголовки, ёмкий текст, чёткая логика."
    elif plan == "standard":
        level = "Сделай качественную презентацию хорошего уровня: понятные заголовки, подробный текст, логичная структура."
    else:
        level = "Сделай нормальную аккуратную презентацию: понятные заголовки, краткий полезный текст, простая структура."

    prompt = level + "\n\n"
    prompt += f"Тема: {topic}\n"
    prompt += f"Количество слайдов: {slides}\n"
    prompt += f"Стиль: {style}\n\n"
    prompt += "Верни ответ строго в JSON:\n"
    prompt += "{\n"
    prompt += '    "title": "Название презентации",\n'
    prompt += '    "slides": [\n'
    prompt += '        {"title": "Заголовок слайда", "content": "Текст слайда"}\n'
    prompt += "    ]\n"
    prompt += "}"
    return prompt

async def ask_grok(prompt: str) -> str:
    try:
        response = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный создатель презентаций. Отвечай только на русском. Всегда возвращай валидный JSON, когда тебя просят."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=3000
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Ошибка: {e}"

def get_user(user_id: int):
    if user_id not in users_db:
        users_db[user_id] = {
            "name": "",
            "category": None,
            "plan": "free",
            "generations": 0,
            "history": []
        }
    return users_db[user_id]

def can_generate(user_id: int) -> bool:
    user = get_user(user_id)
    return user["generations"] < PLAN_LIMITS.get(user["plan"], 3)

# ====================== КЛАВИАТУРЫ ======================
def category_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="👨‍🎓 Студент")],
        [KeyboardButton(text="💼 Предприниматель")],
        [KeyboardButton(text="🏠 Для себя")]
    ], resize_keyboard=True)

def main_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="📊 Сделать презентацию")],
        [KeyboardButton(text="📄 Сделать документ")],
        [KeyboardButton(text="📋 Готовые шаблоны")],
        [KeyboardButton(text="📁 Моя история"), KeyboardButton(text="ℹ️ Мой тариф")]
    ], resize_keyboard=True)

def slides_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="5 слайдов"), KeyboardButton(text="8 слайдов")],
        [KeyboardButton(text="10 слайдов"), KeyboardButton(text="15 слайдов")]
    ], resize_keyboard=True)

def style_kb(plan: str):
    styles = STYLES.get(plan, STYLES["free"])
    buttons = [[KeyboardButton(text=s)] for s in styles]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

# ====================== /START ======================
@dp.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    user = get_user(message.from_user.id)
    name = message.from_user.first_name or "друг"
    user["name"] = name

    await message.answer(
        f"Привет, {name}! 👋\n\n"
        "Я помогаю создавать презентации и документы.\n"
        "Выбери категорию:",
        reply_markup=category_kb()
    )
    await state.set_state(Form.waiting_category)

@dp.message(Form.waiting_category)
async def process_category(message: Message, state: FSMContext):
    user = get_user(message.from_user.id)
    text = message.text

    if "Студент" in text:
        user["category"] = "student"
        reply = "Отлично! Помогу с учебными презентациями."
    elif "Предприниматель" in text:
        user["category"] = "business"
        reply = "Отлично! Помогу с бизнес-презентациями."
    elif "Для себя" in text:
        user["category"] = "personal"
        reply = "Отлично! Помогу с презентациями и документами."
    else:
        await message.answer("Выбери кнопку.")
        return

    await message.answer(f"{reply}\n\nЧто будем делать?", reply_markup=main_kb())
    await state.clear()

# ====================== ПРЕЗЕНТАЦИЯ ======================
@dp.message(F.text == "📊 Сделать презентацию")
async def start_presentation(message: Message, state: FSMContext):
    if not can_generate(message.from_user.id):
        await message.answer("Лимит генераций закончился.")
        return
    await message.answer("Напиши тему презентации:")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_topic)
async def process_topic(message: Message, state: FSMContext):
    await state.update_data(topic=message.text)
    await message.answer("Сколько слайдов?", reply_markup=slides_kb())
    await state.set_state(Form.waiting_slides)

@dp.message(Form.waiting_slides)
async def process_slides(message: Message, state: FSMContext):
    slides = 8
    if "5" in message.text:
        slides = 5
    elif "10" in message.text:
        slides = 10
    elif "15" in message.text:
        slides = 15

    await state.update_data(slides=slides)
    user = get_user(message.from_user.id)
    await message.answer("Выбери стиль:", reply_markup=style_kb(user["plan"]))
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_style)
async def process_style(message: Message, state: FSMContext):
    style = message.text
    data = await state.get_data()
    await state.update_data(style=style)

    await message.answer("Делаю пробный вариант...")

    prompt = f"Тема: {data['topic']}\nСлайдов: {data['slides']}\nСтиль: {style}\n\nСделай короткий образец структуры презентации (название + список слайдов)."
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)

    await message.answer(f"Пробный вариант:\n\n{sample}\n\nЕсли нравится — напиши «делай»")
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text.lower().in_(["делай", "да", "ок", "хорошо", "подтверждаю"]))
async def confirm_generate(message: Message, state: FSMContext):
    data = await state.get_data()
    user_id = message.from_user.id
    user = get_user(user_id)

    await message.answer("Делаю финальную версию...")

    prompt = get_quality_prompt(
        plan=user["plan"],
        topic=data["topic"],
        slides=data["slides"],
        style=data["style"]
    )

    raw = await ask_grok(prompt)

    try:
        start = raw.find("{")
        end = raw.rfind("}") + 1
        content = json.loads(raw[start:end])
    except:
        await message.answer("Ошибка при генерации. Попробуй ещё раз.")
        await state.clear()
        return

    bg_color = STYLE_COLORS.get(data.get("style"), (40, 40, 40))
    text_color = (230, 230, 230) if sum(bg_color) < 300 else (30, 30, 30)

    # PPTX
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*bg_color)

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(2.6), Inches(11.7), Inches(2))
    p = title_box.text_frame.paragraphs[0]
    p.text = content.get("title", "Презентация")
    p.font.size = Pt(34)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*text_color)
    p.alignment = PP_ALIGN.CENTER

    for s in content.get("slides", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(*bg_color)

        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(12), Inches(1))
        p = title_box.text_frame.paragraphs[0]
        p.text = s.get("title", "")
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*text_color)

        content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(12), Inches(5.2))
        tf = content_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(*text_color)

    pptx_path = f"pres_{user_id}.pptx"
    prs.save(pptx_path)

    # PDF
    pdf_path = f"pres_{user_id}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height - 50, content.get("title", "Presentation")[:65])

    y = height - 90
    for i, s in enumerate(content.get("slides", []), 1):
        if y < 80:
            c.showPage()
            y = height - 50
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, f"{i}. {s.get('title', '')[:70]}")
        y -= 18
        c.setFont("Helvetica", 10)
        text = s.get("content", "")[:220]
        c.drawString(40, y, text[:90])
        y -= 14
        if len(text) > 90:
            c.drawString(40, y, text[90:180])
            y -= 14
        y -= 16

    c.save()

    await message.answer_document(FSInputFile(pptx_path), caption="Презентация (PPTX)")
    await message.answer_document(FSInputFile(pdf_path), caption="PDF версия")

    user["generations"] += 1
    user["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await message.answer("Готово!", reply_markup=main_kb())
    await state.clear()

# ====================== ОСТАЛЬНОЕ ======================
@dp.message(F.text == "📄 Сделать документ")
async def make_doc(message: Message):
    await message.answer("Генерация документов скоро будет доступна.")

@dp.message(F.text == "📋 Готовые шаблоны")
async def templates(message: Message):
    await message.answer("Шаблоны скоро появятся.")

@dp.message(F.text == "📁 Моя история")
async def history(message: Message):
    user = get_user(message.from_user.id)
    if not user["history"]:
        await message.answer("История пустая.")
        return
    await message.answer("История:\n\n" + "\n".join(user["history"][-10:]))

@dp.message(F.text == "ℹ️ Мой тариф")
async def my_plan(message: Message):
    user = get_user(message.from_user.id)
    names = {
        "free": "Бесплатный",
        "basic": "Базовый (199 ₽)",
        "standard": "Стандарт (499 ₽)",
        "premium": "Премиум (999 ₽)"
    }
    limit = PLAN_LIMITS.get(user["plan"], 3)
    await message.answer(f"Тариф: {names.get(user['plan'])}\nГенераций: {user['generations']} из {limit}")

# ====================== АДМИНКА ======================
@dp.message(Command("grant"))
async def grant(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    try:
        parts = message.text.split()
        uid = int(parts[1])
        plan = parts[2] if len(parts) > 2 else "premium"
        get_user(uid)["plan"] = plan
        await message.answer(f"Выдан тариф {plan} пользователю {uid}")
    except:
        await message.answer("Формат: /grant user_id [basic/standard/premium]")

@dp.message(Command("revoke"))
async def revoke(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    try:
        uid = int(message.text.split()[1])
        get_user(uid)["plan"] = "free"
        await message.answer(f"Сброшен тариф у {uid}")
    except:
        await message.answer("Формат: /revoke user_id")

@dp.message(Command("users"))
async def users_list(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    text = "Пользователи:\n"
    for uid, d in list(users_db.items())[:20]:
        text += f"{uid} | {d.get('name')} | {d.get('plan')} | {d.get('generations')}\n"
    await message.answer(text or "Пусто")

# ====================== ЗАПУСК ======================
async def main():
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
