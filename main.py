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
ADMIN_IDS = [123456789]  # ← СВОЙ TELEGRAM ID СЮДА

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
async def ask_grok(prompt: str) -> str:
    try:
        response = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный создатель презентаций и документов. Отвечай только на русском, чётко и структурированно."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=2500
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Ошибка Grok: {e}"

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
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="👨‍🎓 Студент")],
            [KeyboardButton(text="💼 Предприниматель")],
            [KeyboardButton(text="🏠 Для себя")]
        ],
        resize_keyboard=True
    )

def main_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📊 Сделать презентацию")],
            [KeyboardButton(text="📄 Сделать документ")],
            [KeyboardButton(text="📋 Готовые шаблоны")],
            [KeyboardButton(text="📁 Моя история"), KeyboardButton(text="ℹ️ Мой тариф")]
        ],
        resize_keyboard=True
    )

def slides_kb():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="5 слайдов"), KeyboardButton(text="8 слайдов")],
            [KeyboardButton(text="10 слайдов"), KeyboardButton(text="15 слайдов")]
        ],
        resize_keyboard=True
    )

def style_kb(plan: str):
    styles = STYLES.get(plan, STYLES["free"])
    buttons = [[KeyboardButton(text=style)] for style in styles]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

# ====================== /START ======================
@dp.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    user = get_user(message.from_user.id)
    name = message.from_user.first_name or "друг"
    user["name"] = name

    await message.answer(
        f"Привет, {name}! 👋\n\n"
        "Я помогаю быстро создавать презентации, документы и готовые шаблоны.\n\n"
        "Выбери, кто ты:",
        reply_markup=category_kb()
    )
    await state.set_state(Form.waiting_category)

@dp.message(Form.waiting_category)
async def process_category(message: Message, state: FSMContext):
    user = get_user(message.from_user.id)
    text = message.text

    if "Студент" in text:
        user["category"] = "student"
        reply = "Отлично! Помогу с презентациями, рефератами и учебными работами."
    elif "Предприниматель" in text:
        user["category"] = "business"
        reply = "Отлично! Помогу с презентациями, коммерческими предложениями и договорами."
    elif "Для себя" in text:
        user["category"] = "personal"
        reply = "Отлично! Помогу с презентациями, договорами и бытовыми документами."
    else:
        await message.answer("Пожалуйста, выбери одну из кнопок.")
        return

    await message.answer(f"{reply}\n\nЧто будем делать?", reply_markup=main_kb())
    await state.clear()

# ====================== МЕНЮ ======================
@dp.message(F.text == "📊 Сделать презентацию")
async def start_presentation(message: Message, state: FSMContext):
    if not can_generate(message.from_user.id):
        await message.answer("Лимит генераций на твоём тарифе закончился.")
        return

    await message.answer("Напиши тему презентации (можно коротко):")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_topic)
async def process_topic(message: Message, state: FSMContext):
    await state.update_data(topic=message.text)
    await message.answer("Сколько слайдов нужно?", reply_markup=slides_kb())
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
    await message.answer("Выбери стиль оформления:", reply_markup=style_kb(user["plan"]))
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_style)
async def process_style(message: Message, state: FSMContext):
    style = message.text
    data = await state.get_data()
    await state.update_data(style=style)

    await message.answer("Делаю пробный вариант... Подожди 15–25 секунд.")

    prompt = f"""
Тема презентации: {data['topic']}
Количество слайдов: {data['slides']}
Стиль: {style}

Сделай краткий образец структуры:
- Название презентации
- Список слайдов (номер + заголовок)

Ответ должен быть коротким и понятным.
"""
    sample = await ask_grok(prompt)

    await state.update_data(sample=sample)
    await message.answer(
        f"Пробный вариант:\n\n{sample}\n\n"
        "Если подходит — напиши «делай»\n"
        "Если нет — просто начни заново через меню."
    )
    await state.set_state(Form.waiting_confirm)

# ====================== ПОДТВЕРЖДЕНИЕ И ГЕНЕРАЦИЯ ======================
@dp.message(Form.waiting_confirm, F.text.lower().in_(["делай", "да", "ок", "хорошо", "подтверждаю"]))
async def confirm_generate(message: Message, state: FSMContext):
    data = await state.get_data()
    user_id = message.from_user.id
    user = get_user(user_id)

    if not data.get("topic"):
        await message.answer("Сначала выбери тему презентации.")
        await state.clear()
        return

    await message.answer("Делаю финальную версию... Это займёт 20–40 секунд.")

    prompt = f"""
Создай полную структуру презентации.

Тема: {data['topic']}
Количество слайдов: {data['slides']}
Стиль: {data['style']}

Верни ответ строго в формате JSON:
{{
    "title": "Название презентации",
    "slides": [
        {{"title": "Заголовок слайда", "content": "Текст слайда"}}
    ]
}}
"""
    raw = await ask_grok(prompt)

    try:
        start = raw.find("{")
        end = raw.rfind("}") + 1
        content = json.loads(raw[start:end])
    except Exception:
        await message.answer("Не удалось разобрать ответ. Попробуй ещё раз.")
        await state.clear()
        return

    bg_color = STYLE_COLORS.get(data.get("style"), (40, 40, 40))
    text_color = (230, 230, 230) if sum(bg_color) < 300 else (30, 30, 30)

    # ===== PPTX =====
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Титульный слайд
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*bg_color)

    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11.3), Inches(1.5))
    p = title_box.text_frame.paragraphs[0]
    p.text = content.get("title", "Презентация")
    p.font.size = Pt(36)
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

        content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(12), Inches(5))
        tf = content_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(*text_color)

    pptx_path = f"pres_{user_id}.pptx"
    prs.save(pptx_path)

    # ===== PDF =====
    pdf_path = f"pres_{user_id}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(40, height - 40, content.get("title", "Презентация")[:70])

    y = height - 80
    for i, s in enumerate(content.get("slides", []), 1):
        if y < 80:
            c.showPage()
            y = height - 40
        c.setFont("Helvetica-Bold", 12)
        c.drawString(40, y, f"{i}. {s.get('title', '')[:80]}")
        y -= 18
        c.setFont("Helvetica", 10)
        text = s.get("content", "")[:200]
        c.drawString(40, y, text[:95])
        y -= 15
        if len(text) > 95:
            c.drawString(40, y, text[95:190])
            y -= 15
        y -= 12

    c.save()

    await message.answer_document(FSInputFile(pptx_path), caption="Редактируемая версия (PPTX)")
    await message.answer_document(FSInputFile(pdf_path), caption="PDF-версия")

    user["generations"] += 1
    user["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title', 'Презентация')}")

    await message.answer("Готово!", reply_markup=main_kb())
    await state.clear()

# ====================== ДОКУМЕНТЫ И ШАБЛОНЫ ======================
@dp.message(F.text == "📄 Сделать документ")
async def make_document(message: Message):
    await message.answer(
        "Пока генерация произвольных документов в разработке.\n"
        "Скоро здесь можно будет создавать коммерческие предложения, отчёты и письма."
    )

@dp.message(F.text == "📋 Готовые шаблоны")
async def templates_menu(message: Message):
    await message.answer(
        "Готовые шаблоны (в разработке):\n\n"
        "• Европротокол (ДТП)\n"
        "• Договор купли-продажи автомобиля\n"
        "• Карточка предприятия\n"
        "• Акт выполненных работ\n\n"
        "Скоро можно будет просто заполнить данные и скачать готовый файл."
    )

@dp.message(F.text == "📁 Моя история")
async def show_history(message: Message):
    user = get_user(message.from_user.id)
    if not user["history"]:
        await message.answer("История пока пустая.")
        return
    text = "Твоя история:\n\n" + "\n".join(user["history"][-10:])
    await message.answer(text)

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
    await message.answer(
        f"Твой тариф: {names.get(user['plan'])}\n"
        f"Использовано генераций: {user['generations']} из {limit}"
    )

# ====================== АДМИНКА ======================
@dp.message(Command("grant"))
async def admin_grant(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    try:
        parts = message.text.split()
        uid = int(parts[1])
        plan = parts[2] if len(parts) > 2 else "premium"
        get_user(uid)["plan"] = plan
        await message.answer(f"Пользователю {uid} выдан тариф: {plan}")
    except:
        await message.answer("Формат: /grant user_id [basic/standard/premium]")

@dp.message(Command("revoke"))
async def admin_revoke(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    try:
        uid = int(message.text.split()[1])
        get_user(uid)["plan"] = "free"
        await message.answer(f"Тариф пользователя {uid} сброшен на free")
    except:
        await message.answer("Формат: /revoke user_id")

@dp.message(Command("users"))
async def admin_users(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    text = "Пользователи:\n\n"
    for uid, data in list(users_db.items())[:25]:
        text += f"{uid} | {data.get('name')} | {data.get('plan')} | gen: {data.get('generations')}\n"
    await message.answer(text or "Пока никого нет")

@dp.message(Command("stats"))
async def admin_stats(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    total = len(users_db)
    gens = sum(u.get("generations", 0) for u in users_db.values())
    await message.answer(f"Всего пользователей: {total}\nВсего генераций: {gens}")

# ====================== ЗАПУСК ======================
async def main():
    print("Бот запущен...")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
