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
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
ADMIN_IDS = [909828109]

client = AsyncOpenAI(api_key=XAI_API_KEY, base_url="https://api.x.ai/v1")
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

users_db = {}

PLAN_LIMITS = {"free": 3, "basic": 10, "standard": 40, "premium": 100}

STYLES = {
    "free": ["Графит", "Снег"],
    "basic": ["Графит", "Снег", "Океан"],
    "standard": ["Графит", "Снег", "Океан", "Ночь", "Мята", "Коралл", "Уголь"],
    "premium": ["Графит", "Снег", "Океан", "Ночь", "Мята", "Коралл", "Уголь",
                "Эспрессо", "Лаванда", "Неон", "Сканди", "Изумруд"]
}

STYLE_COLORS = {
    "Графит": (40, 40, 40), "Снег": (245, 245, 245), "Океан": (20, 40, 80),
    "Ночь": (15, 15, 15), "Мята": (220, 240, 230), "Коралл": (255, 230, 220),
    "Уголь": (30, 30, 30), "Эспрессо": (50, 30, 20), "Лаванда": (230, 220, 240),
    "Неон": (10, 10, 20), "Сканди": (240, 235, 225), "Изумруд": (20, 50, 40)
}

class Form(StatesGroup):
    waiting_category = State()
    waiting_topic = State()
    waiting_slides = State()
    waiting_style = State()
    waiting_confirm = State()
    waiting_doc_topic = State()
    waiting_template_data = State()

def get_user(user_id: int):
    if user_id not in users_db:
        users_db[user_id] = {"name": "", "category": None, "plan": "free", "generations": 0, "history": []}
    return users_db[user_id]

def can_generate(user_id: int) -> bool:
    user = get_user(user_id)
    return user["generations"] < PLAN_LIMITS.get(user["plan"], 3)

async def ask_grok(prompt: str) -> str:
    try:
        response = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный помощник. Отвечай только на русском."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=3000
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Ошибка: {e}"

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
    return ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text=s)] for s in styles], resize_keyboard=True)

def templates_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="Европротокол")],
        [KeyboardButton(text="Договор купли-продажи авто")],
        [KeyboardButton(text="Карточка предприятия")],
        [KeyboardButton(text="Акт выполненных работ")],
        [KeyboardButton(text="⬅️ Назад")]
    ], resize_keyboard=True)

@dp.message(Command("start"))
async def cmd_start(message: Message, state: FSMContext):
    user = get_user(message.from_user.id)
    name = message.from_user.first_name or "друг"
    user["name"] = name
    await message.answer(f"Привет, {name}! 👋\n\nВыбери категорию:", reply_markup=category_kb())
    await state.set_state(Form.waiting_category)

@dp.message(Form.waiting_category)
async def process_category(message: Message, state: FSMContext):
    user = get_user(message.from_user.id)
    text = message.text
    if "Студент" in text:
        user["category"] = "student"
    elif "Предприниматель" in text:
        user["category"] = "business"
    elif "Для себя" in text:
        user["category"] = "personal"
    else:
        await message.answer("Выбери кнопку.")
        return
    await message.answer("Что будем делать?", reply_markup=main_kb())
    await state.clear()

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
    if "5" in message.text: slides = 5
    elif "10" in message.text: slides = 10
    elif "15" in message.text: slides = 15
    await state.update_data(slides=slides)
    user = get_user(message.from_user.id)
    await message.answer("Выбери стиль:", reply_markup=style_kb(user["plan"]))
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_style)
async def process_style(message: Message, state: FSMContext):
    data = await state.get_data()
    await state.update_data(style=message.text)
    await message.answer("Делаю пробный вариант...")

    prompt = f"""Тема: {data['topic']}
Слайдов: {data['slides']}
Стиль: {message.text}

Сделай короткий красивый образец структуры презентации обычным текстом (не JSON):
Название: ...
1. ...
2. ...
"""
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

    prompt = f"""Создай презентацию.
Тема: {data['topic']}
Слайдов: {data['slides']}
Стиль: {data['style']}

Верни строго JSON:
{{
    "title": "Название",
    "slides": [{{"title": "Заголовок", "content": "Текст"}}]
}}"""
    raw = await ask_grok(prompt)
    try:
        content = json.loads(raw[raw.find("{"):raw.rfind("}")+1])
    except:
        await message.answer("Ошибка генерации.")
        await state.clear()
        return

    bg = STYLE_COLORS.get(data.get("style"), (40, 40, 40))
    tc = (230, 230, 230) if sum(bg) < 300 else (30, 30, 30)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

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
        cb = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(12), Inches(5))
        p = cb.text_frame.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(*tc)

    pptx_path = f"pres_{user_id}.pptx"
    prs.save(pptx_path)

    pdf_path = f"pres_{user_id}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, h-50, content.get("title", "")[:60])
    y = h - 90
    for i, s in enumerate(content.get("slides", []), 1):
        if y < 70:
            c.showPage()
            y = h - 50
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y, f"{i}. {s.get('title','')[:70]}")
        y -= 16
        c.setFont("Helvetica", 9)
        c.drawString(40, y, (s.get("content","") or "")[:90])
        y -= 20
    c.save()

    await message.answer_document(FSInputFile(pptx_path), caption="PPTX")
    await message.answer_document(FSInputFile(pdf_path), caption="PDF")
    user["generations"] += 1
    user["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await message.answer("Готово!", reply_markup=main_kb())
    await state.clear()

@dp.message(F.text == "📄 Сделать документ")
async def start_document(message: Message, state: FSMContext):
    if not can_generate(message.from_user.id):
        await message.answer("Лимит генераций закончился.")
        return
    await message.answer("Напиши, какой документ нужен:")
    await state.set_state(Form.waiting_doc_topic)

@dp.message(Form.waiting_doc_topic)
async def process_document(message: Message, state: FSMContext):
    user_id = message.from_user.id
    user = get_user(user_id)
    await message.answer("Делаю документ...")

    prompt = f"Создай полноценный документ на тему: {message.text}\nСделай структурированный текст."
    text = await ask_grok(prompt)

    doc = Document()
    doc.add_heading(message.text[:80], 0)
    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line.strip())

    path = f"doc_{user_id}.docx"
    doc.save(path)

    await message.answer_document(FSInputFile(path), caption="Документ (Word)")
    user["generations"] += 1
    user["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — Документ")
    await message.answer("Готово!", reply_markup=main_kb())
    await state.clear()

@dp.message(F.text == "📋 Готовые шаблоны")
async def templates_menu(message: Message):
    await message.answer("Выбери шаблон:", reply_markup=templates_kb())

@dp.message(F.text == "⬅️ Назад")
async def back_to_main(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Главное меню:", reply_markup=main_kb())

@dp.message(F.text.in_(["Европротокол", "Договор купли-продажи авто", "Карточка предприятия", "Акт выполненных работ"]))
async def choose_template(message: Message, state: FSMContext):
    await state.update_data(template=message.text)
    await message.answer(f"Выбран: {message.text}\n\nНапиши данные (ФИО, даты, суммы и т.д.):")
    await state.set_state(Form.waiting_template_data)

@dp.message(Form.waiting_template_data)
async def process_template(message: Message, state: FSMContext):
    data = await state.get_data()
    template = data.get("template")
    user_id = message.from_user.id
    user = get_user(user_id)

    await message.answer("Формирую документ...")

    prompt = f"Составь официальный документ «{template}».\nДанные: {message.text}\nСделай полный текст."
    text = await ask_grok(prompt)

    doc = Document()
    doc.add_heading(template, 0)
    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line.strip())

    path = f"template_{user_id}.docx"
    doc.save(path)

    await message.answer_document(FSInputFile(path), caption=template)
    user["generations"] += 1
    user["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {template}")
    await message.answer("Готово!", reply_markup=main_kb())
    await state.clear()

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
    names = {"free": "Бесплатный", "basic": "Базовый (199 ₽)", "standard": "Стандарт (499 ₽)", "premium": "Премиум (999 ₽)"}
    limit = PLAN_LIMITS.get(user["plan"], 3)
    await message.answer(f"Тариф: {names.get(user['plan'])}\nГенераций: {user['generations']} из {limit}")

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

async def main():
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
