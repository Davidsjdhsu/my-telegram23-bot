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
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
ADMIN_IDS = [909828109]

client = AsyncOpenAI(api_key=XAI_API_KEY, base_url="https://api.x.ai/v1")
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())
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

def get_user(uid):
    if uid not in users_db:
        users_db[uid] = {"name": "", "category": None, "plan": "free", "generations": 0, "history": []}
    return users_db[uid]

def can_generate(uid):
    u = get_user(uid)
    return u["generations"] < PLAN_LIMITS.get(u["plan"], 3)

async def ask_grok(prompt: str) -> str:
    try:
        r = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный помощник. Отвечай только на русском."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=3500
        )
        return r.choices[0].message.content
    except Exception as e:
        return f"Ошибка: {e}"

def get_prompt(plan, topic, slides, style):
    if plan == "premium":
        level = (
            "Ты топовый презентационный стратег уровня McKinsey. "
            "Сделай максимально сильную презентацию: мощные заголовки, "
            "глубокий текст (4–7 предложений), storytelling, каждый слайд ценный."
        )
    elif plan == "standard":
        level = (
            "Сделай качественную презентацию (примерно на 50% лучше обычной): "
            "сильные заголовки, хороший объём текста (3–5 предложений), логичная структура."
        )
    else:
        level = (
            "Сделай нормальную аккуратную презентацию: "
            "понятные заголовки, краткий полезный текст (2–4 предложения)."
        )
    return f"""{level}

Тема: {topic}
Количество слайдов: {slides}
Стиль: {style}

Верни ТОЛЬКО валидный JSON:
{{
  "title": "Название",
  "slides": [{{"title": "Заголовок", "content": "Текст"}}]
}}"""

def make_chart(path):
    try:
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.bar(["Этап 1", "Этап 2", "Этап 3", "Этап 4"], [30, 55, 70, 90], color="#4A90E2")
        ax.set_title("Динамика")
        ax.set_ylim(0, 100)
        plt.tight_layout()
        plt.savefig(path, dpi=120, bbox_inches="tight", facecolor="white")
        plt.close()
        return True
    except:
        return False

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

def style_kb(plan):
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=s)] for s in STYLES.get(plan, STYLES["free"])],
        resize_keyboard=True
    )

def templates_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="Европротокол")],
        [KeyboardButton(text="Договор купли-продажи авто")],
        [KeyboardButton(text="Карточка предприятия")],
        [KeyboardButton(text="Акт выполненных работ")],
        [KeyboardButton(text="⬅️ Назад")]
    ], resize_keyboard=True)

# ====================== START ======================
@dp.message(Command("start"))
async def cmd_start(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    u["name"] = m.from_user.first_name or "друг"
    await m.answer(f"Привет, {u['name']}! 👋\n\nВыбери категорию:", reply_markup=category_kb())
    await state.set_state(Form.waiting_category)

@dp.message(Form.waiting_category)
async def process_category(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    t = m.text or ""
    if "Студент" in t: u["category"] = "student"
    elif "Предприниматель" in t: u["category"] = "business"
    elif "Для себя" in t: u["category"] = "personal"
    else:
        await m.answer("Выбери кнопку.")
        return
    await m.answer("Что будем делать?", reply_markup=main_kb())
    await state.clear()

# ====================== ПРЕЗЕНТАЦИЯ ======================
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
    if "5" in (m.text or ""): slides = 5
    elif "10" in (m.text or ""): slides = 10
    elif "15" in (m.text or ""): slides = 15
    await state.update_data(slides=slides)
    u = get_user(m.from_user.id)
    await m.answer("Выбери стиль:", reply_markup=style_kb(u["plan"]))
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_style)
async def process_style(m: Message, state: FSMContext):
    data = await state.get_data()
    await state.update_data(style=m.text)
    await m.answer("Делаю пробный вариант...")
    prompt = (
        f"Тема: {data['topic']}\nСлайдов: {data['slides']}\nСтиль: {m.text}\n\n"
        "Сделай короткий красивый образец структуры обычным текстом (не JSON)."
    )
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
        content = json.loads(raw[raw.find("{"):raw.rfind("}")+1])
    except:
        await m.answer("Ошибка генерации. Попробуй ещё раз.")
        await state.clear()
        return

    bg = STYLE_COLORS.get(data.get("style"), (40, 40, 40))
    tc = (230, 230, 230) if sum(bg) < 300 else (30, 30, 30)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*bg)
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.7), Inches(2.2))
    p = tb.text_frame.paragraphs[0]
    p.text = content.get("title", "Презентация")
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*tc)
    p.alignment = PP_ALIGN.CENTER

    chart_path = f"chart_{uid}.png"
    has_chart = make_chart(chart_path)

    for idx, s in enumerate(content.get("slides", [])):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg_shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = RGBColor(*bg)

        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.1), Inches(1))
        p = tb.text_frame.paragraphs[0]
        p.text = s.get("title", "")
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(*tc)

        if has_chart and idx == 1 and u["plan"] in ("standard", "premium"):
            try:
                slide.shapes.add_picture(chart_path, Inches(1.5), Inches(1.6), width=Inches(10))
            except:
                pass
        else:
            cb = slide.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12.1), Inches(5.3))
            tf = cb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = s.get("content", "")
            p.font.size = Pt(18)
            p.font.color.rgb = RGBColor(*tc)

    pptx_path = f"pres_{uid}.pptx"
    prs.save(pptx_path)

    # Simple PDF
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
        c.drawString(40, y, (s.get("content", "") or "")[:95])
        y -= 22
    c.save()

    await m.answer_document(FSInputFile(pptx_path), caption="PPTX")
    await m.answer_document(FSInputFile(pdf_path), caption="PDF")
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await m.answer("Готово!", reply_markup=main_kb())
    await state.clear()

# ====================== ДОКУМЕНТ ======================
@dp.message(F.text == "📄 Сделать документ")
async def start_doc(m: Message, state: FSMContext):
    if not can_generate(m.from_user.id):
        await m.answer("Лимит генераций закончился.")
        return
    await m.answer("Напиши, какой документ нужен:")
    await state.set_state(Form.waiting_doc_topic)

@dp.message(Form.waiting_doc_topic)
async def process_doc(m: Message, state: FSMContext):
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Делаю документ...")
    text = await ask_grok(f"Создай полноценный структурированный документ на тему: {m.text}")
    doc = Document()
    doc.add_heading((m.text or "Документ")[:80], 0)
    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line.strip())
    path = f"doc_{uid}.docx"
    doc.save(path)
    await m.answer_document(FSInputFile(path), caption="Документ")
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — Документ")
    await m.answer("Готово!", reply_markup=main_kb())
    await state.clear()

# ====================== ШАБЛОНЫ ======================
@dp.message(F.text == "📋 Готовые шаблоны")
async def templates_menu(m: Message):
    await m.answer("Выбери шаблон:", reply_markup=templates_kb())

@dp.message(F.text == "⬅️ Назад")
async def back(m: Message, state: FSMContext):
    await state.clear()
    await m.answer("Главное меню:", reply_markup=main_kb())

@dp.message(F.text.in_(["Европротокол", "Договор купли-продажи авто", "Карточка предприятия", "Акт выполненных работ"]))
async def choose_template(m: Message, state: FSMContext):
    await state.update_data(template=m.text)
    hints = {
        "Европротокол": "Напиши данные:\nМесто ДТП, дата, время, марки авто, ФИО водителей, адреса, полисы, повреждения",
        "Договор купли-продажи авто": "Напиши данные:\nФИО продавца и покупателя, паспортные данные, марка/модель, VIN, год, пробег, цена, дата",
        "Карточка предприятия": "Напиши данные:\nНазвание, ИНН, ОГРН, адрес, директор, телефон, email",
        "Акт выполненных работ": "Напиши данные:\nЗаказчик, исполнитель, перечень работ, сумма, дата"
    }
    await m.answer(hints.get(m.text, "Напиши данные:"))
    await state.set_state(Form.waiting_template_data)

@dp.message(Form.waiting_template_data)
async def process_template(m: Message, state: FSMContext):
    data = await state.get_data()
    template = data.get("template", "Документ")
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Формирую документ...")

    prompt = f"""Составь официальный документ «{template}» на основе этих данных:
{m.text}

Сделай полный структурированный текст документа на русском языке, готовый к использованию.
"""
    text = await ask_grok(prompt)

    doc = Document()
    doc.add_heading(template, 0)
    for line in text.split("\n"):
        if line.strip():
            doc.add_paragraph(line.strip())
    path = f"template_{uid}.docx"
    doc.save(path)

    await m.answer_document(FSInputFile(path), caption=template)
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {template}")
    await m.answer("Готово!", reply_markup=main_kb())
    await state.clear()

# ====================== ИСТОРИЯ / ТАРИФ ======================
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
    await m.answer(f"Тариф: {names.get(u['plan'])}\nГенераций: {u['generations']} из {PLAN_LIMITS.get(u['plan'], 3)}")

# ====================== АДМИНКА ======================
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
