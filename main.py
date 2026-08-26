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

# Лимиты по символам и доп. информации
TOPIC_LIMITS = {
    "free": 100,
    "basic": 150,
    "standard": 300,
    "premium": 600
}

EXTRA_LIMITS = {
    "free": {"times": 0, "chars": 0},
    "basic": {"times": 1, "chars": 300},
    "standard": {"times": 2, "chars": 600},
    "premium": {"times": 3, "chars": 1200}
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
    waiting_extra = State()

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

def get_prompt(plan: str, topic: str, slides: int, style: str, extra: str = "") -> str:
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

    extra_part = f"\nДополнительная информация от пользователя: {extra}" if extra else ""

    return f"""{level}

Тема: {topic}
Количество слайдов: {slides}
Стиль: {style}{extra_part}

Верни ТОЛЬКО валидный JSON:
{{
  "title": "Название презентации",
  "slides": [
    {{"title": "Заголовок слайда", "content": "Текст слайда"}}
  ]
}}"""

def add_placeholder(slide, left, top, width, height, text, tc):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(80, 80, 80)
    shape.line.color.rgb = RGBColor(150, 150, 150)

    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = RGBColor(*tc)
    p.alignment = PP_ALIGN.CENTER

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

def confirm_kb(plan: str, extra_used: int):
    buttons = [[KeyboardButton(text="✅ Делать полную версию")]]
    limits = EXTRA_LIMITS.get(plan, EXTRA_LIMITS["free"])
    if extra_used < limits["times"]:
        buttons.append([KeyboardButton(text="➕ Добавить информацию")])
    buttons.append([KeyboardButton(text="✏️ Изменить тему")])
    buttons.append([KeyboardButton(text="🎨 Изменить стиль")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

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
    u = get_user(m.from_user.id)
    limit = TOPIC_LIMITS.get(u["plan"], 100)
    await m.answer(f"Напиши тему презентации (до {limit} символов):")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_topic)
async def process_topic(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    limit = TOPIC_LIMITS.get(u["plan"], 100)
    text = m.text or ""
    if len(text) > limit:
        await m.answer(f"Слишком длинно. Максимум {limit} символов. Напиши короче:")
        return
    await state.update_data(topic=text, extra="", extra_used=0)
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

    prompt = f"""Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Стиль: {m.text}
Дополнительно: {data.get('extra', '')}

Сделай короткий образец структуры презентации обычным текстом:
Название: ...
1. ...
2. ...
3. ...
Не используй JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)

    u = get_user(m.from_user.id)
    extra_used = data.get("extra_used", 0)
    await m.answer(
        f"Пробный вариант:\n\n{sample}\n\nВыбери действие:",
        reply_markup=confirm_kb(u["plan"], extra_used)
    )
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text == "✏️ Изменить тему")
async def change_topic(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    limit = TOPIC_LIMITS.get(u["plan"], 100)
    await m.answer(f"Напиши новую тему (до {limit} символов):")
    await state.set_state(Form.waiting_topic)

@dp.message(Form.waiting_confirm, F.text == "🎨 Изменить стиль")
async def change_style(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    await m.answer("Выбери новый стиль:", reply_markup=style_kb(u["plan"]))
    await state.set_state(Form.waiting_style)

@dp.message(Form.waiting_confirm, F.text == "➕ Добавить информацию")
async def add_extra(m: Message, state: FSMContext):
    data = await state.get_data()
    u = get_user(m.from_user.id)
    limits = EXTRA_LIMITS.get(u["plan"], EXTRA_LIMITS["free"])
    extra_used = data.get("extra_used", 0)

    if extra_used >= limits["times"]:
        await m.answer("Лимит добавлений информации исчерпан.")
        return

    await m.answer(f"Напиши дополнительную информацию (до {limits['chars']} символов):")
    await state.set_state(Form.waiting_extra)

@dp.message(Form.waiting_extra)
async def process_extra(m: Message, state: FSMContext):
    data = await state.get_data()
    u = get_user(m.from_user.id)
    limits = EXTRA_LIMITS.get(u["plan"], EXTRA_LIMITS["free"])
    text = m.text or ""

    if len(text) > limits["chars"]:
        await m.answer(f"Слишком длинно. Максимум {limits['chars']} символов. Напиши короче:")
        return

    old_extra = data.get("extra", "")
    new_extra = (old_extra + "\n" + text).strip() if old_extra else text
    extra_used = data.get("extra_used", 0) + 1

    await state.update_data(extra=new_extra, extra_used=extra_used)
    await m.answer("Обновляю пробный вариант...")

    prompt = f"""Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Стиль: {data.get('style')}
Дополнительно: {new_extra}

Сделай короткий образец структуры презентации обычным текстом:
Название: ...
1. ...
2. ...
3. ...
Не используй JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)

    await m.answer(
        f"Обновлённый пробный вариант:\n\n{sample}\n\nВыбери действие:",
        reply_markup=confirm_kb(u["plan"], extra_used)
    )
    await state.set_state(Form.waiting_confirm)

@dp.message(Form.waiting_confirm, F.text.in_(["✅ Делать полную версию", "делай", "да", "ок", "хорошо", "подтверждаю"]))
async def confirm_generate(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Делаю финальную версию...")

    prompt = get_prompt(
        u["plan"],
        data.get("topic", ""),
        data.get("slides", 8),
        data.get("style", "Графит"),
        data.get("extra", "")
    )
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

    for idx, s in enumerate(content.get("slides", [])):
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

        cb = slide.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(6.5), Inches(5.5))
        tf = cb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(16)
        p.font.color.rgb = RGBColor(*tc)

        if idx % 2 == 0:
            add_placeholder(slide, 7.5, 1.5, 5.2, 2.5, "Вставь сюда фото", tc)
            add_placeholder(slide, 7.5, 4.3, 5.2, 2.3, "Вставь сюда график", tc)
        else:
            add_placeholder(slide, 7.5, 1.5, 5.2, 5.1, "Вставь сюда фото", tc)

    pptx_path = f"pres_{uid}.pptx"
    prs.save(pptx_path)

    # PDF с кириллицей
    pdf_path = f"pres_{uid}.pdf"
    c = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4

    font_name = "Helvetica"
    font_bold = "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
        font_name = "DejaVu"
        font_bold = "DejaVuBold"
    except:
        pass

    c.setFont(font_bold, 14)
    c.drawString(40, h - 50, content.get("title", "")[:65])
    y = h - 90
    for i, s in enumerate(content.get("slides", []), 1):
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
