import os
import asyncio
import json
import random
from datetime import datetime
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from openai import AsyncOpenAI
from docx import Document
from docx.shared import Pt as DocxPt, Cm, RGBColor as DocxRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
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
    waiting_mode = State()
    waiting_topic = State()
    waiting_user_text = State()
    waiting_theme = State()
    waiting_slides = State()
    waiting_confirm = State()
    waiting_extra = State()
    waiting_style = State()
    waiting_word_mode = State()
    waiting_word_kind = State()
    waiting_word_topic = State()
    waiting_word_text = State()
    waiting_word_size = State()
    waiting_word_confirm = State()
    waiting_word_extra = State()

THEMES = {
    "nature": {"bg": (248, 246, 241), "ink": (22, 22, 24), "mid": (70, 70, 74), "mute": (130, 128, 124), "line": (22, 22, 24),
               "photo": "photorealistic nature photography, cinematic sunlight, no text, no watermark"},
    "business": {"bg": (16, 16, 18), "ink": (245, 245, 247), "mid": (196, 196, 200), "mute": (120, 120, 126), "line": (212, 175, 90),
                 "photo": "premium business photography, architecture, cinematic, no text, no watermark"},
    "tech": {"bg": (8, 16, 28), "ink": (240, 246, 255), "mid": (176, 196, 220), "mute": (110, 130, 155), "line": (90, 170, 230),
             "photo": "futuristic technology photography, cinematic, no text, no watermark"},
    "school": {"bg": (250, 249, 246), "ink": (28, 32, 40), "mid": (60, 64, 72), "mute": (120, 124, 132), "line": (40, 90, 180),
               "photo": "clear educational photo, bright, no text, no watermark"},
    "fashion": {"bg": (252, 250, 247), "ink": (18, 18, 18), "mid": (70, 66, 62), "mute": (140, 134, 128), "line": (18, 18, 18),
                "photo": "editorial fashion photography, magazine look, no text, no watermark"},
    "history": {"bg": (245, 237, 224), "ink": (48, 32, 20), "mid": (90, 70, 50), "mute": (140, 120, 95), "line": (140, 90, 40),
                "photo": "historical documentary photography, museums, archives, cinematic, no text"},
    "science": {"bg": (244, 248, 252), "ink": (16, 32, 56), "mid": (50, 70, 95), "mute": (110, 130, 150), "line": (20, 90, 160),
                "photo": "scientific photography, labs, space, macro details, cinematic, no text"},
    "sport": {"bg": (18, 18, 20), "ink": (250, 250, 252), "mid": (210, 210, 214), "mute": (140, 140, 146), "line": (230, 70, 40),
              "photo": "dynamic sports photography, motion, stadium light, cinematic, no text"},
    "travel": {"bg": (247, 243, 236), "ink": (32, 28, 24), "mid": (80, 70, 60), "mute": (130, 120, 110), "line": (180, 120, 60),
               "photo": "travel photography, cities and landscapes, golden hour, cinematic, no text"},
    "food": {"bg": (252, 248, 242), "ink": (40, 24, 16), "mid": (90, 60, 40), "mute": (140, 110, 90), "line": (180, 80, 40),
             "photo": "food photography, editorial restaurant style, no text, no watermark"},
    "art": {"bg": (20, 18, 22), "ink": (248, 244, 238), "mid": (200, 190, 180), "mute": (140, 130, 125), "line": (220, 180, 120),
            "photo": "art gallery photography, paintings, sculpture, cinematic, no text"},
    "eco": {"bg": (236, 244, 236), "ink": (20, 40, 24), "mid": (50, 80, 55), "mute": (100, 125, 105), "line": (40, 110, 60),
            "photo": "ecology photography, forests, clean energy, cinematic, no text"},
    "minimal": {"bg": (250, 250, 250), "ink": (18, 18, 18), "mid": (70, 70, 70), "mute": (140, 140, 140), "line": (18, 18, 18),
                "photo": "minimalist photography, clean composition, negative space, no text"},
    "default": {"bg": (248, 246, 241), "ink": (22, 22, 24), "mid": (70, 70, 74), "mute": (130, 128, 124), "line": (22, 22, 24),
                "photo": "cinematic photorealistic photo, no text, no watermark"}
}

THEME_LABELS = {
    "nature": "Природа", "business": "Бизнес", "tech": "Технологии", "school": "Учёба",
    "fashion": "Мода", "history": "История", "science": "Наука", "sport": "Спорт",
    "travel": "Путешествия", "food": "Еда", "art": "Искусство", "eco": "Экология",
    "minimal": "Минимализм", "default": "Универсальный"
}

ANGLES = [
    "через неожиданный факт", "через историю одного примера", "через контраст до и после",
    "через вопрос к зрителю", "через три сильных тезиса", "в стиле National Geographic",
    "в стиле Apple keynote", "в стиле модного журнала", "как лекция сильного преподавателя",
    "как премиальный pitch deck"
]


def pick_theme(topic: str):
    t = (topic or "").lower()
    rules = [
        ("nature", ["животн", "природ", "океан", "кит", "лес", "моря", "птиц", "растен"]),
        ("business", ["бизнес", "компани", "продаж", "финанс", "инвест", "стартап", "рынок"]),
        ("tech", ["крипт", "техно", "ai", "ии", "робот", "софт", "нейро", "код", "гаджет"]),
        ("school", ["школ", "универ", "урок", "студент", "доклад", "класс"]),
        ("fashion", ["мод", "стиль", "бренд", "одежд"]),
        ("history", ["истори", "войн", "древн", "импери", "век"]),
        ("science", ["наук", "физик", "хими", "космос", "медицин", "биологи"]),
        ("sport", ["спорт", "футбол", "тренир", "олимп", "матч"]),
        ("travel", ["путешеств", "город", "страна", "туризм", "поездк"]),
        ("food", ["еда", "кухн", "рецепт", "ресторан", "блюд"]),
        ("art", ["искусств", "живопис", "музей", "театр", "музык"]),
        ("eco", ["эколог", "климат", "мусор", "переработ"]),
        ("minimal", ["минимал", "чисто", "просто"]),
    ]
    for name, keys in rules:
        if any(k in t for k in keys):
            return name, THEMES[name]
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
                {"role": "system", "content": "Ты арт-директор презентаций. Пиши как живой сильный автор, не как нейросеть. Без канцелярита и шаблонных фраз: «в современном мире», «является», «следует отметить», «данный», «невозможно переоценить». Чередуй короткие и длинные предложения. Заголовки живые. Каждый ответ уникален. Исправляй ошибки. Только русский."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.95,
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
                    print("Image saved:", path)
                    return True
                if info.get("status") in ("failed", "canceled"):
                    print("Replicate failed:", info)
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
        [KeyboardButton(text="📄 Сделать документ Word")],
        [KeyboardButton(text="📁 Моя история"), KeyboardButton(text="ℹ️ Мой тариф")]
    ], resize_keyboard=True)


def mode_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="✨ Сгенерировать с ИИ")],
        [KeyboardButton(text="📝 Вставить свой текст")]
    ], resize_keyboard=True)


def slides_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="5️⃣ 5 слайдов"), KeyboardButton(text="8️⃣ 8 слайдов")],
        [KeyboardButton(text="🔟 10 слайдов")]
    ], resize_keyboard=True)


def confirm_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="✅ Делать полную версию")],
        [KeyboardButton(text="➕ Добавить или изменить информацию")],
        [KeyboardButton(text="🎨 Изменить стиль")],
        [KeyboardButton(text="✏️ Изменить тему")]
    ], resize_keyboard=True)


def word_kind_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="📄 Обычный документ")],
        [KeyboardButton(text="🎓 Реферат")],
        [KeyboardButton(text="📝 Договор купли-продажи")],
        [KeyboardButton(text="🏠 Договор аренды")],
        [KeyboardButton(text="💼 Коммерческое предложение")],
        [KeyboardButton(text="📋 Заявление")],
        [KeyboardButton(text="🧾 Доверенность")]
    ], resize_keyboard=True)


def word_size_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="Короткий")],
        [KeyboardButton(text="Средний")],
        [KeyboardButton(text="Подробный")]
    ], resize_keyboard=True)


def word_confirm_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="✅ Собрать документ")],
        [KeyboardButton(text="➕ Добавить или изменить информацию")],
        [KeyboardButton(text="✏️ Изменить запрос")]
    ], resize_keyboard=True)


def style_kb(include_keep=False):
    rows = []
    if include_keep:
        rows.append([KeyboardButton(text="👍 Оставить этот стиль")])
    rows.extend([
        [KeyboardButton(text="Природа"), KeyboardButton(text="Бизнес")],
        [KeyboardButton(text="Технологии"), KeyboardButton(text="Учёба")],
        [KeyboardButton(text="История"), KeyboardButton(text="Наука")],
        [KeyboardButton(text="Спорт"), KeyboardButton(text="Путешествия")],
        [KeyboardButton(text="Еда"), KeyboardButton(text="Искусство")],
        [KeyboardButton(text="Экология"), KeyboardButton(text="Минимализм")],
        [KeyboardButton(text="Мода"), KeyboardButton(text="Универсальный")]
    ])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


STYLE_BY_LABEL = {v: k for k, v in THEME_LABELS.items()}


@dp.message(Command("start"))
async def cmd_start(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    u["name"] = m.from_user.first_name or "друг"
    await m.answer(
        f"Привет, {u['name']} 👋\n\n"
        "Я собираю красивые презентации: текст, стиль и живые фото.\n"
        "Можно набросать тему — или прислать свой текст, я его поправлю.\n\n"
        "Нажми кнопку ниже и начнём.",
        reply_markup=main_kb()
    )
    await state.clear()


@dp.message(F.text.in_(["Сделать презентацию", "📊 Сделать презентацию"]))
async def start_pres(m: Message, state: FSMContext):
    if not can_generate(m.from_user.id):
        await m.answer("Лимит генераций закончился.")
        return
    await m.answer(
        "Как делаем презентацию?\n\n"
        "✨ Сгенерировать с ИИ — я сам придумаю текст, стиль и фото.\n"
        "📝 Вставить свой текст — пришли материал, я поправлю ошибки и соберу слайды.",
        reply_markup=mode_kb()
    )
    await state.set_state(Form.waiting_mode)


@dp.message(Form.waiting_mode, F.text.in_(["Сгенерировать с ИИ", "✨ Сгенерировать с ИИ"]))
async def mode_ai(m: Message, state: FSMContext):
    await state.update_data(mode="ai", user_text="")
    await m.answer("Напиши тему. Можно коротко, например: киты, крипта, школа.")
    await state.set_state(Form.waiting_topic)


@dp.message(Form.waiting_mode, F.text.in_(["Вставить свой текст", "📝 Вставить свой текст"]))
async def mode_user(m: Message, state: FSMContext):
    await state.update_data(mode="user")
    await m.answer("Пришли свой текст 📝\nМожно черновик — я поправлю ошибки и соберу структуру.")
    await state.set_state(Form.waiting_user_text)


@dp.message(Form.waiting_topic)
async def process_topic(m: Message, state: FSMContext):
    text = m.text or ""
    if len(text) > 500:
        await m.answer("Слишком длинно. Максимум 500 символов.")
        return
    name, _ = pick_theme(text)
    await state.update_data(topic=text, extra="", extra_used=0, theme_name=name)
    await m.answer(
        f"🎨 По теме подходит стиль: {THEME_LABELS.get(name, name)}.\nОставить его или выбрать другой?",
        reply_markup=style_kb(include_keep=True)
    )
    await state.set_state(Form.waiting_theme)


@dp.message(Form.waiting_user_text)
async def process_user_text(m: Message, state: FSMContext):
    text = m.text or ""
    if len(text) < 40:
        await m.answer("Текста мало. Пришли хотя бы несколько абзацев.")
        return
    if len(text) > 4000:
        await m.answer("Слишком длинно. Сократи до 4000 символов.")
        return
    name, _ = pick_theme(text)
    topic = text[:80].replace("\n", " ")
    await state.update_data(user_text=text, topic=topic, extra="", extra_used=0, theme_name=name)
    await m.answer(
        f"🎨 По тексту подходит стиль: {THEME_LABELS.get(name, name)}.\nОставить его или выбрать другой?",
        reply_markup=style_kb(include_keep=True)
    )
    await state.set_state(Form.waiting_theme)


@dp.message(Form.waiting_theme)
async def process_theme(m: Message, state: FSMContext):
    label = m.text or ""
    if label not in ("Оставить этот стиль", "👍 Оставить этот стиль"):
        name = STYLE_BY_LABEL.get(label)
        if not name:
            await m.answer("Выбери стиль кнопкой ниже.", reply_markup=style_kb(include_keep=True))
            return
        await state.update_data(theme_name=name)
    await m.answer("Сколько слайдов сделать?", reply_markup=slides_kb())
    await state.set_state(Form.waiting_slides)


@dp.message(Form.waiting_slides)
async def process_slides(m: Message, state: FSMContext):
    slides = 8
    t = m.text or ""
    if "5" in t:
        slides = 5
    elif "10" in t:
        slides = 10
    data = await state.get_data()
    angle = random.choice(ANGLES)
    await state.update_data(slides=slides, angle=angle)
    await m.answer("Секунду, собираю пробный вариант…", reply_markup=ReplyKeyboardRemove())

    if data.get("mode") == "user":
        prompt = f"""Пользователь прислал свой текст. Исправь ошибки и собери структуру презентации.
Текст:
{data.get('user_text')}
Слайдов: {slides}
Угол подачи: {angle}
Короткий план текстом: Название + 3 пункта. Без JSON."""
    else:
        prompt = f"""Тема: {data.get('topic')}
Слайдов: {slides}
Доп: {data.get('extra')}
Угол подачи: {angle}
Короткий уникальный план: название и 3 пункта. Без JSON."""

    sample = await ask_grok(prompt)
    theme_name = data.get("theme_name", "default")
    await state.update_data(sample=sample)
    await m.answer(
        f"Черновик готов ✅\n\n{sample}\n\n"
        f"Стиль: {THEME_LABELS.get(theme_name, theme_name)}\n\n"
        "Если всё ок — собираем полную версию.",
        reply_markup=confirm_kb()
    )
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(["Изменить тему", "✏️ Изменить тему"]))
async def change_topic(m: Message, state: FSMContext):
    data = await state.get_data()
    if data.get("mode") == "user":
        await m.answer("Пришли новый текст:")
        await state.set_state(Form.waiting_user_text)
    else:
        await m.answer("Напиши новую тему:")
        await state.set_state(Form.waiting_topic)


@dp.message(Form.waiting_confirm, F.text.in_(["Изменить стиль", "🎨 Изменить стиль"]))
async def change_style(m: Message, state: FSMContext):
    await m.answer("Выбери стиль оформления:", reply_markup=style_kb())
    await state.set_state(Form.waiting_style)


@dp.message(Form.waiting_style)
async def process_style(m: Message, state: FSMContext):
    label = m.text or ""
    name = STYLE_BY_LABEL.get(label)
    if not name:
        await m.answer("Выбери стиль кнопкой ниже.", reply_markup=style_kb())
        return
    data = await state.get_data()
    await state.update_data(theme_name=name)
    await m.answer(
        f"Стиль изменён: {label}\n\nЧерновик:\n\n{data.get('sample')}\n\nВыбери действие:",
        reply_markup=confirm_kb()
    )
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(["Добавить информацию", "➕ Добавить информацию", "➕ Добавить или изменить информацию"]))
async def add_extra(m: Message, state: FSMContext):
    if (await state.get_data()).get("extra_used", 0) >= 3:
        await m.answer("Лимит добавлений исчерпан.", reply_markup=confirm_kb())
        return
    await m.answer("Напиши, что добавить или изменить. Максимум 800 символов.")
    await state.set_state(Form.waiting_extra)


@dp.message(Form.waiting_extra)
async def process_extra(m: Message, state: FSMContext):
    data = await state.get_data()
    text = m.text or ""
    if len(text) > 800:
        await m.answer("Слишком длинно. Максимум 800 символов.")
        return
    extra = ((data.get("extra") or "") + "\n" + text).strip()
    angle = random.choice(ANGLES)
    await state.update_data(extra=extra, extra_used=data.get("extra_used", 0) + 1, angle=angle)
    await m.answer("Обновляю черновик…")
    if data.get("mode") == "user":
        prompt = f"""Исправь и обнови структуру.
Исходный текст:
{data.get('user_text')}
Дополнительно:
{extra}
Слайдов: {data.get('slides')}
Угол: {angle}
Короткий план: название и 3 пункта. Без JSON."""
    else:
        prompt = f"""Тема: {data.get('topic')}
Доп: {extra}
Угол: {angle}
Новый короткий план, 3 пункта. Без JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)
    await m.answer(f"Обновлённый черновик ✅\n\n{sample}\n\nВыбери действие:", reply_markup=confirm_kb())
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(["Делать полную версию", "✅ Делать полную версию", "делай", "да", "ок"]))
async def confirm_generate(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Собираю презентацию с картинками. Это займёт 1–2 минуты.")

    theme_name = data.get("theme_name") or pick_theme(data.get("topic", ""))[0]
    colors = THEMES.get(theme_name, THEMES["default"])
    angle = data.get("angle") or random.choice(ANGLES)
    layouts = [0, 1, 2]
    random.shuffle(layouts)

    if data.get("mode") == "user":
        raw = await ask_grok(f"""Собери уникальную презентацию из текста пользователя.
Исправь ошибки, сохрани смысл.
Текст:
{data.get('user_text')}
Доп:
{data.get('extra')}
Слайдов: {data.get('slides')}
Угол: {angle}
Стиль: {theme_name}
Заголовок слайда 3–6 слов, как у человека, не как у ИИ.
Текст: 2 живых абзаца, без канцелярита.
Фото: живой кадр, не стоковый ИИ-шаблон.
Только JSON:
{{"title":"...","slides":[{{"title":"...","content":"абзац1\\n\\nабзац2","image_prompt":"unique cinematic scene"}}]}}""")
    else:
        raw = await ask_grok(f"""Собери уникальную презентацию уровня лучшего журнала.
Тема: {data.get('topic')}
Слайдов: {data.get('slides')}
Доп: {data.get('extra')}
Угол: {angle}
Стиль: {theme_name}
Заголовок 3–6 слов, живой, не шаблонный.
Текст: 2 живых абзаца, без канцелярита и следов ИИ.
Фото: живой кадр, не стоковый ИИ-шаблон.
Только JSON:
{{"title":"...","slides":[{{"title":"...","content":"абзац1\\n\\nабзац2","image_prompt":"unique cinematic scene"}}]}}""")

    try:
        content = json.loads(raw[raw.find("{"):raw.rfind("}") + 1])
    except Exception:
        await m.answer("Не собрал текст. Нажми ещё раз «Сделать презентацию».", reply_markup=main_kb())
        await state.clear()
        return

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slides_data = content.get("slides", [])
    n = len(slides_data)

    cover_img = None
    cover_src = f"/tmp/{uid}_cover.png"
    if await generate_image(f"{content.get('title')}, wide cinematic opening shot, {colors['photo']}", cover_src):
        cover_wide = f"/tmp/{uid}_cover_w.png"
        cover(cover_src, cover_wide, 1920, 1080)
        cover_img = cover_wide

    images = []
    for i, s in enumerate(slides_data):
        src = f"/tmp/{uid}_{i}_{random.randint(1000, 9999)}.png"
        prompt = f"{s.get('image_prompt') or s.get('title')}, {colors['photo']}, unique composition"
        ok = await generate_image(prompt, src)
        if ok:
            wide, tall = f"{src}_w.png", f"{src}_t.png"
            cover(src, wide, 1920, 1080)
            cover(src, tall, 1260, 1500)
            images.append((wide, tall))
        else:
            images.append(None)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    rect(slide, 0, 0, 13.333, 7.5, colors["bg"])
    if cover_img:
        slide.shapes.add_picture(cover_img, Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
        rect(slide, 0, 4.7, 13.333, 2.8, colors["bg"])
    txt(slide, 0.7, 5.0, 12, 1.5, content.get("title", "Презентация"), 40, colors["ink"], True)
    txt(slide, 0.7, 6.6, 12, 0.4, "01  /  введение", 13, colors["mute"])

    for idx, s in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        rect(slide, 0, 0, 13.333, 7.5, colors["bg"])
        layout = layouts[idx % 3]
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
            txt(slide, 7.05, 6.95, 5.5, 0.3, f"{idx + 2:02}  /  {n + 1:02}", 12, colors["mute"])
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
            txt(slide, 0.7, 6.95, 5.5, 0.3, f"{idx + 2:02}  /  {n + 1:02}", 12, colors["mute"])

    pptx_path = f"pres_{uid}.pptx"
    prs.save(pptx_path)
    pdf_path = f"pres_{uid}.pdf"
    pdf = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    fb = "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
        fb = "DejaVuBold"
    except Exception:
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

    await m.answer_document(FSInputFile(pptx_path), caption="📊 PPTX — открывай этот файл")
    await m.answer_document(FSInputFile(pdf_path), caption="📄 PDF — текстовая копия без фото")
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await m.answer(
        "Готово ✅\n\n"
        "Открывай именно PPTX в PowerPoint, Keynote или Google Презентациях.\n\n"
        "Если на телефоне все фото одинаковые, это не ошибка файла. "
        "Так бывает в предпросмотре Telegram, WPS и встроенных «Документах». "
        "Открой тот же файл на другом устройстве или в нормальном редакторе презентаций.",
        reply_markup=main_kb()
    )
    await state.clear()


def _run(p, text, size=12, bold=False, name="Times New Roman", align=None):
    if align is not None:
        p.alignment = align
    r = p.add_run(text or "")
    r.bold = bold
    r.font.size = DocxPt(size)
    r.font.name = name
    r.font.color.rgb = DocxRGB(0, 0, 0)
    return r


def _p(doc, text, size=12, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=6, first=None, line=1.15):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = DocxPt(before)
    p.paragraph_format.space_after = DocxPt(after)
    p.paragraph_format.line_spacing = line
    if first is not None:
        p.paragraph_format.first_line_indent = Cm(first)
    _run(p, text, size, bold)
    return p


def _body(doc, sections, indent=True, head_center=False):
    for block in sections or []:
        head = (block.get("title") or "").strip()
        body = block.get("content") or ""
        if head:
            _p(
                doc, head, 13, True,
                align=WD_ALIGN_PARAGRAPH.CENTER if head_center else WD_ALIGN_PARAGRAPH.LEFT,
                before=10, after=6
            )
        for para in [x.strip() for x in body.split("\n") if x.strip()]:
            _p(doc, para, 12, first=1.25 if indent else None, after=6, line=1.15)


def build_word(path, title, sections, kind="doc", meta=None):
    meta = meta or {}
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2)
    title = title or "Документ"

    if kind == "referat":
        sec.left_margin = Cm(3)
        sec.right_margin = Cm(1.5)
        _p(doc, meta.get("org") or "Министерство образования", 14, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        _p(doc, meta.get("school") or "[Название учебного заведения]", 14, align=WD_ALIGN_PARAGRAPH.CENTER)
        for _ in range(4):
            doc.add_paragraph()
        _p(doc, "РЕФЕРАТ", 20, True, WD_ALIGN_PARAGRAPH.CENTER, after=8)
        _p(doc, title, 16, True, WD_ALIGN_PARAGRAPH.CENTER)
        for _ in range(6):
            doc.add_paragraph()
        _p(doc, f"Выполнил: {meta.get('author') or '[ФИО студента]'}", 14, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0)
        _p(doc, f"Проверил: {meta.get('teacher') or '[ФИО преподавателя]'}", 14, align=WD_ALIGN_PARAGRAPH.RIGHT)
        for _ in range(6):
            doc.add_paragraph()
        _p(doc, f"{meta.get('city') or '[Город]'} {meta.get('year') or '2026'}", 14, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_page_break()
        _body(doc, sections, indent=True, head_center=True)

    elif kind in ("dkp", "rent"):
        cap = "ДОГОВОР КУПЛИ-ПРОДАЖИ" if kind == "dkp" else "ДОГОВОР АРЕНДЫ"
        _p(doc, cap, 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, title if title not in (cap, "Документ") else "№ ______", 12, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        _body(doc, sections, indent=True, head_center=False)
        left = "ПРОДАВЕЦ" if kind == "dkp" else "АРЕНДОДАТЕЛЬ"
        right = "ПОКУПАТЕЛЬ" if kind == "dkp" else "АРЕНДАТОР"
        _p(doc, f"{left}                    {right}", 12, True, before=18)
        _p(doc, "__________ / [ФИО] /          __________ / [ФИО] /", 12)

    elif kind == "offer":
        _p(doc, "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", 11, True, after=0)
        _p(doc, meta.get("number") or "", 10, after=12)
        _p(doc, title, 20, True, after=4)
        _body(doc, sections, indent=False, head_center=False)

    elif kind == "statement":
        _p(doc, meta.get("to") or "Директору [организация]", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0)
        _p(doc, meta.get("to_name") or "[ФИО руководителя]", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0)
        _p(doc, f"от {meta.get('from') or '[должность, ФИО]'}", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=18)
        _p(doc, "ЗАЯВЛЕНИЕ", 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=14)
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, f"{meta.get('date') or '[дата]'}                    __________ / {meta.get('author') or '[ФИО]'} /", 12, before=24)

    elif kind == "proxy":
        _p(doc, "ДОВЕРЕННОСТЬ", 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, f"Подпись доверителя: __________ / {meta.get('author') or '[ФИО]'} /", 12, before=20)

    else:
        _p(doc, meta.get("org") or "", 11, True, after=0)
        _p(doc, title, 16, True, WD_ALIGN_PARAGRAPH.CENTER, before=8, after=12)
        _body(doc, sections, indent=True, head_center=False)
        if meta.get("sign"):
            _p(doc, "С уважением,", 12, before=16, after=0)
            _p(doc, meta.get("sign"), 12)

    doc.save(path)


@dp.message(F.text.in_(["Сделать документ Word", "📄 Сделать документ Word"]))
async def start_word(m: Message, state: FSMContext):
    if not can_generate(m.from_user.id):
        await m.answer("Лимит генераций закончился.")
        return
    await m.answer("Какой документ нужен?", reply_markup=word_kind_kb())
    await state.set_state(Form.waiting_word_kind)


@dp.message(Form.waiting_word_kind)
async def word_kind(m: Message, state: FSMContext):
    t = (m.text or "").lower()
    kind = "doc"
    if "реферат" in t:
        kind = "referat"
    elif "купл" in t:
        kind = "dkp"
    elif "аренд" in t:
        kind = "rent"
    elif "коммерч" in t or "предлож" in t:
        kind = "offer"
    elif "заявлен" in t:
        kind = "statement"
    elif "доверен" in t:
        kind = "proxy"
    await state.update_data(word_kind=kind, extra="", extra_used=0)
    await m.answer(
        "Как собираем документ?\n\n"
        "✨ Сгенерировать с ИИ — я сам напишу текст.\n"
        "📝 Вставить свой текст — пришли материал, я поправлю и оформлю.",
        reply_markup=mode_kb()
    )
    await state.set_state(Form.waiting_word_mode)


@dp.message(Form.waiting_word_mode, F.text.in_(["Сгенерировать с ИИ", "✨ Сгенерировать с ИИ"]))
async def word_mode_ai(m: Message, state: FSMContext):
    await state.update_data(mode="ai", user_text="")
    await m.answer("Напиши тему документа. Например: доклад про китов или договор на машину.")
    await state.set_state(Form.waiting_word_topic)


@dp.message(Form.waiting_word_mode, F.text.in_(["Вставить свой текст", "📝 Вставить свой текст"]))
async def word_mode_user(m: Message, state: FSMContext):
    await state.update_data(mode="user")
    await m.answer("Пришли текст. Для договора укажи стороны, предмет, цену и дату, если они уже есть.")
    await state.set_state(Form.waiting_word_text)


@dp.message(Form.waiting_word_topic)
async def word_topic(m: Message, state: FSMContext):
    text = m.text or ""
    if len(text) > 500:
        await m.answer("Слишком длинно. Максимум 500 символов.")
        return
    await state.update_data(topic=text)
    await m.answer("Какой объём?", reply_markup=word_size_kb())
    await state.set_state(Form.waiting_word_size)


@dp.message(Form.waiting_word_text)
async def word_user_text(m: Message, state: FSMContext):
    text = m.text or ""
    if len(text) < 30:
        await m.answer("Текста мало. Пришли чуть больше деталей.")
        return
    if len(text) > 4000:
        await m.answer("Слишком длинно. Сократи до 4000 символов.")
        return
    await state.update_data(user_text=text, topic=text[:80].replace("\n", " "))
    await m.answer("Какой объём?", reply_markup=word_size_kb())
    await state.set_state(Form.waiting_word_size)


@dp.message(Form.waiting_word_size)
async def word_size(m: Message, state: FSMContext):
    t = (m.text or "").lower()
    size = "short"
    if "подр" in t:
        size = "long"
    elif "сред" in t:
        size = "medium"
    data = await state.get_data()
    await state.update_data(word_size=size)
    await m.answer("Собираю черновик…", reply_markup=ReplyKeyboardRemove())
    size_map = {"short": "1–2 страницы, коротко", "medium": "2–4 страницы, нормально", "long": "4–6 страниц, подробно"}
    kind = data.get("word_kind", "doc")
    kind_name = {
        "doc": "обычный документ",
        "referat": "реферат для университета",
        "dkp": "договор купли-продажи",
        "rent": "договор аренды",
        "offer": "коммерческое предложение: о компании, задача клиента, решение и услуги, этапы, сроки, бюджет, контакты",
        "statement": "заявление",
        "proxy": "доверенность",
    }.get(kind, "документ")
    prompt = f"""Собери черновик: {kind_name}.
Тема/данные: {data.get('topic')}
Текст пользователя: {data.get('user_text')}
Доп: {data.get('extra')}
Объём: {size_map[size]}
Исправь ошибки. Если данных не хватает, поставь пропуски [указать ...].
Обычным текстом: название и 4 пункта структуры. Без JSON."""
    sample = await ask_grok(prompt)
    await state.update_data(sample=sample)
    await m.answer(f"Черновик готов ✅\n\n{sample}\n\nЕсли всё ок — собираем файл.", reply_markup=word_confirm_kb())
    await state.set_state(Form.waiting_word_confirm)


@dp.message(Form.waiting_word_confirm, F.text.in_(["Изменить запрос", "✏️ Изменить запрос"]))
async def word_change(m: Message, state: FSMContext):
    data = await state.get_data()
    if data.get("mode") == "user":
        await m.answer("Пришли новый текст:")
        await state.set_state(Form.waiting_word_text)
    else:
        await m.answer("Напиши новую тему:")
        await state.set_state(Form.waiting_word_topic)


@dp.message(Form.waiting_word_confirm, F.text.in_(["Добавить информацию", "➕ Добавить информацию", "➕ Добавить или изменить информацию"]))
async def word_add(m: Message, state: FSMContext):
    if (await state.get_data()).get("extra_used", 0) >= 3:
        await m.answer("Лимит добавлений исчерпан.", reply_markup=word_confirm_kb())
        return
    await m.answer("Напиши, что добавить или изменить. Максимум 800 символов.")
    await state.set_state(Form.waiting_word_extra)


@dp.message(Form.waiting_word_extra)
async def word_extra(m: Message, state: FSMContext):
    data = await state.get_data()
    text = m.text or ""
    if len(text) > 800:
        await m.answer("Слишком длинно. Максимум 800 символов.")
        return
    extra = ((data.get("extra") or "") + "\n" + text).strip()
    await state.update_data(extra=extra, extra_used=data.get("extra_used", 0) + 1)
    await m.answer("Обновляю черновик…")
    sample = await ask_grok(
        f"Обнови черновик документа.\nТема: {data.get('topic')}\nТекст: {data.get('user_text')}\nДоп: {extra}\nКороткий план. Без JSON."
    )
    await state.update_data(sample=sample)
    await m.answer(f"Обновлённый черновик ✅\n\n{sample}\n\nВыбери действие:", reply_markup=word_confirm_kb())
    await state.set_state(Form.waiting_word_confirm)


@dp.message(Form.waiting_word_confirm, F.text.in_(["Собрать документ", "✅ Собрать документ", "делай", "да", "ок"]))
async def word_build(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    await m.answer("Собираю Word. Обычно это быстрее презентации.")
    kind = data.get("word_kind", "doc")
    size = data.get("word_size", "medium")
    size_map = {"short": "короткий", "medium": "средний", "long": "подробный"}
    kind_name = {
        "doc": "обычный документ",
        "referat": "реферат для университета. Титульный лист, содержание, введение, 2–3 главы, заключение, список литературы. Текст живой, связный, как для сдачи преподавателю, не набор абзацев.",
        "dkp": "договор купли-продажи",
        "rent": "договор аренды",
        "offer": "коммерческое предложение: о компании, задача клиента, решение и услуги, этапы, сроки, бюджет, контакты",
        "statement": "заявление",
        "proxy": "доверенность",
    }.get(kind, "документ")
    raw = await ask_grok(f"""Собери Word: {kind_name}.
Данные: {data.get('topic')}
Текст пользователя: {data.get('user_text')}
Доп: {data.get('extra')}
Объём: {size_map[size]}
Пиши как живой человек. Недостающие данные: [указать ...].
Только JSON:
{{"title":"...","meta":{{"org":"Министерство образования","school":"[учебное заведение]","author":"[ФИО]","teacher":"[преподаватель]","city":"[город]","year":"2026"}},"sections":[{{"title":"Введение","content":"абзац1\\n\\nабзац2"}}]}}""")
    try:
        content = json.loads(raw[raw.find("{"):raw.rfind("}") + 1])
    except Exception:
        await m.answer("Не собрал текст. Попробуй ещё раз.", reply_markup=main_kb())
        await state.clear()
        return

    docx_path = f"doc_{uid}.docx"
    build_word(
        docx_path,
        content.get("title", "Документ"),
        content.get("sections", []),
        kind,
        content.get("meta") or {}
    )

    pdf_path = f"doc_{uid}.pdf"
    pdf = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    fn, fb = "Helvetica", "Helvetica-Bold"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
        pdfmetrics.registerFont(TTFont("DejaVuBold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))
        fn, fb = "DejaVu", "DejaVuBold"
    except Exception:
        pass
    pdf.setFont(fb, 14)
    pdf.drawString(40, h - 50, content.get("title", "")[:65])
    y = h - 80
    for s in content.get("sections", []):
        if y < 70:
            pdf.showPage()
            y = h - 50
        pdf.setFont(fb, 11)
        pdf.drawString(40, y, (s.get("title") or "")[:70])
        y -= 16
        pdf.setFont(fn, 9)
        text = (s.get("content") or "")[:400]
        while text:
            if y < 50:
                pdf.showPage()
                y = h - 50
            pdf.drawString(40, y, text[:90])
            text = text[90:]
            y -= 12
        y -= 10
    pdf.save()

    await m.answer_document(FSInputFile(docx_path), caption="📄 Word — этот файл можно править")
    await m.answer_document(FSInputFile(pdf_path), caption="📄 PDF-копия")
    u["generations"] += 1
    u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
    await m.answer("Документ готов ✅", reply_markup=main_kb())
    await state.clear()


@dp.message(F.text.in_(["Моя история", "📁 Моя история"]))
async def history(m: Message):
    u = get_user(m.from_user.id)
    await m.answer("История пустая." if not u["history"] else "📁 История:\n\n" + "\n".join(u["history"][-10:]))


@dp.message(F.text.in_(["Мой тариф", "ℹ️ Мой тариф"]))
async def my_plan(m: Message):
    u = get_user(m.from_user.id)
    limit = PLAN_LIMITS.get(u["plan"], 15)
    await m.answer(f"ℹ️ Генераций: {u['generations']} из {limit}\nОсталось: {max(0, limit - u['generations'])}")


@dp.message(Command("grant"))
async def grant(m: Message):
    if m.from_user.id not in ADMIN_IDS:
        return
    try:
        uid = int(m.text.split()[1])
        get_user(uid)["plan"] = "premium"
        await m.answer(f"Доступ выдан пользователю {uid}")
    except Exception:
        await m.answer("Формат: /grant user_id")


async def main():
    print("Бот запущен")
    print("REPLICATE TOKEN:", "YES" if REPLICATE_API_TOKEN else "NO")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
