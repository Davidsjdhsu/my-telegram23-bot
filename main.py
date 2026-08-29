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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_CANDIDATES = [
    (os.path.join(BASE_DIR, "fonts", "DejaVuSans.ttf"), os.path.join(BASE_DIR, "fonts", "DejaVuSans-Bold.ttf")),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
]


def safe_filename(title: str, fallback: str = "Документ") -> str:
    """Делает из темы/заголовка документа безопасное имя файла:
    убирает символы, которые нельзя использовать в имени файла (Windows/Telegram),
    обрезает длину, чтобы не упереться в лимиты, и подставляет fallback, если
    заголовок пустой."""
    import re as _re
    text = (title or "").strip() or fallback
    text = _re.sub(r'[\\/:*?"<>|\n\r\t]', " ", text)
    text = _re.sub(r"\s+", " ", text).strip()
    text = text[:80].strip() or fallback
    return text


def wrap_lines(text, font, size, max_width):
    """Переносит текст по словам, а не по количеству символов,
    чтобы слова не рвались посередине."""
    words = text.split()
    lines, cur = [], ""
    for word in words:
        test = f"{cur} {word}".strip()
        if pdfmetrics.stringWidth(test, font, size) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    return lines


def register_pdf_fonts():
    """Регистрирует шрифт с поддержкой кириллицы для reportlab.
    Сначала пробует шрифт, который лежит рядом со скриптом (папка fonts/),
    затем — типичные системные пути. Если ничего не нашлось, возвращает
    базовые Helvetica-шрифты и печатает предупреждение, чтобы не было
    тихой поломки кириллицы в PDF."""
    for regular, bold in FONT_CANDIDATES:
        if os.path.exists(regular) and os.path.exists(bold):
            try:
                pdfmetrics.registerFont(TTFont("CyrRegular", regular))
                pdfmetrics.registerFont(TTFont("CyrBold", bold))
                return "CyrRegular", "CyrBold"
            except Exception as e:
                print("Не удалось зарегистрировать шрифт", regular, ":", e)
    print("ВНИМАНИЕ: шрифт с кириллицей не найден — русский текст в PDF будет нечитаемым. "
          "Положи DejaVuSans.ttf и DejaVuSans-Bold.ttf в папку fonts/ рядом со скриптом.")
    return "Helvetica", "Helvetica-Bold"


BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
REPLICATE_API_TOKEN = os.getenv("REPLICATE_API_TOKEN")
ADMIN_IDS = [909828109]

if not BOT_TOKEN:
    raise RuntimeError("BOT_TOKEN не задан в переменных окружения")
if not XAI_API_KEY:
    raise RuntimeError("XAI_API_KEY не задан в переменных окружения")
if not REPLICATE_API_TOKEN:
    print("ВНИМАНИЕ: REPLICATE_API_TOKEN не задан — генерация изображений будет недоступна")

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
    waiting_word_category = State()
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
        users_db[uid] = {"name": "", "plan": "premium", "generations": 0, "history": [], "busy": False}
    return users_db[uid]


def can_generate(uid):
    u = get_user(uid)
    return u["generations"] < PLAN_LIMITS.get(u["plan"], 15)


def start_job(uid):
    """Помечает пользователя как занятого дорогой генерацией.
    Возвращает False, если задача уже выполняется (защита от двойного нажатия)."""
    u = get_user(uid)
    if u.get("busy"):
        return False
    u["busy"] = True
    return True


def finish_job(uid):
    get_user(uid)["busy"] = False


# Telegram режет любое сообщение длиннее 4096 символов и просто не отправляет его
# (ошибка на стороне API), а не обрезает — из-за этого черновик документа при
# "полном раскрытии темы" (когда нейросеть пишет заметно больше, чем короткий план)
# иногда падал в общий обработчик ошибок с "Что-то пошло не так. Попробуй ещё раз".
# send_draft() при коротком черновике отправляет его текстом как раньше, а при
# длинном — красивым PDF-файлом с именем по теме, а не рвёт на несколько сообщений.
TELEGRAM_MAX_LEN = 4000  # с запасом от лимита в 4096


async def send_draft(m: Message, text: str, title: str = "Черновик", reply_markup=None, note: str = ""):
    """Отправляет черновик пользователю: если он укладывается в лимит Telegram —
    обычным сообщением, а если длиннее (типично для "полного раскрытия темы") —
    красивым PDF-файлом с именем по теме документа, чтобы не дробить текст на
    несколько сообщений и не терять читаемость."""
    if len(text) <= TELEGRAM_MAX_LEN:
        await m.answer(text, reply_markup=reply_markup)
        return
    uid = m.from_user.id
    pdf_path = f"/tmp/draft_{uid}.pdf"
    pdf = canvas.Canvas(pdf_path, pagesize=A4)
    w, h = A4
    fn, fb = register_pdf_fonts()
    pdf.setFont(fb, 14)
    pdf.drawString(40, h - 50, safe_filename(title)[:65])
    y = h - 80
    pdf.setFont(fn, 10)
    for para in text.split("\n"):
        if not para.strip():
            y -= 10
            continue
        for line in wrap_lines(para, fn, 10, w - 80):
            if y < 50:
                pdf.showPage()
                y = h - 50
                pdf.setFont(fn, 10)
            pdf.drawString(40, y, line)
            y -= 14
    pdf.save()
    fname = safe_filename(title, fallback="Черновик")
    caption = f"📄 Черновик получился объёмным, поэтому вот файлом{(' — ' + note) if note else ''}"
    await m.answer_document(FSInputFile(pdf_path, filename=f"{fname} - черновик.pdf"), caption=caption, reply_markup=reply_markup)
    try:
        os.remove(pdf_path)
    except OSError:
        pass


def _balanced_json_spans(text: str):
    """Находит все сбалансированные по фигурным скобкам блоки {...} в тексте,
    корректно игнорируя скобки внутри строк."""
    spans = []
    depth = 0
    in_string = False
    escape = False
    start = None
    for i, ch in enumerate(text):
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and start is not None:
                    spans.append((start, i + 1))
    return spans


def extract_json(raw: str):
    """Надёжно достаёт JSON-объект из ответа модели: убирает markdown-обёртку
    ```json ... ``` и перебирает все сбалансированные по скобкам блоки
    (от самого большого к самому маленькому), а не просто первую/последнюю
    скобку в тексте — так лишний текст модели вокруг JSON не ломает разбор."""
    if not raw:
        raise ValueError("Пустой ответ модели")
    text = raw.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    spans = _balanced_json_spans(text)
    if not spans:
        raise ValueError("В ответе нет JSON-объекта")
    for start, end in sorted(spans, key=lambda s: s[1] - s[0], reverse=True):
        try:
            return json.loads(text[start:end])
        except json.JSONDecodeError:
            continue
    raise ValueError("Не найден валидный JSON-блок")


GROK_ERROR_PREFIX = "__GROK_ERROR__"


async def ask_grok(prompt: str, max_tokens: int = 4000) -> str:
    try:
        r = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты арт-директор презентаций. Пиши как живой сильный автор, не как нейросеть, и так, чтобы текст не триггерил детекторы ИИ-генерации: без канцелярита и шаблонных фраз («в современном мире», «является», «следует отметить», «данный», «невозможно переоценить», «таким образом», «подводя итог»), без идеально симметричной структуры абзацев и слишком гладких переходов. Чередуй короткие и длинные предложения неравномерно. Заголовки живые. Каждый ответ уникален. Исправляй ошибки. Только русский."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.95,
            max_tokens=max_tokens
        )
        return r.choices[0].message.content
    except Exception as e:
        print("Grok API error:", e)
        return f"{GROK_ERROR_PREFIX}{e}"


def grok_failed(text: str) -> bool:
    return (text or "").startswith(GROK_ERROR_PREFIX)


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


def mode_kb(show_template=False):
    rows = [
        [KeyboardButton(text="✨ Сгенерировать с ИИ")],
        [KeyboardButton(text="📝 Вставить свой текст")],
    ]
    if show_template:
        rows.append([KeyboardButton(text="📄 Скачать шаблон")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


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


# Полный банк из 30 видов документов, разбитых на 3 категории (физлица/юрлица/учёба).
# В каждой категории есть "основные" (видны сразу) и "дополнительные" (открываются
# кнопкой "Показать ещё документы"). Итого 8 основных на весь бот + 22 дополнительных.
STUDY_KINDS = {"doc", "referat", "report", "essay", "notes", "coursework"}

# Инструкция, снижающая вероятность срабатывания детекторов ИИ-текста (не антиплагиата —
# антиплагиат по совпадениям и так не страдает, т.к. текст пишется с нуля). Применяется
# только к учебным документам и презентациям, где стиль текста имеет значение;
# юридические документы намеренно формальные и канцелярские по своей природе, их не трогаем.
ANTI_AI_DETECTOR_STYLE = (
    "Пиши так, чтобы текст не выглядел сгенерированным нейросетью и не триггерил "
    "детекторы ИИ-текста: избегай слишком гладких переходов и идеально симметричной "
    "структуры абзацев, чередуй короткие и длинные предложения неравномерно, не начинай "
    "несколько абзацев одинаковыми конструкциями, избегай канцелярских клише "
    "('в современном мире', 'является', 'следует отметить', 'таким образом', "
    "'нельзя переоценить', 'играет важную роль', 'подводя итог'). "
    "Формулировки должны звучать как у обычного студента: где уместно — чуть менее "
    "выверенные обороты, конкретные примеры и детали по теме, живая логика рассуждения, "
    "а не шаблонная академическая гладкость."
)

WORD_CATEGORIES = {
    "physical": {
        "title": "👤 Для физлиц",
        "main": ["dkp", "rent", "proxy", "statement"],
        "more": ["act", "loan", "claim", "consent", "marriage_contract", "gift", "lawsuit", "alimony"],
    },
    "entity": {
        "title": "🏢 Для юрлиц / ИП",
        "main": ["offer", "services", "employment", "work_act"],
        "more": ["supply", "agency", "joint_activity", "nonresidential_rent", "cession", "nda",
                 "self_employed", "warranty_letter"],
    },
    "study": {
        "title": "🎓 Для учёбы",
        "main": ["doc", "referat", "report", "essay", "notes", "coursework"],
        "more": [],
    },
}

KIND_LABELS = {
    "doc": "📄 Обычный документ",
    "referat": "🎓 Реферат",
    "report": "🎤 Доклад",
    "essay": "✍️ Эссе",
    "notes": "📓 Конспект",
    "coursework": "📚 Курсовая работа",
    "dkp": "📝 Договор купли-продажи",
    "rent": "🏠 Договор аренды",
    "proxy": "🧾 Доверенность",
    "statement": "📋 Заявление",
    "act": "📦 Акт приёма-передачи",
    "loan": "💰 Расписка / договор займа",
    "claim": "⚠️ Претензия",
    "consent": "✈️ Согласие на выезд ребёнка",
    "marriage_contract": "💍 Брачный договор",
    "gift": "🎁 Договор дарения",
    "lawsuit": "⚖️ Исковое заявление",
    "alimony": "👶 Соглашение об алиментах",
    "offer": "💼 Коммерческое предложение",
    "services": "🤝 Договор оказания услуг",
    "employment": "👔 Трудовой договор",
    "work_act": "✅ Акт выполненных работ",
    "supply": "🚚 Договор поставки",
    "agency": "🕴️ Агентский договор",
    "joint_activity": "🤝 Договор о совместной деятельности",
    "nonresidential_rent": "🏢 Аренда нежилого помещения",
    "cession": "🔄 Договор цессии",
    "nda": "🔒 Соглашение о неразглашении (NDA)",
    "self_employed": "🧑‍💻 Договор с самозанятым",
    "warranty_letter": "✉️ Гарантийное письмо",
}
LABEL_TO_KIND = {v: k for k, v in KIND_LABELS.items()}

BTN_MORE_DOCS = "➕ Показать ещё документы"
BTN_BACK_TO_CATEGORIES = "⬅️ Назад к категориям"


def word_category_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=WORD_CATEGORIES["physical"]["title"])],
        [KeyboardButton(text=WORD_CATEGORIES["entity"]["title"])],
        [KeyboardButton(text=WORD_CATEGORIES["study"]["title"])],
    ], resize_keyboard=True)


def word_kind_kb(category, more=False):
    cat = WORD_CATEGORIES.get(category, WORD_CATEGORIES["physical"])
    ids = list(cat["main"]) + (list(cat["more"]) if more else [])
    rows = [[KeyboardButton(text=KIND_LABELS[k])] for k in ids]
    if cat["more"] and not more:
        rows.append([KeyboardButton(text=BTN_MORE_DOCS)])
    rows.append([KeyboardButton(text=BTN_BACK_TO_CATEGORIES)])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def word_size_kb():
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text="Поверхностное раскрытие темы")],
        [KeyboardButton(text="Полное раскрытие темы")]
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


@dp.message(Command("cancel"))
async def cmd_cancel(m: Message, state: FSMContext):
    await state.clear()
    finish_job(m.from_user.id)
    await m.answer("Отменил текущее действие. Начнём заново 👇", reply_markup=main_kb())


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
    await m.answer("Напиши тему. Можно коротко, например: киты, крипта, школа.", reply_markup=ReplyKeyboardRemove())
    await state.set_state(Form.waiting_topic)


@dp.message(Form.waiting_mode, F.text.in_(["Вставить свой текст", "📝 Вставить свой текст"]))
async def mode_user(m: Message, state: FSMContext):
    await state.update_data(mode="user")
    await m.answer("Пришли свой текст 📝\nМожно черновик — я поправлю ошибки и соберу структуру.", reply_markup=ReplyKeyboardRemove())
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
        prompt = f"""Ты — редактор презентаций высокого уровня. Пользователь прислал свой текст.
Исправь ошибки, сохрани смысл и факты, но выстрой драматургию: у презентации должна быть чёткая идея,
а не пересказ по порядку.
Текст:
{data.get('user_text')}
Слайдов: {slides}
Угол подачи: {angle}
Дай короткий план текстом: цепляющее название (не общее, а конкретное и живое) и 3 пункта,
каждый — не абстрактная категория, а конкретный тезис с сутью, который будет на слайде.
Без JSON, без вводных фраз, без нумерации "Слайд 1"."""
    else:
        prompt = f"""Ты — редактор презентаций высокого уровня, который умеет находить неочевидный
и интересный угол в любой теме, избегая шаблонных заголовков вроде "Введение" или "Что это такое".
Тема: {data.get('topic')}
Слайдов: {slides}
Доп: {data.get('extra')}
Угол подачи: {angle}
Дай короткий уникальный план текстом: цепляющее конкретное название и 3 пункта — каждый должен быть
содержательным тезисом (что именно будет рассказано), а не общей категорией.
Без JSON, без вводных фраз, без нумерации "Слайд 1"."""

    sample = await ask_grok(prompt)
    if grok_failed(sample):
        await m.answer(
            "Не получилось получить ответ от нейросети. Попробуй ещё раз через минуту.",
            reply_markup=slides_kb()
        )
        return
    theme_name = data.get("theme_name", "default")
    await state.update_data(sample=sample)
    await send_draft(
        m,
        f"Черновик готов ✅\n\n{sample}\n\n"
        f"Стиль: {THEME_LABELS.get(theme_name, theme_name)}\n\n"
        "Если всё ок — собираем полную версию.",
        title=data.get("topic", "Презентация"),
        reply_markup=confirm_kb()
    )
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(["Изменить тему", "✏️ Изменить тему"]))
async def change_topic(m: Message, state: FSMContext):
    data = await state.get_data()
    if data.get("mode") == "user":
        await m.answer("Пришли новый текст:", reply_markup=ReplyKeyboardRemove())
        await state.set_state(Form.waiting_user_text)
    else:
        await m.answer("Напиши новую тему:", reply_markup=ReplyKeyboardRemove())
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
    await m.answer("Напиши, что добавить или изменить. Максимум 800 символов.", reply_markup=ReplyKeyboardRemove())
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
    if grok_failed(sample):
        await m.answer(
            "Не получилось обновить черновик — нейросеть не ответила. Прошлый черновик остался как был.",
            reply_markup=confirm_kb()
        )
        await state.set_state(Form.waiting_confirm)
        return
    await state.update_data(sample=sample)
    await send_draft(m, f"Обновлённый черновик ✅\n\n{sample}\n\nВыбери действие:", title=data.get("topic", "Презентация"), reply_markup=confirm_kb())
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(["Делать полную версию", "✅ Делать полную версию", "делай", "да", "ок"]))
async def confirm_generate(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    if not start_job(uid):
        await m.answer("Уже собираю предыдущую версию, подожди немного 🙂")
        return
    await m.answer("Собираю презентацию с картинками. Это займёт 1–2 минуты.")
    try:

        theme_name = data.get("theme_name") or pick_theme(data.get("topic", ""))[0]
        colors = THEMES.get(theme_name, THEMES["default"])
        angle = data.get("angle") or random.choice(ANGLES)
        layouts = [0, 1, 2]
        random.shuffle(layouts)

        if data.get("mode") == "user":
            common_rules = """
    Правила качества (обязательны для каждого слайда):
    - Заголовок слайда: 3-6 слов, конкретный и живой, называет суть тезиса, а не общую категорию
      (плохо: "Введение", "Общая информация"; хорошо: конкретный факт или вывод).
    - Текст: ровно 2 абзаца, каждый 2-4 предложения. Первый абзац - конкретный факт, пример или наблюдение,
      второй - что это значит или почему это важно. Никаких вводных фраз типа "в этом слайде" или "стоит отметить".
    - Между слайдами не повторяй одну и ту же структуру фразы - разнообразь подачу (факт, вопрос, сравнение, история).
    - Никаких общих фраз без содержания ("это важная тема", "мир меняется") - только конкретика: цифры, имена,
      примеры, детали, если они есть в материале или логично следуют из темы.
    - Фото: описание живого кадра (реальная сцена, человек, объект, место), не стоковый шаблон и не абстракция."""
            raw = await ask_grok(f"""Собери уникальную презентацию из текста пользователя.
    Исправь ошибки, сохрани смысл и все факты из текста.
    Текст:
    {data.get('user_text')}
    Доп:
    {data.get('extra')}
    Слайдов: {data.get('slides')}
    Угол: {angle}
    Стиль: {theme_name}
    {common_rules}
    Только JSON:
    {{"title":"...","slides":[{{"title":"...","content":"абзац1\n\nабзац2","image_prompt":"unique cinematic scene"}}]}}""")
        else:
            common_rules = """
    Правила качества (обязательны для каждого слайда):
    - Заголовок слайда: 3-6 слов, конкретный и живой, называет суть тезиса, а не общую категорию
      (плохо: "Введение", "Общая информация"; хорошо: конкретный факт или вывод).
    - Текст: ровно 2 абзаца, каждый 2-4 предложения. Первый абзац - конкретный факт, пример или наблюдение,
      второй - что это значит или почему это важно. Никаких вводных фраз типа "в этом слайде" или "стоит отметить".
    - Между слайдами не повторяй одну и ту же структуру фразы - разнообразь подачу (факт, вопрос, сравнение, история).
    - Никаких общих фраз без содержания ("это важная тема", "мир меняется") - только конкретика: цифры, имена,
      примеры, детали, логично следующие из темы.
    - Фото: описание живого кадра (реальная сцена, человек, объект, место), не стоковый шаблон и не абстракция."""
            raw = await ask_grok(f"""Собери уникальную презентацию уровня лучшего журнала на эту тему,
    найди неочевидный и интересный угол, избегай банальностей.
    Тема: {data.get('topic')}
    Слайдов: {data.get('slides')}
    Доп: {data.get('extra')}
    Угол: {angle}
    Стиль: {theme_name}
    {common_rules}
    Только JSON:
    {{"title":"...","slides":[{{"title":"...","content":"абзац1\n\nабзац2","image_prompt":"unique cinematic scene"}}]}}""")

        try:
            content = extract_json(raw)
            if not isinstance(content.get("slides"), list) or not content["slides"]:
                raise ValueError("В ответе модели нет слайдов")
        except Exception as e:
            print("Presentation JSON parse error:", e)
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
        raw_sources = []
        for i, s in enumerate(slides_data):
            src = f"/tmp/{uid}_{i}_{random.randint(1000, 9999)}.png"
            prompt = f"{s.get('image_prompt') or s.get('title')}, {colors['photo']}, unique composition"
            ok = await generate_image(prompt, src)
            if ok:
                raw_sources.append(src)
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
        fn, fb = register_pdf_fonts()

        # Титульная "обложка" копии — как в самой презентации.
        pdf.setFont(fb, 22)
        for line in wrap_lines(content.get("title", "Презентация"), fb, 22, w - 80):
            pdf.drawString(40, h - 90, line)
            break  # заголовок короткий по промпту, одной строки почти всегда достаточно
        pdf.setFont(fn, 11)
        pdf.drawString(40, h - 115, "Текстовая копия презентации")
        pdf.showPage()

        y = h - 60
        for i, s in enumerate(slides_data, 1):
            if y < 100:
                pdf.showPage()
                y = h - 60
            pdf.setFont(fb, 14)
            for line in wrap_lines(f"{i}. {s.get('title', '')}", fb, 14, w - 80):
                if y < 60:
                    pdf.showPage()
                    y = h - 60
                    pdf.setFont(fb, 14)
                pdf.drawString(40, y, line)
                y -= 18
            y -= 6
            pdf.setFont(fn, 11)
            paragraphs = [x.strip() for x in (s.get("content") or "").split("\n") if x.strip()]
            for para in paragraphs:
                for line in wrap_lines(para, fn, 11, w - 80):
                    if y < 60:
                        pdf.showPage()
                        y = h - 60
                        pdf.setFont(fn, 11)
                    pdf.drawString(40, y, line)
                    y -= 15
                y -= 8
            y -= 14
        pdf.save()

        pres_fname = safe_filename(content.get("title"), fallback="Презентация")
        await m.answer_document(FSInputFile(pptx_path, filename=f"{pres_fname}.pptx"), caption="📊 PPTX — открывай этот файл")
        await m.answer_document(FSInputFile(pdf_path, filename=f"{pres_fname}.pdf"), caption="📄 PDF — полная текстовая копия (без фото)")
        u["generations"] += 1
        u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
        for p in (cover_src, cover_img, pptx_path, pdf_path, *raw_sources, *[f for pair in images if pair for f in pair]):
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        await m.answer(
            "Готово ✅\n\n"
            "Открывай именно PPTX в PowerPoint, Keynote или Google Презентациях.\n\n"
            "Если на телефоне все фото одинаковые, это не ошибка файла. "
            "Так бывает в предпросмотре Telegram, WPS и встроенных «Документах». "
            "Открой тот же файл на другом устройстве или в нормальном редакторе презентаций.",
            reply_markup=main_kb()
        )
        await state.clear()
    finally:
        finish_job(uid)


@dp.message(Form.waiting_confirm)
async def waiting_confirm_fallback(m: Message, state: FSMContext):
    """Ловит любой текст, не совпавший с кнопками выше, чтобы диалог никогда
    не зависал без ответа."""
    await m.answer(
        "Не понял. Выбери действие кнопкой ниже, или напиши /cancel, чтобы начать заново.",
        reply_markup=confirm_kb()
    )


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


def _body(doc, sections, indent=True, head_center=False, size=12, line=1.15, first=1.25, head_size=13):
    for block in sections or []:
        head = (block.get("title") or "").strip()
        body = block.get("content") or ""
        if head:
            _p(
                doc, head, head_size, True,
                align=WD_ALIGN_PARAGRAPH.CENTER if head_center else WD_ALIGN_PARAGRAPH.LEFT,
                before=10, after=6
            )
        for para in [x.strip() for x in body.split("\n") if x.strip()]:
            _p(doc, para, size, first=first if indent else None, after=6, line=line)


# Семейства новых юридических документов: чтобы не дублировать почти одинаковый
# код форматирования 22 раза, новые виды сгруппированы по "семьям" (двусторонний
# договор / акт / заявление в суд-организацию / односторонний документ), и
# build_word() рендерит их по общему шаблону, подставляя только заголовок и роли сторон.
CONTRACT_FAMILY = {
    "marriage_contract": {"caption": "БРАЧНЫЙ ДОГОВОР", "left": "СУПРУГ", "right": "СУПРУГА"},
    "gift": {"caption": "ДОГОВОР ДАРЕНИЯ", "left": "ДАРИТЕЛЬ", "right": "ОДАРЯЕМЫЙ"},
    "supply": {"caption": "ДОГОВОР ПОСТАВКИ", "left": "ПОСТАВЩИК", "right": "ПОКУПАТЕЛЬ"},
    "agency": {"caption": "АГЕНТСКИЙ ДОГОВОР", "left": "ПРИНЦИПАЛ", "right": "АГЕНТ"},
    "joint_activity": {"caption": "ДОГОВОР О СОВМЕСТНОЙ ДЕЯТЕЛЬНОСТИ", "left": "СТОРОНА 1", "right": "СТОРОНА 2"},
    "nonresidential_rent": {"caption": "ДОГОВОР АРЕНДЫ НЕЖИЛОГО ПОМЕЩЕНИЯ", "left": "АРЕНДОДАТЕЛЬ", "right": "АРЕНДАТОР"},
    "cession": {"caption": "ДОГОВОР ЦЕССИИ (УСТУПКИ ПРАВА ТРЕБОВАНИЯ)", "left": "ЦЕДЕНТ", "right": "ЦЕССИОНАРИЙ"},
    "nda": {"caption": "СОГЛАШЕНИЕ О НЕРАЗГЛАШЕНИИ (NDA)", "left": "СТОРОНА 1", "right": "СТОРОНА 2"},
    "employment": {"caption": "ТРУДОВОЙ ДОГОВОР", "left": "РАБОТОДАТЕЛЬ", "right": "РАБОТНИК"},
    "self_employed": {"caption": "ДОГОВОР С САМОЗАНЯТЫМ", "left": "ЗАКАЗЧИК", "right": "ИСПОЛНИТЕЛЬ"},
    "services": {"caption": "ДОГОВОР ОКАЗАНИЯ УСЛУГ", "left": "ЗАКАЗЧИК", "right": "ИСПОЛНИТЕЛЬ"},
}

ACT_FAMILY = {
    "work_act": {"caption": "АКТ ВЫПОЛНЕННЫХ РАБОТ (ОКАЗАННЫХ УСЛУГ)"},
}

STATEMENT_FAMILY = {
    "lawsuit": {
        "caption": "ИСКОВОЕ ЗАЯВЛЕНИЕ",
        "closing": "На основании изложенного и в соответствии с законодательством РФ прошу суд удовлетворить исковые требования.",
    },
}

DECLARATION_FAMILY = {
    "alimony": {
        "caption": "СОГЛАШЕНИЕ ОБ УПЛАТЕ АЛИМЕНТОВ",
        "intro": "Я, {author}, обязуюсь уплачивать алименты на условиях, указанных ниже:",
    },
    "warranty_letter": {
        "caption": "ГАРАНТИЙНОЕ ПИСЬМО",
        "intro": "Настоящим письмом {author} гарантирует следующее:",
    },
}

# Единая схема мета-полей под каждый тип документа. Используется и в промпте для ИИ
# (как образец JSON, который нужно заполнить), и напрямую как список плейсхолдеров
# для режима "Скачать шаблон" (без обращения к ИИ).
META_SCHEMAS = {
    "referat": '{"org":"Министерство образования","school":"[учебное заведение]","author":"[ФИО студента]","teacher":"[ФИО преподавателя]","city":"[город]","year":"2026"}',
    "coursework": '{"org":"Министерство образования","school":"[учебное заведение]","specialty":"[специальность]","author":"[ФИО студента]","teacher":"[ФИО научного руководителя]","city":"[город]","year":"2026"}',
    "report": '{"school":"[учебное заведение]","author":"[ФИО, класс/курс]"}',
    "essay": '{"school":"[учебное заведение]","author":"[ФИО]"}',
    "notes": '{}',
    "dkp": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "rent": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "offer": '{"number":"[номер КП]"}',
    "act": '{"city":"[город]","date":"«___» __________ 20___ г.","basis":"[договор №/основание]","from_party":"[передающая сторона, ФИО/наименование]","to_party":"[принимающая сторона, ФИО/наименование]"}',
    "statement": '{"to":"[должность руководителя]","to_name":"[ФИО руководителя]","from":"[должность, ФИО заявителя]","date":"[дата]","author":"[ФИО]"}',
    "proxy": '{"city":"[город]","date":"«___» __________ 20___ г.","author":"[ФИО доверителя]"}',
    "loan": '{"city":"[город]","date":"«___» __________ 20___ г.","from_party":"[заимодавец, ФИО/паспорт]","to_party":"[заёмщик, ФИО/паспорт]"}',
    "claim": '{"to":"[адресат претензии]","from":"[заявитель, ФИО/контакты]","date":"[дата]","basis":"[договор №/основание]"}',
    "consent": '{"city":"[город]","date":"«___» __________ 20___ г.","author":"[ФИО родителя]"}',
    "marriage_contract": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "gift": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "lawsuit": '{"to":"[наименование суда]","from":"[ФИО истца, адрес, контакты]","date":"[дата]"}',
    "alimony": '{"city":"[город]","date":"«___» __________ 20___ г.","author":"[ФИО плательщика]"}',
    "services": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "employment": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "work_act": '{"city":"[город]","date":"«___» __________ 20___ г.","basis":"[договор №/основание]","from_party":"[исполнитель]","to_party":"[заказчик]"}',
    "supply": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "agency": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "joint_activity": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "nonresidential_rent": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "cession": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "nda": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "self_employed": '{"city":"[город]","date":"«___» __________ 20___ г."}',
    "warranty_letter": '{"city":"[город]","date":"«___» __________ 20___ г.","author":"[наименование/ФИО гаранта]"}',
    "doc": '{"org":"","sign":""}',
}

# Описание вида документа для промпта ИИ (что именно должно быть в разделах).
WORD_KIND_DESC = {
    "doc": "обычный документ",
    "referat": "реферат для университета. Титульный лист, содержание, введение (актуальность темы, цель и задачи), 2–3 главы, заключение с выводами (не пересказ содержания), список литературы не менее 5 источников. Текст живой, связный, как для сдачи преподавателю, не набор абзацев.",
    "report": "доклад для школы/университета. Титульный лист (упрощённый), краткое вступление, основная часть по теме с конкретными фактами, короткий вывод. Компактно и по существу, рассчитан на устное зачитывание.",
    "essay": "эссе. Без титульного листа или с упрощённым. Личная позиция автора, аргументированное рассуждение с примерами, живой связный текст, не сухой пересказ.",
    "notes": "конспект. Без титульного листа. Сжатое изложение материала по пунктам и подпунктам, ключевые определения выделены, без художественных отступлений — удобно для повторения перед экзаменом.",
    "coursework": "курсовая работа. Титульный лист, содержание, введение (актуальность темы, объект и предмет, цель и задачи исследования), 2-3 главы с подразделами (например 1.1, 1.2 и 2.1, 2.2) — теоретическая и практическая часть, заключение с самостоятельными выводами по каждой задаче (не пересказ содержания глав), список литературы не менее 5 источников. Это черновик-каркас под доработку с научным руководителем, а не финальная работа.",
    "dkp": "договор купли-продажи: предмет договора, цена и порядок расчётов, права и обязанности сторон, порядок передачи товара, ответственность сторон, срок действия и заключительные положения",
    "rent": "договор аренды: предмет аренды с точным описанием, срок аренды, размер и порядок внесения арендной платы, права и обязанности сторон, порядок передачи и возврата имущества, ответственность сторон",
    "offer": "коммерческое предложение уровня 2026 года: заголовок с конкретной выгодой или болью клиента (не 'КП от компании'), короткое описание сути в 1-2 абзаца, о компании, что именно входит в предложение, стоимость и условия оплаты, сроки и этапы, почему стоит выбрать именно нас (доказательства, кейсы), контакты и призыв к действию с дедлайном",
    "act": "акт приёма-передачи: дата и место составления, номер и основание (реквизиты договора), полные данные передающей и принимающей стороны, подробный перечень передаваемого имущества/товара/документов с количеством и состоянием, отметка об отсутствии претензий, место для подписей обеих сторон",
    "statement": "заявление",
    "proxy": "доверенность: кто выдаёт (доверитель), кому (представитель), точный перечень полномочий, срок действия",
    "loan": "расписка / договор займа: заимодавец и заёмщик (ФИО, паспортные данные), сумма займа цифрами и прописью, срок возврата, проценты (если есть) или указание на беспроцентный заём, порядок возврата, ответственность за просрочку (неустойка/пени)",
    "claim": "досудебная претензия: реквизиты адресата и заявителя, описание нарушения со ссылкой на договор/закон, конкретное требование (сумма, срок исполнения), срок для добровольного удовлетворения, предупреждение об обращении в суд",
    "consent": "согласие на выезд ребёнка за границу: ФИО и паспортные данные родителя (доверителя), ФИО, дата рождения и данные свидетельства о рождении/паспорта ребёнка, ФИО сопровождающего (если есть), страна/страны выезда, срок действия согласия",
    "marriage_contract": "брачный договор: ФИО супругов, реквизиты свидетельства о браке, режим собственности на имущество (совместное/раздельное/долевое), перечень конкретного имущества и порядок его раздела в случае развода",
    "gift": "договор дарения: ФИО и паспортные данные дарителя и одаряемого, точное описание предмета дарения, отсутствие встречного предоставления, момент перехода права собственности",
    "lawsuit": "исковое заявление в суд: наименование суда, данные истца и ответчика, цена иска (если применимо), обстоятельства дела по порядку, правовое обоснование со ссылками на закон, исковые требования, перечень приложений",
    "alimony": "соглашение об уплате алиментов: ФИО плательщика и получателя, ФИО и дата рождения ребёнка, размер алиментов (доля дохода или твёрдая сумма), периодичность и способ выплаты, индексация",
    "services": "договор возмездного оказания услуг: заказчик и исполнитель, точное описание услуги, стоимость и порядок оплаты, сроки оказания, порядок сдачи-приёмки, ответственность сторон",
    "employment": "трудовой договор: работодатель и работник, должность и трудовая функция, дата начала работы, режим рабочего времени, оклад и порядок выплаты зарплаты, права и обязанности сторон, испытательный срок (если есть)",
    "work_act": "акт выполненных работ (оказанных услуг): реквизиты основного договора, заказчик и исполнитель, перечень выполненных работ/услуг с объёмом, стоимостью и сроками, отметка о соответствии качества, отсутствие претензий",
    "supply": "договор поставки: поставщик и покупатель, наименование, ассортимент и количество товара, цена, сроки и порядок поставки, порядок приёмки товара по количеству и качеству, ответственность за просрочку",
    "agency": "агентский договор: принципал и агент, точный перечень юридических и фактических действий, которые агент совершает от имени и за счёт принципала, размер и порядок выплаты агентского вознаграждения, срок действия",
    "joint_activity": "договор о совместной деятельности (простого товарищества): участники, общая цель, вклад каждого участника (деньги, имущество, навыки), порядок ведения общих дел, распределение прибыли и убытков",
    "nonresidential_rent": "договор аренды нежилого помещения: арендодатель и арендатор, точный адрес и площадь помещения, назначение использования, срок аренды, размер и порядок внесения арендной платы, порядок передачи и возврата помещения",
    "cession": "договор цессии (уступки права требования): цедент и цессионарий, реквизиты первоначального обязательства/договора, объём и содержание уступаемого права требования, цена уступки, момент перехода права",
    "nda": "соглашение о неразглашении (NDA): стороны, точное определение конфиденциальной информации, цель передачи информации, обязательства по неразглашению и ограничению доступа, срок действия обязательств, ответственность за нарушение",
    "self_employed": "договор с самозанятым (на выполнение работ/оказание услуг): заказчик и исполнитель-самозанятый (ФИО, ИНН, номер чека НПД), предмет договора, стоимость, сроки, порядок сдачи-приёмки и оплаты",
    "warranty_letter": "гарантийное письмо: кому адресовано, кто гарантирует, чёткая суть гарантии (оплата, выполнение обязательства, устранение недостатков), срок исполнения гарантии",
}


def template_sections_for(kind):
    """Плейсхолдер-структура документа для режима 'Скачать шаблон' (без ИИ)."""
    if kind in ("dkp", "rent") or kind in CONTRACT_FAMILY:
        return [
            {"title": "Предмет договора", "content": "[точное описание предмета договора]"},
            {"title": "Цена и порядок расчётов", "content": "[сумма, порядок и сроки оплаты]"},
            {"title": "Права и обязанности сторон", "content": "[права и обязанности каждой стороны]"},
            {"title": "Ответственность сторон", "content": "[ответственность за нарушение условий]"},
            {"title": "Заключительные положения", "content": "[срок действия, порядок разрешения споров, реквизиты сторон]"},
        ]
    if kind == "act" or kind in ACT_FAMILY:
        return [{"title": "", "content": "[перечень передаваемого/выполненного — наименование, количество, состояние, стоимость]"}]
    if kind in ("statement", "claim") or kind in STATEMENT_FAMILY:
        return [{"title": "", "content": "Прошу [указать суть просьбы, требования или обстоятельства дела]."}]
    if kind in ("proxy", "consent") or kind in DECLARATION_FAMILY:
        return [{"title": "", "content": "[точный перечень полномочий или условий]"}]
    if kind == "loan":
        return [{"title": "", "content": "Сумма займа: [сумма цифрами и прописью]. Срок возврата: [дата]. Проценты: [указать или «беспроцентный»]."}]
    if kind == "offer":
        return [
            {"title": "Суть предложения", "content": "[что именно предлагается]"},
            {"title": "Стоимость и условия", "content": "[цена, порядок оплаты, сроки]"},
        ]
    return [{"title": "", "content": "[текст документа]"}]


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
        # Поля, шрифт, интервал и отступ ниже соответствуют официальным требованиям
        # вузов к оформлению учебных работ (Times New Roman 14pt, интервал 1.2,
        # поля 20мм со всех сторон, абзацный отступ 1см).
        sec.left_margin = Cm(2)
        sec.right_margin = Cm(2)
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
        _body(doc, sections, indent=True, head_center=True, size=14, line=1.2, first=1.0, head_size=14)

    elif kind == "coursework":
        sec.left_margin = Cm(2)
        sec.right_margin = Cm(2)
        _p(doc, meta.get("org") or "Министерство образования", 14, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
        _p(doc, meta.get("school") or "[Название учебного заведения]", 14, align=WD_ALIGN_PARAGRAPH.CENTER)
        if meta.get("specialty"):
            _p(doc, f"Специальность: {meta.get('specialty')}", 12, align=WD_ALIGN_PARAGRAPH.CENTER)
        for _ in range(4):
            doc.add_paragraph()
        _p(doc, "КУРСОВАЯ РАБОТА", 20, True, WD_ALIGN_PARAGRAPH.CENTER, after=8)
        _p(doc, title, 16, True, WD_ALIGN_PARAGRAPH.CENTER)
        for _ in range(6):
            doc.add_paragraph()
        _p(doc, f"Выполнил: {meta.get('author') or '[ФИО студента]'}", 14, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0)
        _p(doc, f"Научный руководитель: {meta.get('teacher') or '[ФИО преподавателя]'}", 14, align=WD_ALIGN_PARAGRAPH.RIGHT)
        for _ in range(6):
            doc.add_paragraph()
        _p(doc, f"{meta.get('city') or '[Город]'} {meta.get('year') or '2026'}", 14, align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.add_page_break()
        _body(doc, sections, indent=True, head_center=True, size=14, line=1.2, first=1.0, head_size=14)

    elif kind == "report":
        if meta.get("school") or meta.get("author"):
            _p(doc, meta.get("school") or "", 11, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
            _p(doc, f"Выполнил: {meta.get('author')}" if meta.get("author") else "", 11, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        _p(doc, "ДОКЛАД", 12, True, WD_ALIGN_PARAGRAPH.CENTER, after=0)
        _p(doc, title, 18, True, WD_ALIGN_PARAGRAPH.CENTER, before=4, after=14)
        _body(doc, sections, indent=True, head_center=False)

    elif kind == "essay":
        if meta.get("school") or meta.get("author"):
            _p(doc, meta.get("school") or "", 11, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
            _p(doc, f"{meta.get('author')}" if meta.get("author") else "", 11, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        _p(doc, "ЭССЕ", 12, True, WD_ALIGN_PARAGRAPH.CENTER, after=0)
        _p(doc, title, 18, True, WD_ALIGN_PARAGRAPH.CENTER, before=4, after=14)
        _body(doc, sections, indent=True, head_center=False)

    elif kind == "notes":
        _p(doc, "КОНСПЕКТ", 12, True, WD_ALIGN_PARAGRAPH.CENTER, after=0)
        _p(doc, title, 18, True, WD_ALIGN_PARAGRAPH.CENTER, before=4, after=14)
        _body(doc, sections, indent=False, head_center=False)

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
        _p(doc, title, 20, True, after=14)
        for i, block in enumerate(sections or [], 1):
            head = (block.get("title") or "").strip()
            body = block.get("content") or ""
            if head:
                _p(doc, f"{i}. {head}", 14, True, before=12, after=6)
            for para in [x.strip() for x in body.split("\n") if x.strip()]:
                _p(doc, para, 12, after=6, line=1.2)

    elif kind == "act":
        _p(doc, "АКТ ПРИЁМА-ПЕРЕДАЧИ", 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, title if title not in ("АКТ ПРИЁМА-ПЕРЕДАЧИ", "Документ") else "№ ______", 12,
           align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        if meta.get("basis"):
            _p(doc, f"Основание: {meta.get('basis')}", 12, after=10)
        _p(
            doc,
            f"{meta.get('from_party') or '[передающая сторона, ФИО/наименование]'}, именуемый(ая) в дальнейшем "
            "«Передающая сторона», с одной стороны, и "
            f"{meta.get('to_party') or '[принимающая сторона, ФИО/наименование]'}, именуемый(ая) в дальнейшем "
            "«Принимающая сторона», с другой стороны, составили настоящий акт о нижеследующем:",
            12, after=10, first=1.25
        )
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, "Претензий по количеству, качеству и комплектности стороны друг к другу не имеют.", 12, before=8, after=18)
        _p(doc, "ПЕРЕДАЛ                    ПРИНЯЛ", 12, True, before=6)
        _p(doc, "__________ / [ФИО] /          __________ / [ФИО] /", 12)

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

    elif kind == "loan":
        _p(doc, "РАСПИСКА В ПОЛУЧЕНИИ ДЕНЕЖНЫХ СРЕДСТВ", 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, "(договор займа)", 11, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        _p(
            doc,
            f"{meta.get('from_party') or '[Заимодавец, ФИО, паспортные данные]'}, именуемый(ая) в дальнейшем "
            "«Заимодавец», передал(а) в долг, а "
            f"{meta.get('to_party') or '[Заёмщик, ФИО, паспортные данные]'}, именуемый(ая) в дальнейшем "
            "«Заёмщик», получил(а) в долг денежные средства на нижеследующих условиях:",
            12, after=10, first=1.25
        )
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, "ЗАИМОДАВЕЦ                    ЗАЁМЩИК", 12, True, before=18)
        _p(doc, "__________ / [ФИО] /          __________ / [ФИО] /", 12)

    elif kind == "claim":
        _p(doc, meta.get("to") or "[адресат претензии]", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0)
        _p(doc, f"от {meta.get('from') or '[ФИО, контакты заявителя]'}", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=18)
        _p(doc, "ПРЕТЕНЗИЯ", 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=4)
        if meta.get("basis"):
            _p(doc, f"Основание: {meta.get('basis')}", 11, align=WD_ALIGN_PARAGRAPH.CENTER, after=12)
        _body(doc, sections, indent=True, head_center=False)
        _p(
            doc,
            "В случае неудовлетворения настоящей претензии в указанный срок я буду вынужден(а) "
            "обратиться в суд за защитой своих прав со взысканием всех сопутствующих расходов.",
            12, before=10, after=18
        )
        _p(doc, f"{meta.get('date') or '[дата]'}                    __________ / {meta.get('from') or '[ФИО]'} /", 12)

    elif kind == "consent":
        _p(doc, "СОГЛАСИЕ", 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, "на выезд несовершеннолетнего ребёнка за границу", 12, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        _p(
            doc,
            f"Я, {meta.get('author') or '[ФИО родителя, паспортные данные]'}, даю согласие на выезд "
            "моего несовершеннолетнего ребёнка за пределы Российской Федерации на условиях, указанных ниже:",
            12, after=10, first=1.25
        )
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, f"Подпись: __________ / {meta.get('author') or '[ФИО]'} /", 12, before=20)

    elif kind in CONTRACT_FAMILY:
        fam = CONTRACT_FAMILY[kind]
        _p(doc, fam["caption"], 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, title if title not in (fam["caption"], "Документ") else "№ ______", 12,
           align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, f"{fam['left']}                    {fam['right']}", 12, True, before=18)
        _p(doc, "__________ / [ФИО/наименование] /          __________ / [ФИО/наименование] /", 12)

    elif kind in ACT_FAMILY:
        fam = ACT_FAMILY[kind]
        _p(doc, fam["caption"], 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=2)
        _p(doc, title if title not in (fam["caption"], "Документ") else "№ ______", 12,
           align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        if meta.get("basis"):
            _p(doc, f"Основание: {meta.get('basis')}", 12, after=10)
        _p(
            doc,
            f"{meta.get('from_party') or '[исполнитель, ФИО/наименование]'}, именуемый(ая) в дальнейшем "
            "«Исполнитель», с одной стороны, и "
            f"{meta.get('to_party') or '[заказчик, ФИО/наименование]'}, именуемый(ая) в дальнейшем "
            "«Заказчик», с другой стороны, составили настоящий акт о нижеследующем:",
            12, after=10, first=1.25
        )
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, "Претензий по объёму, качеству и срокам стороны друг к другу не имеют.", 12, before=8, after=18)
        _p(doc, "ИСПОЛНИТЕЛЬ                    ЗАКАЗЧИК", 12, True, before=6)
        _p(doc, "__________ / [ФИО] /          __________ / [ФИО] /", 12)

    elif kind in STATEMENT_FAMILY:
        fam = STATEMENT_FAMILY[kind]
        _p(doc, meta.get("to") or "[наименование суда/адресата]", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=0)
        _p(doc, f"от {meta.get('from') or '[ФИО, адрес, контакты]'}", 12, align=WD_ALIGN_PARAGRAPH.RIGHT, after=18)
        _p(doc, fam["caption"], 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=4)
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, fam["closing"], 12, before=10, after=18)
        _p(doc, f"{meta.get('date') or '[дата]'}                    __________ / {meta.get('from') or '[ФИО]'} /", 12)

    elif kind in DECLARATION_FAMILY:
        fam = DECLARATION_FAMILY[kind]
        _p(doc, fam["caption"], 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=10)
        p = doc.add_paragraph()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT)
        _run(p, f"{meta.get('city') or 'г. _______________'}\t{meta.get('date') or '«___» __________ 20___ г.'}", 12)
        _p(doc, fam["intro"].format(author=meta.get("author") or "[ФИО]"), 12, after=10, first=1.25)
        _body(doc, sections, indent=True, head_center=False)
        _p(doc, f"Подпись: __________ / {meta.get('author') or '[ФИО]'} /", 12, before=20)

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
    await m.answer("Для кого документ?", reply_markup=word_category_kb())
    await state.set_state(Form.waiting_word_category)


@dp.message(Form.waiting_word_category)
async def word_category(m: Message, state: FSMContext):
    t = (m.text or "").lower()
    cat = None
    if "физ" in t:
        cat = "physical"
    elif "юр" in t or "ип" in t:
        cat = "entity"
    elif "учеб" in t or "учёб" in t:
        cat = "study"
    if not cat:
        await m.answer("Выбери категорию кнопкой ниже.", reply_markup=word_category_kb())
        return
    await state.update_data(word_category=cat)
    await m.answer("Какой документ нужен?", reply_markup=word_kind_kb(cat))
    await state.set_state(Form.waiting_word_kind)


@dp.message(Form.waiting_word_kind, F.text == BTN_MORE_DOCS)
async def word_kind_more(m: Message, state: FSMContext):
    data = await state.get_data()
    cat = data.get("word_category", "physical")
    await m.answer("Ещё документы:", reply_markup=word_kind_kb(cat, more=True))


@dp.message(Form.waiting_word_kind, F.text == BTN_BACK_TO_CATEGORIES)
async def word_kind_back(m: Message, state: FSMContext):
    await m.answer("Для кого документ?", reply_markup=word_category_kb())
    await state.set_state(Form.waiting_word_category)


@dp.message(Form.waiting_word_kind)
async def word_kind(m: Message, state: FSMContext):
    kind = LABEL_TO_KIND.get((m.text or "").strip())
    if not kind:
        data = await state.get_data()
        cat = data.get("word_category", "physical")
        await m.answer("Выбери документ кнопкой ниже.", reply_markup=word_kind_kb(cat))
        return
    await state.update_data(word_kind=kind, extra="", extra_used=0)
    show_template = kind not in STUDY_KINDS
    hint = (
        "Как собираем документ?\n\n"
        "✨ Сгенерировать с ИИ — я сам напишу текст.\n"
        "📝 Вставить свой текст — пришли материал, я поправлю и оформлю."
    )
    if show_template:
        hint += "\n📄 Скачать шаблон — пустой бланк с пропусками для заполнения от руки."
    await m.answer(hint, reply_markup=mode_kb(show_template))
    await state.set_state(Form.waiting_word_mode)



WORD_KIND_HINTS = {
    "doc": {
        "ai": "Напиши тему документа. Например: доклад про историю китов.",
        "user": "Пришли текст документа — я поправлю грамотность и оформлю по стандарту.",
    },
    "referat": {
        "ai": "Напиши тему реферата. Также укажи, если знаешь: учебное заведение, свои ФИО, ФИО преподавателя, город и год — иначе оставлю поля пустыми для заполнения.",
        "user": "Пришли текст реферата (или тезисы). Также укажи учебное заведение, свои ФИО, ФИО преподавателя, город и год, если нужно их вставить.",
    },
    "report": {
        "ai": "Напиши тему доклада. Также укажи, если знаешь: учебное заведение, свои ФИО, класс/курс — иначе оставлю поля пустыми для заполнения.",
        "user": "Пришли текст доклада (или тезисы). Также укажи учебное заведение, свои ФИО и класс/курс, если нужно их вставить.",
    },
    "essay": {
        "ai": "Напиши тему эссе и свою позицию по ней, если она уже есть. Укажи, если знаешь: учебное заведение, свои ФИО.",
        "user": "Пришли текст эссе (или тезисы, свои мысли по теме). Укажи учебное заведение и свои ФИО, если нужно их вставить.",
    },
    "notes": {
        "ai": "Напиши тему или раздел, по которому нужен конспект (например: конспект лекции по клеточной биологии).",
        "user": "Пришли материал, по которому нужно сделать конспект — я структурирую его в сжатом виде по пунктам.",
    },
    "coursework": {
        "ai": "Напиши тему курсовой работы. Укажи, если знаешь: учебное заведение, специальность, свои ФИО, ФИО научного руководителя, город и год.",
        "user": "Пришли текст или тезисы курсовой работы. Укажи учебное заведение, специальность, свои ФИО, ФИО научного руководителя, город и год, если нужно их вставить.",
    },
    "dkp": {
        "ai": "Напиши, что покупается/продаётся (например: продажа автомобиля). Укажи, если знаешь: ФИО и паспортные данные продавца и покупателя, точное описание вещи, цену, порядок оплаты, дату и место.",
        "user": "Пришли данные для договора купли-продажи: ФИО и паспортные данные продавца и покупателя, описание предмета продажи, цену, порядок оплаты, дату и место составления.",
    },
    "rent": {
        "ai": "Напиши, что сдаётся в аренду (например: аренда квартиры). Укажи, если знаешь: данные сторон, точный адрес/описание объекта, срок аренды, сумму и порядок оплаты.",
        "user": "Пришли данные для договора аренды: ФИО сторон, точный адрес или описание объекта аренды, срок, сумму и порядок оплаты, кто оплачивает коммунальные услуги/ремонт.",
    },
    "offer": {
        "ai": "Напиши, что за услугу/товар предлагаете и кому. Укажи, если знаешь: название компании, суть предложения, цену и условия, срок действия КП.",
        "user": "Пришли данные для коммерческого предложения: название компании, суть предложения, цену и условия оплаты, сроки, контакты.",
    },
    "act": {
        "ai": "Напиши, что передаётся и по какому договору (например: акт передачи оборудования по договору №12). Укажи, если знаешь: кто передаёт, кто принимает, перечень предметов.",
        "user": "Пришли данные для акта приёма-передачи: номер и дату основного договора, кто передаёт и кто принимает (ФИО/организация), перечень передаваемого с количеством и состоянием.",
    },
    "statement": {
        "ai": "Напиши суть заявления и кому оно адресовано.",
        "user": "Пришли текст заявления: кому адресовано (должность, ФИО), от кого, суть просьбы или требования, дата.",
    },
    "proxy": {
        "ai": "Напиши, на какие действия нужна доверенность. Укажи, если знаешь: ФИО и паспортные данные доверителя и представителя, срок действия.",
        "user": "Пришли данные для доверенности: ФИО и паспортные данные доверителя и представителя, точный перечень полномочий, срок действия, дату составления.",
    },
    "loan": {
        "ai": "Напиши сумму и условия займа. Укажи, если знаешь: ФИО сторон, срок возврата, проценты.",
        "user": "Пришли данные для расписки/договора займа: ФИО и паспортные данные заимодавца и заёмщика, сумму, срок возврата, проценты (если есть).",
    },
    "claim": {
        "ai": "Напиши суть претензии и к кому она адресована. Укажи, если знаешь: от кого претензия, основание (договор/факт), требование.",
        "user": "Пришли данные для претензии: кому адресована (ФИО/организация), от кого, суть нарушения, конкретное требование, срок ответа.",
    },
    "consent": {
        "ai": "Напиши, кто едет и куда (например: согласие на выезд сына в Турцию). Укажи, если знаешь: ФИО и данные родителя, ФИО и дату рождения ребёнка, с кем и куда он выезжает, срок действия согласия.",
        "user": "Пришли данные для согласия на выезд ребёнка: ФИО родителя (доверителя), ФИО и дату рождения ребёнка, с кем именно и в какую страну выезжает ребёнок, на какой срок.",
    },
    "marriage_contract": {
        "ai": "Напиши, какие имущественные вопросы супруги хотят закрепить в брачном договоре. Укажи, если знаешь: ФИО супругов, режим имущества (совместное/раздельное), конкретное имущество.",
        "user": "Пришли данные для брачного договора: ФИО обоих супругов, реквизиты свидетельства о браке, какое имущество и на каких условиях делится.",
    },
    "gift": {
        "ai": "Напиши, что и кому дарится (например: дарение квартиры сыну). Укажи, если знаешь: ФИО дарителя и одаряемого, точное описание предмета дарения.",
        "user": "Пришли данные для договора дарения: ФИО и паспортные данные дарителя и одаряемого, точное описание предмета дарения, документы-основания (если есть).",
    },
    "lawsuit": {
        "ai": "Напиши суть спора и в какой суд обращаешься. Укажи, если знаешь: ФИО истца и ответчика, обстоятельства дела, какие требования заявляешь.",
        "user": "Пришли данные для искового заявления: наименование суда, ФИО/данные истца и ответчика, обстоятельства дела, исковые требования, цену иска (если есть).",
    },
    "alimony": {
        "ai": "Напиши, кто и на кого платит алименты, и на каких условиях (сумма, периодичность).",
        "user": "Пришли данные для соглашения об алиментах: ФИО плательщика и получателя, ФИО и дату рождения ребёнка, размер и периодичность выплат.",
    },
    "services": {
        "ai": "Напиши, какую услугу и кто оказывает. Укажи, если знаешь: заказчика и исполнителя, суть услуги, стоимость и сроки.",
        "user": "Пришли данные для договора оказания услуг: заказчик и исполнитель, точное описание услуги, стоимость, порядок оплаты, сроки оказания.",
    },
    "employment": {
        "ai": "Напиши должность и условия работы. Укажи, если знаешь: работодателя, работника, оклад, дату начала работы, режим работы.",
        "user": "Пришли данные для трудового договора: работодатель и работник (ФИО, паспортные данные), должность, оклад, режим работы, дату начала работы.",
    },
    "work_act": {
        "ai": "Напиши, какие работы или услуги выполнены и по какому договору. Укажи, если знаешь: заказчика и исполнителя, перечень работ, стоимость.",
        "user": "Пришли данные для акта выполненных работ: номер и дату договора, заказчик и исполнитель, перечень выполненных работ/услуг с объёмом и стоимостью.",
    },
    "supply": {
        "ai": "Напиши, какой товар поставляется и кем. Укажи, если знаешь: поставщика и покупателя, товар, количество, цену, сроки поставки.",
        "user": "Пришли данные для договора поставки: поставщик и покупатель, наименование и количество товара, цена, сроки и порядок поставки и оплаты.",
    },
    "agency": {
        "ai": "Напиши, какие действия агент совершает в интересах принципала. Укажи, если знаешь: стороны, суть поручения, вознаграждение агента.",
        "user": "Пришли данные для агентского договора: принципал и агент, точное описание поручаемых действий, размер и порядок выплаты вознаграждения.",
    },
    "joint_activity": {
        "ai": "Напиши цель совместной деятельности и вклад каждой стороны.",
        "user": "Пришли данные для договора о совместной деятельности: стороны, цель, вклад каждого участника, порядок распределения прибыли и расходов.",
    },
    "nonresidential_rent": {
        "ai": "Напиши, какое нежилое помещение сдаётся в аренду. Укажи, если знаешь: стороны, адрес и площадь, срок аренды, сумму.",
        "user": "Пришли данные для аренды нежилого помещения: арендодатель и арендатор, точный адрес и площадь, срок аренды, размер и порядок оплаты.",
    },
    "cession": {
        "ai": "Напиши, какое право требования и по какому обязательству уступается.",
        "user": "Пришли данные для договора цессии: цедент и цессионарий, реквизиты первоначального обязательства, сумма и объём уступаемого права, цена уступки.",
    },
    "nda": {
        "ai": "Напиши, какую конфиденциальную информацию нужно защитить и между кем.",
        "user": "Пришли данные для соглашения о неразглашении: стороны, что считается конфиденциальной информацией, срок действия обязательств, ответственность за разглашение.",
    },
    "self_employed": {
        "ai": "Напиши, какую работу выполняет самозанятый и для кого.",
        "user": "Пришли данные для договора с самозанятым: заказчик и исполнитель (ФИО, ИНН самозанятого), суть работ, стоимость, сроки, порядок оплаты.",
    },
    "warranty_letter": {
        "ai": "Напиши, что именно гарантируется и кому адресовано письмо.",
        "user": "Пришли данные для гарантийного письма: кому адресовано, кто гарантирует, суть гарантии, срок исполнения.",
    },
}


def word_hint(kind: str, mode: str) -> str:
    return WORD_KIND_HINTS.get(kind, WORD_KIND_HINTS["doc"])[mode]


@dp.message(Form.waiting_word_mode, F.text.in_(["Сгенерировать с ИИ", "✨ Сгенерировать с ИИ"]))
async def word_mode_ai(m: Message, state: FSMContext):
    await state.update_data(mode="ai", user_text="")
    data = await state.get_data()
    await m.answer(word_hint(data.get("word_kind", "doc"), "ai"), reply_markup=ReplyKeyboardRemove())
    await state.set_state(Form.waiting_word_topic)


@dp.message(Form.waiting_word_mode, F.text.in_(["Вставить свой текст", "📝 Вставить свой текст"]))
async def word_mode_user(m: Message, state: FSMContext):
    await state.update_data(mode="user")
    data = await state.get_data()
    await m.answer(word_hint(data.get("word_kind", "doc"), "user"), reply_markup=ReplyKeyboardRemove())
    await state.set_state(Form.waiting_word_text)


@dp.message(Form.waiting_word_mode, F.text.in_(["Скачать шаблон", "📄 Скачать шаблон"]))
async def word_mode_template(m: Message, state: FSMContext):
    data = await state.get_data()
    kind = data.get("word_kind", "doc")
    uid = m.from_user.id
    if kind in STUDY_KINDS:
        await m.answer("Для учебных документов шаблон не предусмотрен.", reply_markup=mode_kb(False))
        return
    if not start_job(uid):
        await m.answer("Уже собираю документ, подожди немного 🙂")
        return
    await m.answer("Собираю шаблон…", reply_markup=ReplyKeyboardRemove())
    try:
        u = get_user(uid)
        tpl_path = f"/tmp/tpl_{uid}.docx"
        try:
            meta = json.loads(META_SCHEMAS.get(kind, META_SCHEMAS["doc"]))
        except Exception:
            meta = {}
        build_word(tpl_path, "Документ", template_sections_for(kind), kind, meta)
        await m.answer_document(FSInputFile(tpl_path), caption="📄 Пустой шаблон — заполни пропуски от руки или в Word и распечатай")
        try:
            os.remove(tpl_path)
        except Exception:
            pass
        u["generations"] += 1
        u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — шаблон: {KIND_LABELS.get(kind, kind)}")
        await m.answer("Готово ✅", reply_markup=main_kb())
        await state.clear()
    finally:
        finish_job(uid)


# Выбор объёма (короткий/средний/подробный) имеет смысл только там, где объём реально
# влияет на глубину раскрытия темы: реферат, курсовая, эссе, доклад. Для конспекта
# он не нужен — конспект по своей природе всегда сжатый и короткий. Для обычного
# документа тоже убран — там слишком неопределённый формат, чтобы объём был осмысленным
# параметром. Для юридических документов (договоров, доверенностей, актов и т.п.) объём
# определяется их обязательной структурой по ГК РФ, а не пожеланием пользователя.
WORD_SIZE_KINDS = {"referat", "report", "essay", "coursework"}


async def word_after_input(m: Message, state: FSMContext):
    data = await state.get_data()
    kind = data.get("word_kind", "doc")
    if kind in WORD_SIZE_KINDS:
        await m.answer("Какой объём?", reply_markup=word_size_kb())
        await state.set_state(Form.waiting_word_size)
    else:
        await word_build_draft(m, state, data, "short")


@dp.message(Form.waiting_word_topic)
async def word_topic(m: Message, state: FSMContext):
    text = m.text or ""
    if len(text) > 500:
        await m.answer("Слишком длинно. Максимум 500 символов.")
        return
    await state.update_data(topic=text)
    await word_after_input(m, state)


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
    await word_after_input(m, state)


async def word_build_draft(m: Message, state: FSMContext, data: dict, size: str):
    await state.update_data(word_size=size)
    await m.answer("Собираю черновик…", reply_markup=ReplyKeyboardRemove())
    size_map = {"short": "поверхностное раскрытие темы — только суть и ключевые моменты, без глубокого разбора деталей и подпунктов, но по-настоящему содержательно, объём определяй по теме, не режь искусственно", "long": "полное раскрытие темы — подробно, с деталями, подпунктами и глубоким разбором, объём определяй по теме"}
    kind = data.get("word_kind", "doc")
    kind_name = WORD_KIND_DESC.get(kind, "документ")
    prompt = f"""Собери черновик-план (не полный текст): {kind_name}.
Тема/данные: {data.get('topic')}
Текст пользователя: {data.get('user_text')}
Доп: {data.get('extra')}
Итоговый документ будет иметь такой объём: {size_map[size]}. Но сейчас нужен не сам документ,
а короткий план для согласования с пользователем.
Исправь ошибки. {"Пиши формальным юридическим/деловым языком, без канцелярита-воды." if kind not in STUDY_KINDS else ANTI_AI_DETECTOR_STYLE}
Важно: никогда не придумывай паспортные данные, суммы, даты или ФИО, которых нет в данных пользователя -
для них ставь пропуски [указать ...]. Если документ содержательный (реферат, коммерческое предложение) -
план должен отражать суть темы конкретно, а не общими фразами.
Обычным текстом, коротко: название и 4 пункта структуры (каждый пункт — одна строка, без раскрытия
содержания). Без JSON."""
    sample = await ask_grok(prompt)
    if grok_failed(sample):
        await m.answer(
            "Не получилось получить ответ от нейросети. Попробуй ещё раз через минуту.",
            reply_markup=word_size_kb() if kind in WORD_SIZE_KINDS else word_confirm_kb()
        )
        return
    await state.update_data(sample=sample)
    await send_draft(m, f"Черновик готов ✅\n\n{sample}\n\nЕсли всё ок — собираем файл.", title=data.get("topic") or kind_name, reply_markup=word_confirm_kb())
    await state.set_state(Form.waiting_word_confirm)


@dp.message(Form.waiting_word_size)
async def word_size(m: Message, state: FSMContext):
    t = (m.text or "").lower()
    size = "long" if "полн" in t else "short"
    data = await state.get_data()
    await word_build_draft(m, state, data, size)


@dp.message(Form.waiting_word_confirm, F.text.in_(["Изменить запрос", "✏️ Изменить запрос"]))
async def word_change(m: Message, state: FSMContext):
    data = await state.get_data()
    if data.get("mode") == "user":
        await m.answer("Пришли новый текст:", reply_markup=ReplyKeyboardRemove())
        await state.set_state(Form.waiting_word_text)
    else:
        await m.answer("Напиши новую тему:", reply_markup=ReplyKeyboardRemove())
        await state.set_state(Form.waiting_word_topic)


@dp.message(Form.waiting_word_confirm, F.text.in_(["Добавить информацию", "➕ Добавить информацию", "➕ Добавить или изменить информацию"]))
async def word_add(m: Message, state: FSMContext):
    if (await state.get_data()).get("extra_used", 0) >= 3:
        await m.answer("Лимит добавлений исчерпан.", reply_markup=word_confirm_kb())
        return
    await m.answer("Напиши, что добавить или изменить. Максимум 800 символов.", reply_markup=ReplyKeyboardRemove())
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
    if grok_failed(sample):
        await m.answer(
            "Не получилось обновить черновик — нейросеть не ответила. Прошлый черновик остался как был.",
            reply_markup=word_confirm_kb()
        )
        await state.set_state(Form.waiting_word_confirm)
        return
    await state.update_data(sample=sample)
    await send_draft(m, f"Обновлённый черновик ✅\n\n{sample}\n\nВыбери действие:", title=data.get("topic") or WORD_KIND_DESC.get(data.get("word_kind", "doc"), "документ"), reply_markup=word_confirm_kb())
    await state.set_state(Form.waiting_word_confirm)


@dp.message(Form.waiting_word_confirm, F.text.in_(["Собрать документ", "✅ Собрать документ", "делай", "да", "ок"]))
async def word_build(m: Message, state: FSMContext):
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    if not start_job(uid):
        await m.answer("Уже собираю предыдущий документ, подожди немного 🙂")
        return
    await m.answer("Собираю Word. Обычно это быстрее презентации.")
    kind = data.get("word_kind", "doc")
    size = data.get("word_size", "short")
    size_map = {"short": "поверхностное раскрытие темы — только суть и ключевые моменты, без глубокого разбора деталей и подпунктов, но по-настоящему содержательно, объём определяй по теме, не режь искусственно", "long": "полное раскрытие темы — подробно, с деталями, подпунктами и глубоким разбором, объём определяй по теме"}
    kind_name = WORD_KIND_DESC.get(kind, "документ")
    # Схема мета-полей своя под каждый тип документа (модульный словарь META_SCHEMAS,
    # он же переиспользуется в режиме "Скачать шаблон") — раньше запрашивалась
    # всегда одна и та же (реферат/договор), из-за чего для заявления,
    # доверенности, КП поля вроде "кому"/"от кого" не заполнялись моделью
    # и в документе оставались заглушки, даже если пользователь их указал.
    meta_schema = META_SCHEMAS.get(kind, META_SCHEMAS["doc"])
    style_rule = ANTI_AI_DETECTOR_STYLE if kind in STUDY_KINDS else (
        "Пиши формальным юридическим/деловым языком, грамотно и точно по формулировкам ГК РФ, где применимо."
    )
    # Финальная сборка — это не короткий черновик-план, а весь текст документа целиком,
    # поэтому фиксированного лимита в 4000 токенов не хватало на курсовую/реферат с
    # "полным раскрытием темы" (там нужно 6000-10000+ слов), и модель тихо обрезала
    # содержание. Лимит теперь подбирается по типу документа и выбранной глубине.
    # Русский текст обычно занимает 1.5-2 токена на слово, плюс накладные расходы
    # на JSON-обёртку (кавычки, экранирование, ключи полей) - лимит с запасом,
    # чтобы модель физически не упёрлась в потолок на середине последнего раздела.
    if kind == "coursework":
        gen_max_tokens = 20000 if size == "long" else 8000
    elif kind == "referat":
        gen_max_tokens = 12000 if size == "long" else 6000
    elif kind in WORD_SIZE_KINDS:  # report, essay
        gen_max_tokens = 7000 if size == "long" else 4000
    else:
        gen_max_tokens = 6000
    # Расплывчатые формулировки вроде "несколько содержательных абзацев" модель
    # игнорирует и всё равно пишет коротко, даже с большим лимитом токенов -
    # нужны точные числовые ориентиры по объёму на раздел, иначе она работает
    # по привычке писать компактно, независимо от того, сколько токенов доступно.
    length_hint = ""
    if kind == "coursework":
        if size == "long":
            length_hint = "\nЭто полноценная курсовая работа для сдачи, суммарный объём всего документа — 6000-9000 слов. Каждый содержательный раздел (кроме титульного листа, содержания и списка литературы) должен быть НЕ МЕНЕЕ 600-900 слов (это примерно 4-6 полноценных абзацев с фактами, примерами, анализом, а не общими фразами) - пиши подробно и разворачивай мысль, а не сжимай её в 2-3 предложения."
        else:
            length_hint = "\nСуммарный объём документа — 2000-3000 слов. Каждый содержательный раздел (кроме титульного листа, содержания и списка литературы) — примерно 200-300 слов, по существу, без искусственного разжижения."
    elif kind == "referat":
        if size == "long":
            length_hint = "\nЭто полноценный реферат для сдачи, суммарный объём всего документа — 3000-5000 слов. Каждый содержательный раздел — НЕ МЕНЕЕ 400-600 слов (несколько развёрнутых абзацев с фактами и анализом), не сжимай в 2-3 предложения."
        else:
            length_hint = "\nСуммарный объём документа — 1200-1800 слов. Каждый содержательный раздел — примерно 150-250 слов, по существу."
    elif kind in WORD_SIZE_KINDS:  # report, essay
        if size == "long":
            length_hint = "\nСуммарный объём документа — 1500-2500 слов. Каждый содержательный раздел — примерно 250-400 слов, развёрнуто, с конкретикой."
        else:
            length_hint = "\nСуммарный объём документа — 600-1000 слов. Каждый содержательный раздел — примерно 100-180 слов, по существу, без воды."
    # Модель склонна писать заключение как краткий пересказ глав и не проверять
    # число источников - это отдельная, часто игнорируемая инструкция, поэтому
    # прописываем её явно, а не полагаемся на общее описание вида документа.
    structure_hint = ""
    if kind in ("coursework", "referat"):
        structure_hint = (
            "\nЗаключение пиши как самостоятельные выводы по каждой задаче из введения "
            "и рекомендации по теме - не пересказывай содержание глав.\n"
            "Раздел со списком литературы должен содержать не менее 5 источников "
            "(автор/название/год/издательство или ссылка), оформленных по ГОСТ."
        )
    raw = await ask_grok(f"""Собери Word: {kind_name}.
Данные: {data.get('topic')}
Текст пользователя: {data.get('user_text')}
Доп: {data.get('extra')}
Объём: {size_map[size]}{length_hint}{structure_hint}
{style_rule}
Если в данных пользователя есть даты, ФИО, паспортные данные, суммы, названия сторон - подставь их в meta и в текст
точно как есть, ничего не меняя и не придумывая.
Никогда не выдумывай паспортные данные, суммы, даты или ФИО, которых нет в данных пользователя -
для недостающих данных ставь [указать ...].
Для содержательных документов (реферат, коммерческое предложение, заявление) разделы должны раскрывать
тему конкретно и по существу, без воды и общих фраз.
Только JSON:
{{"title":"...","meta":{meta_schema},"sections":[{{"title":"Введение","content":"абзац1\n\nабзац2"}}]}}""", max_tokens=gen_max_tokens)
    try:
        content = extract_json(raw)
        if not isinstance(content.get("sections"), list) or not content["sections"]:
            raise ValueError("В ответе модели нет разделов документа")
    except Exception as e:
        print("Word JSON parse error:", e)
        finish_job(uid)
        await m.answer("Не собрал текст. Попробуй ещё раз.", reply_markup=main_kb())
        await state.clear()
        return

    try:
        docx_path = f"/tmp/doc_{uid}.docx"
        build_word(
            docx_path,
            content.get("title", "Документ"),
            content.get("sections", []),
            kind,
            content.get("meta") or {}
        )

        pdf_path = f"/tmp/doc_{uid}.pdf"
        pdf = canvas.Canvas(pdf_path, pagesize=A4)
        w, h = A4
        fn, fb = register_pdf_fonts()

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
            for line in wrap_lines((s.get("content") or ""), fn, 9, w - 80):
                if y < 50:
                    pdf.showPage()
                    y = h - 50
                    pdf.setFont(fn, 9)
                pdf.drawString(40, y, line)
                y -= 12
            y -= 10
        pdf.save()

        fname = safe_filename(content.get("title"), fallback=WORD_KIND_DESC.get(kind, "Документ"))
        await m.answer_document(FSInputFile(docx_path, filename=f"{fname}.docx"), caption="📄 Word — этот файл можно править")
        await m.answer_document(FSInputFile(pdf_path, filename=f"{fname}.pdf"), caption="📄 PDF-копия")
        u["generations"] += 1
        u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
        for p in (docx_path, pdf_path):
            try:
                os.remove(p)
            except Exception:
                pass
        final_msg = "Документ готов ✅"
        if kind not in STUDY_KINDS:
            final_msg += "\n\n⚠️ Проверь пропуски в тексте перед распечаткой."
        await m.answer(final_msg, reply_markup=main_kb())
        await state.clear()
    finally:
        finish_job(uid)


@dp.message(Form.waiting_word_confirm)
async def waiting_word_confirm_fallback(m: Message, state: FSMContext):
    await m.answer(
        "Не понял. Выбери действие кнопкой ниже, или напиши /cancel, чтобы начать заново.",
        reply_markup=word_confirm_kb()
    )


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
        u = get_user(uid)
        u["plan"] = "premium"
        u["generations"] = 0
        await m.answer(f"Доступ выдан пользователю {uid}, счётчик генераций сброшен")
    except (IndexError, ValueError):
        await m.answer("Формат: /grant user_id")


@dp.errors()
async def global_error_handler(event):
    """Подстраховка на случай непредвиденных ошибок вне двух основных
    обработчиков (те уже защищены через try/finally выше и сами снимают
    busy). Здесь — просто не даём боту молчать и на всякий случай ещё раз
    снимаем блокировку, если получится определить пользователя."""
    print("Необработанная ошибка:", repr(event.exception))
    uid = None
    update = getattr(event, "update", None)
    msg = getattr(update, "message", None) or getattr(update, "callback_query", None)
    if msg is not None:
        user = getattr(msg, "from_user", None)
        uid = getattr(user, "id", None)
    if uid is not None:
        try:
            finish_job(uid)
            await bot.send_message(uid, "Что-то пошло не так. Попробуй ещё раз или напиши /cancel.")
        except Exception as e:
            print("Не смог уведомить пользователя об ошибке:", e)
    return True


async def main():
    print("Бот запущен")
    print("REPLICATE TOKEN:", "YES" if REPLICATE_API_TOKEN else "NO")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
