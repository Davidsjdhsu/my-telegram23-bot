import os
import asyncio
import json
import random
import re
import time
from collections import deque
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
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
import openpyxl
from openpyxl.styles import Font as XlFont, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
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


def safe_filename(title: str, fallback: str = "Документ", author: str = None) -> str:
    """Делает из темы/заголовка документа безопасное имя файла:
    убирает символы, которые нельзя использовать в имени файла (Windows/Telegram),
    обрезает длину, чтобы не упереться в лимиты, и подставляет fallback, если
    заголовок пустой. Если передан author (реальное ФИО, не плейсхолдер вида
    "[ФИО студента]") - добавляет его через тире, чтобы файл не выглядел безлико."""
    import re as _re

    def _clean(s):
        s = _re.sub(r'[\\/:*?"<>|\n\r\t]', " ", s)
        return _re.sub(r"\s+", " ", s).strip()

    text = _clean((title or "").strip() or fallback)
    text = text[:70].strip() or fallback
    if author:
        a = _clean(author.strip())
        if a and not a.startswith("["):  # не подставляем невыполненный плейсхолдер
            a = a[:40]
            text = f"{text[:95 - len(a) - 3]} - {a}"
    text = text[:110].strip() or fallback
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
    waiting_pres_photo_choice = State()
    waiting_pres_photos = State()
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
    waiting_excel_category = State()
    waiting_excel_kind = State()
    waiting_excel_mode = State()
    waiting_excel_topic = State()
    waiting_excel_data = State()
    waiting_excel_startup_data = State()
    waiting_excel_confirm = State()
    waiting_excel_extra = State()
    waiting_language = State()


# ==================== ЯЗЫКИ / i18n ====================
# 7 языков интерфейса и генерации контента. Выбор языка хранится в users_db[uid]["lang"],
# по умолчанию "ru" (для уже существующих пользователей и новых, пока не выбрали другой).
LANGS = {
    "ru": {"flag": "🇷🇺", "name": "Русский", "prompt_name": "русском"},
    "en": {"flag": "🇬🇧", "name": "English", "prompt_name": "английском"},
    "de": {"flag": "🇩🇪", "name": "Deutsch", "prompt_name": "немецком"},
    "ar": {"flag": "🇸🇦", "name": "العربية", "prompt_name": "арабском"},
    "zh": {"flag": "🇨🇳", "name": "中文", "prompt_name": "китайском"},
    "es": {"flag": "🇪🇸", "name": "Español", "prompt_name": "испанском"},
    "fr": {"flag": "🇫🇷", "name": "Français", "prompt_name": "французском"},
}
RTL_LANGS = {"ar"}

# Переводы интерфейса бота (кнопки, служебные сообщения). Ключ - смысловой id строки,
# значение - словарь {код_языка: текст}. Если перевода для конкретного языка нет -
# t() подставит русский вариант, чтобы бот не падал и не показывал пустоту.
TR = {
    "btn_pres": {"ru": "📊 Сделать презентацию", "en": "📊 Create a presentation", "de": "📊 Präsentation erstellen",
                 "ar": "📊 إنشاء عرض تقديمي", "zh": "📊 制作演示文稿", "es": "📊 Crear una presentación", "fr": "📊 Créer une présentation"},
    "btn_word": {"ru": "📄 Сделать документ Word", "en": "📄 Create a Word document", "de": "📄 Word-Dokument erstellen",
                 "ar": "📄 إنشاء مستند Word", "zh": "📄 制作 Word 文档", "es": "📄 Crear un documento Word", "fr": "📄 Créer un document Word"},
    "btn_excel": {"ru": "📈 Сделать таблицу Excel", "en": "📈 Create an Excel table", "de": "📈 Excel-Tabelle erstellen",
                  "ar": "📈 إنشاء جدول Excel", "zh": "📈 制作 Excel 表格", "es": "📈 Crear una tabla Excel", "fr": "📈 Créer un tableau Excel"},
    "btn_history": {"ru": "📁 Моя история", "en": "📁 My history", "de": "📁 Mein Verlauf",
                    "ar": "📁 سجلي", "zh": "📁 我的历史", "es": "📁 Mi historial", "fr": "📁 Mon historique"},
    "btn_plan": {"ru": "ℹ️ Мой тариф", "en": "ℹ️ My plan", "de": "ℹ️ Mein Tarif",
                 "ar": "ℹ️ خطتي", "zh": "ℹ️ 我的套餐", "es": "ℹ️ Mi plan", "fr": "ℹ️ Mon forfait"},
    "btn_language": {"ru": "🌐 Язык", "en": "🌐 Language", "de": "🌐 Sprache",
                     "ar": "🌐 اللغة", "zh": "🌐 语言", "es": "🌐 Idioma", "fr": "🌐 Langue"},
    "btn_main_menu": {"ru": "🏠 Главное меню", "en": "🏠 Main menu", "de": "🏠 Hauptmenü",
                      "ar": "🏠 القائمة الرئيسية", "zh": "🏠 主菜单", "es": "🏠 Menú principal", "fr": "🏠 Menu principal"},
    "btn_ai_generate": {"ru": "✨ Сгенерировать с ИИ", "en": "✨ Generate with AI", "de": "✨ Mit KI generieren",
                        "ar": "✨ إنشاء بالذكاء الاصطناعي", "zh": "✨ 用AI生成", "es": "✨ Generar con IA", "fr": "✨ Générer avec l'IA"},
    "btn_own_text": {"ru": "📝 Вставить свой текст", "en": "📝 Paste your own text", "de": "📝 Eigenen Text einfügen",
                     "ar": "📝 إدراج النص الخاص بك", "zh": "📝 粘贴自己的文本", "es": "📝 Pegar tu propio texto", "fr": "📝 Coller votre propre texte"},
    "btn_own_data": {"ru": "📝 Ввести свои данные", "en": "📝 Enter your own data", "de": "📝 Eigene Daten eingeben",
                     "ar": "📝 إدخال بياناتك الخاصة", "zh": "📝 输入自己的数据", "es": "📝 Introducir tus propios datos", "fr": "📝 Saisir vos propres données"},
    "btn_template": {"ru": "📄 Скачать шаблон", "en": "📄 Download template", "de": "📄 Vorlage herunterladen",
                     "ar": "📄 تنزيل القالب", "zh": "📄 下载模板", "es": "📄 Descargar plantilla", "fr": "📄 Télécharger le modèle"},
    "btn_full_version": {"ru": "✅ Делать полную версию", "en": "✅ Make full version", "de": "✅ Vollversion erstellen",
                         "ar": "✅ إنشاء النسخة الكاملة", "zh": "✅ 生成完整版本", "es": "✅ Crear versión completa", "fr": "✅ Créer la version complète"},
    "btn_add_info": {"ru": "➕ Добавить или изменить информацию", "en": "➕ Add or change information", "de": "➕ Informationen hinzufügen/ändern",
                     "ar": "➕ إضافة أو تعديل المعلومات", "zh": "➕ 添加或修改信息", "es": "➕ Añadir o modificar información", "fr": "➕ Ajouter ou modifier des informations"},
    "btn_change_style": {"ru": "🎨 Изменить стиль", "en": "🎨 Change style", "de": "🎨 Stil ändern",
                         "ar": "🎨 تغيير النمط", "zh": "🎨 更改风格", "es": "🎨 Cambiar estilo", "fr": "🎨 Changer le style"},
    "btn_change_topic": {"ru": "✏️ Изменить тему", "en": "✏️ Change topic", "de": "✏️ Thema ändern",
                         "ar": "✏️ تغيير الموضوع", "zh": "✏️ 更改主题", "es": "✏️ Cambiar el tema", "fr": "✏️ Changer le sujet"},
    "btn_change_query": {"ru": "✏️ Изменить запрос", "en": "✏️ Change request", "de": "✏️ Anfrage ändern",
                         "ar": "✏️ تغيير الطلب", "zh": "✏️ 更改请求", "es": "✏️ Cambiar la solicitud", "fr": "✏️ Modifier la demande"},
    "btn_ai_photos": {"ru": "✨ Все фото через ИИ", "en": "✨ All photos via AI", "de": "✨ Alle Fotos per KI",
                      "ar": "✨ كل الصور بالذكاء الاصطناعي", "zh": "✨ 全部用AI生成图片", "es": "✨ Todas las fotos con IA", "fr": "✨ Toutes les photos via l'IA"},
    "btn_own_photos": {"ru": "📷 Добавить свои фото", "en": "📷 Add your own photos", "de": "📷 Eigene Fotos hinzufügen",
                       "ar": "📷 إضافة صورك الخاصة", "zh": "📷 添加自己的照片", "es": "📷 Añadir tus propias fotos", "fr": "📷 Ajouter vos propres photos"},
    "btn_photos_done": {"ru": "✅ Готово, собери презентацию", "en": "✅ Done, build the presentation", "de": "✅ Fertig, Präsentation erstellen",
                        "ar": "✅ تم، أنشئ العرض التقديمي", "zh": "✅ 完成，生成演示文稿", "es": "✅ Listo, crea la presentación", "fr": "✅ Terminé, créez la présentation"},
    "btn_build_doc": {"ru": "✅ Собрать документ", "en": "✅ Build document", "de": "✅ Dokument erstellen",
                      "ar": "✅ إنشاء المستند", "zh": "✅ 生成文档", "es": "✅ Crear documento", "fr": "✅ Créer le document"},
    "btn_build_table": {"ru": "✅ Собрать таблицу", "en": "✅ Build table", "de": "✅ Tabelle erstellen",
                        "ar": "✅ إنشاء الجدول", "zh": "✅ 生成表格", "es": "✅ Crear tabla", "fr": "✅ Créer le tableau"},
    "btn_more_docs": {"ru": "➕ Показать ещё документы", "en": "➕ Show more documents", "de": "➕ Weitere Dokumente anzeigen",
                      "ar": "➕ عرض المزيد من المستندات", "zh": "➕ 显示更多文档", "es": "➕ Mostrar más documentos", "fr": "➕ Afficher plus de documents"},
    "btn_back_categories": {"ru": "⬅️ Назад к категориям", "en": "⬅️ Back to categories", "de": "⬅️ Zurück zu Kategorien",
                            "ar": "⬅️ العودة إلى الفئات", "zh": "⬅️ 返回分类", "es": "⬅️ Volver a categorías", "fr": "⬅️ Retour aux catégories"},
    "btn_5slides": {"ru": "5️⃣ 5 слайдов", "en": "5️⃣ 5 slides", "de": "5️⃣ 5 Folien", "ar": "5️⃣ 5 شرائح", "zh": "5️⃣ 5张幻灯片", "es": "5️⃣ 5 diapositivas", "fr": "5️⃣ 5 diapositives"},
    "btn_8slides": {"ru": "8️⃣ 8 слайдов", "en": "8️⃣ 8 slides", "de": "8️⃣ 8 Folien", "ar": "8️⃣ 8 شرائح", "zh": "8️⃣ 8张幻灯片", "es": "8️⃣ 8 diapositivas", "fr": "8️⃣ 8 diapositives"},
    "btn_10slides": {"ru": "🔟 10 слайдов", "en": "🔟 10 slides", "de": "🔟 10 Folien", "ar": "🔟 10 شرائح", "zh": "🔟 10张幻灯片", "es": "🔟 10 diapositivas", "fr": "🔟 10 diapositives"},
    "btn_shallow": {"ru": "Поверхностное раскрытие темы", "en": "Brief coverage of the topic", "de": "Oberflächliche Themenbehandlung",
                    "ar": "تغطية سطحية للموضوع", "zh": "简要阐述主题", "es": "Cobertura breve del tema", "fr": "Traitement succinct du sujet"},
    "btn_deep": {"ru": "Полное раскрытие темы", "en": "Full coverage of the topic", "de": "Vollständige Themenbehandlung",
                 "ar": "تغطية كاملة للموضوع", "zh": "全面阐述主题", "es": "Cobertura completa del tema", "fr": "Traitement complet du sujet"},
    "btn_keep_style": {"ru": "👍 Оставить этот стиль", "en": "👍 Keep this style", "de": "👍 Diesen Stil behalten",
                       "ar": "👍 الاحتفاظ بهذا النمط", "zh": "👍 保留此风格", "es": "👍 Mantener este estilo", "fr": "👍 Conserver ce style"},
    "msg_choose_lang": {"ru": "Выбери язык интерфейса и генерации:", "en": "Choose interface and generation language:",
                        "de": "Wähle die Sprache für Oberfläche und Generierung:", "ar": "اختر لغة الواجهة والإنشاء:",
                        "zh": "选择界面和生成内容的语言：", "es": "Elige el idioma de la interfaz y la generación:", "fr": "Choisissez la langue de l'interface et de la génération :"},
    "msg_lang_set": {"ru": "Готово! Дальше всё — интерфейс и документы — будет на этом языке. 👇", "en": "Done! From now on the interface and documents will be in this language. 👇",
                     "de": "Fertig! Ab jetzt sind Oberfläche und Dokumente in dieser Sprache. 👇", "ar": "تم! من الآن ستكون الواجهة والمستندات بهذه اللغة. 👇",
                     "zh": "完成！从现在起界面和文档都将使用该语言。👇", "es": "¡Listo! A partir de ahora la interfaz y los documentos estarán en este idioma. 👇",
                     "fr": "C'est fait ! Désormais l'interface et les documents seront dans cette langue. 👇"},
    "msg_welcome": {
        "ru": "Привет, {name} 👋\n\nЯ собираю красивые презентации, документы Word и таблицы Excel: текст, стиль и оформление — по твоей теме или из твоих данных.\n\nНажми кнопку ниже и начнём.",
        "en": "Hi, {name} 👋\n\nI put together polished presentations, Word documents and Excel tables: text, style and layout — from your topic or your own data.\n\nTap a button below to start.",
        "de": "Hallo, {name} 👋\n\nIch erstelle ansprechende Präsentationen, Word-Dokumente und Excel-Tabellen: Text, Stil und Layout — zu deinem Thema oder aus deinen Daten.\n\nTippe unten auf einen Button, um zu starten.",
        "ar": "مرحباً {name} 👋\n\nأقوم بإعداد عروض تقديمية ومستندات Word وجداول Excel أنيقة: النص والنمط والتصميم — بناءً على موضوعك أو بياناتك.\n\nاضغط الزر أدناه للبدء.",
        "zh": "你好，{name} 👋\n\n我可以帮你制作精美的演示文稿、Word文档和Excel表格：根据你的主题或数据生成文本、风格和排版。\n\n点击下方按钮开始吧。",
        "es": "Hola, {name} 👋\n\nCreo presentaciones, documentos Word y tablas Excel cuidados: texto, estilo y diseño, a partir de tu tema o tus propios datos.\n\nToca un botón abajo para empezar.",
        "fr": "Bonjour {name} 👋\n\nJe crée des présentations, des documents Word et des tableaux Excel soignés : texte, style et mise en page — à partir de votre sujet ou de vos propres données.\n\nAppuyez sur un bouton ci-dessous pour commencer.",
    },
    "msg_main_menu": {"ru": "Главное меню 👇", "en": "Main menu 👇", "de": "Hauptmenü 👇", "ar": "القائمة الرئيسية 👇",
                      "zh": "主菜单 👇", "es": "Menú principal 👇", "fr": "Menu principal 👇"},
    "msg_cancelled": {"ru": "Отменил текущее действие. Начнём заново 👇", "en": "Cancelled the current action. Let's start over 👇",
                      "de": "Aktuelle Aktion abgebrochen. Fangen wir neu an 👇", "ar": "تم إلغاء الإجراء الحالي. لنبدأ من جديد 👇",
                      "zh": "已取消当前操作。重新开始吧 👇", "es": "Acción actual cancelada. Empecemos de nuevo 👇", "fr": "Action en cours annulée. Recommençons 👇"},
    "msg_limit": {"ru": "Лимит генераций закончился.", "en": "Generation limit reached.", "de": "Generierungslimit erreicht.",
                  "ar": "انتهى حد الإنشاء.", "zh": "生成次数已用完。", "es": "Se agotó el límite de generaciones.", "fr": "Limite de générations atteinte."},
    "msg_table_ready": {"ru": "Таблица готова ✅", "en": "Table ready ✅", "de": "Tabelle fertig ✅", "ar": "الجدول جاهز ✅",
                        "zh": "表格已完成 ✅", "es": "Tabla lista ✅", "fr": "Tableau prêt ✅"},
    "msg_doc_ready": {"ru": "Документ готов ✅", "en": "Document ready ✅", "de": "Dokument fertig ✅", "ar": "المستند جاهز ✅",
                      "zh": "文档已完成 ✅", "es": "Documento listo ✅", "fr": "Document prêt ✅"},
    "msg_build_error_table": {"ru": "Не собрал таблицу. Попробуй ещё раз.", "en": "Couldn't build the table. Try again.",
                              "de": "Tabelle konnte nicht erstellt werden. Versuche es erneut.", "ar": "تعذر إنشاء الجدول. حاول مرة أخرى.",
                              "zh": "未能生成表格。请重试。", "es": "No se pudo crear la tabla. Inténtalo de nuevo.", "fr": "Impossible de créer le tableau. Réessayez."},
    "msg_build_error_doc": {"ru": "Не собрал документ. Попробуй ещё раз.", "en": "Couldn't build the document. Try again.",
                            "de": "Dokument konnte nicht erstellt werden. Versuche es erneut.", "ar": "تعذر إنشاء المستند. حاول مرة أخرى.",
                            "zh": "未能生成文档。请重试。", "es": "No se pudo crear el documento. Inténtalo de nuevo.", "fr": "Impossible de créer le document. Réessayez."},
    "msg_for_whom_table": {"ru": "Для кого таблица?", "en": "Who is the table for?", "de": "Für wen ist die Tabelle?",
                           "ar": "لمن هذا الجدول؟", "zh": "表格是给谁用的？", "es": "¿Para quién es la tabla?", "fr": "Pour qui est ce tableau ?"},
    "msg_for_whom_doc": {"ru": "Для кого документ?", "en": "Who is the document for?", "de": "Für wen ist das Dokument?",
                         "ar": "لمن هذا المستند؟", "zh": "文档是给谁用的？", "es": "¿Para quién es el documento?", "fr": "Pour qui est ce document ?"},
    "msg_pick_category": {"ru": "Выбери категорию кнопкой ниже.", "en": "Pick a category using the button below.", "de": "Wähle eine Kategorie mit dem Button unten.",
                          "ar": "اختر فئة من الأزرار أدناه.", "zh": "请用下方按钮选择分类。", "es": "Elige una categoría con el botón de abajo.", "fr": "Choisis une catégorie avec le bouton ci-dessous."},
    "msg_which_table": {"ru": "Какая таблица нужна?", "en": "Which table do you need?", "de": "Welche Tabelle brauchst du?",
                        "ar": "أي جدول تحتاج؟", "zh": "需要哪种表格？", "es": "¿Qué tabla necesitas?", "fr": "Quel tableau te faut-il ?"},
    "msg_which_doc": {"ru": "Какой документ нужен?", "en": "Which document do you need?", "de": "Welches Dokument brauchst du?",
                      "ar": "أي مستند تحتاج؟", "zh": "需要哪种文档？", "es": "¿Qué documento necesitas?", "fr": "Quel document te faut-il ?"},
    "msg_pick_table": {"ru": "Выбери таблицу кнопкой ниже.", "en": "Pick a table using the button below.", "de": "Wähle eine Tabelle mit dem Button unten.",
                       "ar": "اختر جدولاً من الأزرار أدناه.", "zh": "请用下方按钮选择表格。", "es": "Elige una tabla con el botón de abajo.", "fr": "Choisis un tableau avec le bouton ci-dessous."},
    "msg_pick_doc": {"ru": "Выбери документ кнопкой ниже.", "en": "Pick a document using the button below.", "de": "Wähle ein Dokument mit dem Button unten.",
                     "ar": "اختر مستنداً من الأزرار أدناه.", "zh": "请用下方按钮选择文档。", "es": "Elige un documento con el botón de abajo.", "fr": "Choisis un document avec le bouton ci-dessous."},
    "msg_more_docs": {"ru": "Ещё документы:", "en": "More documents:", "de": "Weitere Dokumente:",
                      "ar": "المزيد من المستندات:", "zh": "更多文档：", "es": "Más documentos:", "fr": "Plus de documents :"},
    "msg_build_table_q": {"ru": "Собрать таблицу?", "en": "Build the table?", "de": "Tabelle erstellen?",
                          "ar": "هل ننشئ الجدول؟", "zh": "生成表格吗？", "es": "¿Creo la tabla?", "fr": "On crée le tableau ?"},
    "msg_build_model_q": {"ru": "Собрать финмодель?", "en": "Build the financial model?", "de": "Finanzmodell erstellen?",
                          "ar": "هل ننشئ النموذج المالي؟", "zh": "生成财务模型吗？", "es": "¿Creo el modelo financiero?", "fr": "On crée le modèle financier ?"},
    "msg_resend_data": {"ru": "Пришли исходные данные заново.", "en": "Please send the source data again.", "de": "Sende die Ausgangsdaten erneut.",
                        "ar": "أرسل البيانات الأصلية مرة أخرى.", "zh": "请重新发送原始数据。", "es": "Envía de nuevo los datos originales.", "fr": "Renvoie les données d'origine."},
    "msg_what_add": {"ru": "Что добавить или уточнить?", "en": "What should I add or clarify?", "de": "Was soll ich hinzufügen oder ändern?",
                     "ar": "ما الذي تريد إضافته أو توضيحه؟", "zh": "需要添加或说明什么？", "es": "¿Qué añado o aclaro?", "fr": "Qu'est-ce qu'il faut ajouter ou préciser ?"},
    "msg_choose_action": {"ru": "Выбери действие кнопкой ниже.", "en": "Choose an action using the button below.", "de": "Wähle eine Aktion mit dem Button unten.",
                          "ar": "اختر إجراءً من الأزرار أدناه.", "zh": "请用下方按钮选择操作。", "es": "Elige una acción con el botón de abajo.", "fr": "Choisis une action avec le bouton ci-dessous."},
    "msg_new_text": {"ru": "Пришли новый текст:", "en": "Send the new text:", "de": "Sende den neuen Text:",
                     "ar": "أرسل النص الجديد:", "zh": "请发送新文本：", "es": "Envía el nuevo texto:", "fr": "Envoie le nouveau texte :"},
    "msg_new_topic": {"ru": "Напиши новую тему:", "en": "Write the new topic:", "de": "Schreibe das neue Thema:",
                      "ar": "اكتب الموضوع الجديد:", "zh": "请输入新主题：", "es": "Escribe el nuevo tema:", "fr": "Écris le nouveau sujet :"},
    "msg_extra_limit": {"ru": "Лимит добавлений исчерпан.", "en": "You've reached the edit limit.", "de": "Das Limit für Ergänzungen ist erreicht.",
                        "ar": "لقد استنفدت حد الإضافات.", "zh": "修改次数已用完。", "es": "Se agotó el límite de adiciones.", "fr": "Limite de modifications atteinte."},
    "msg_extra_prompt": {"ru": "Напиши, что добавить или изменить. Максимум 800 символов.", "en": "Write what to add or change. Max 800 characters.",
                         "de": "Schreibe, was hinzugefügt oder geändert werden soll. Max. 800 Zeichen.", "ar": "اكتب ما تريد إضافته أو تغييره. الحد الأقصى ٨٠٠ حرف.",
                         "zh": "请输入要添加或修改的内容，最多800字。", "es": "Escribe qué añadir o cambiar. Máximo 800 caracteres.", "fr": "Écris ce qu'il faut ajouter ou modifier. 800 caractères max."},
    "msg_no_template_study": {"ru": "Для учебных документов шаблон не предусмотрен.", "en": "There's no template for study documents.",
                              "de": "Für Studienarbeiten gibt es keine Vorlage.", "ar": "لا يوجد قالب للمستندات الدراسية.",
                              "zh": "学习类文档暂无模板。", "es": "No hay plantilla para documentos académicos.", "fr": "Il n'y a pas de modèle pour les documents scolaires."},
    "msg_building_template": {"ru": "Собираю шаблон…", "en": "Building the template…", "de": "Vorlage wird erstellt…",
                              "ar": "جاري إنشاء القالب…", "zh": "正在生成模板…", "es": "Creando la plantilla…", "fr": "Création du modèle…"},
    "msg_ready": {"ru": "Готово ✅", "en": "Done ✅", "de": "Fertig ✅", "ar": "تم ✅", "zh": "完成 ✅", "es": "Listo ✅", "fr": "Terminé ✅"},
    "msg_which_size": {"ru": "Какой объём?", "en": "How much detail?", "de": "Welcher Umfang?",
                       "ar": "ما الحجم المطلوب؟", "zh": "需要多少内容？", "es": "¿Qué extensión?", "fr": "Quelle ampleur ?"},
    "msg_building_draft": {"ru": "Собираю черновик…", "en": "Building the draft…", "de": "Entwurf wird erstellt…",
                           "ar": "جاري إعداد المسودة…", "zh": "正在生成草稿…", "es": "Creando el borrador…", "fr": "Création du brouillon…"},
    "msg_no_text": {"ru": "Не собрал текст. Нажми ещё раз «{btn}».", "en": "Couldn't build the text. Tap \"{btn}\" again.",
                    "de": "Text konnte nicht erstellt werden. Tippe erneut auf „{btn}“.", "ar": "تعذر إنشاء النص. اضغط «{btn}» مرة أخرى.",
                    "zh": "未能生成文本，请再次点击「{btn}」。", "es": "No se pudo crear el texto. Vuelve a tocar «{btn}».", "fr": "Impossible de créer le texte. Appuie à nouveau sur « {btn} »."},
    "msg_photo_wait": {"ru": "Фото пока не пришло. Пришли хотя бы одно или вернись к «Все фото через ИИ».",
                       "en": "No photo yet. Send at least one, or go back to \"All photos via AI\".",
                       "de": "Noch kein Foto erhalten. Sende mindestens eins oder gehe zurück zu „Alle Fotos per KI“.",
                       "ar": "لم تصل أي صورة بعد. أرسل صورة واحدة على الأقل أو ارجع إلى «كل الصور بالذكاء الاصطناعي».",
                       "zh": "还没有收到照片。请至少发送一张，或返回「全部用AI生成图片」。",
                       "es": "Aún no llegó ninguna foto. Envía al menos una o vuelve a «Todas las fotos con IA».",
                       "fr": "Aucune photo reçue pour l'instant. Envoie-en au moins une ou reviens à « Toutes les photos via l'IA »."},
    "msg_send_photos": {"ru": "Пришли фото 📷 или нажми «Готово, собери презентацию», когда закончишь.",
                        "en": "Send photos 📷 or tap \"Done, build the presentation\" when finished.",
                        "de": "Sende Fotos 📷 oder tippe auf „Fertig, Präsentation erstellen“, wenn du fertig bist.",
                        "ar": "أرسل الصور 📷 أو اضغط «تم، أنشئ العرض التقديمي» عند الانتهاء.",
                        "zh": "发送照片📷，完成后点击「完成，生成演示文稿」。",
                        "es": "Envía fotos 📷 o toca «Listo, crea la presentación» cuando termines.",
                        "fr": "Envoie des photos 📷 ou appuie sur « Terminé, créez la présentation » quand tu as fini."},
    "msg_topic_prompt": {"ru": "Напиши тему. Можно коротко, например: киты, крипта, школа.",
                         "en": "Write the topic. Can be short, e.g.: whales, crypto, school.",
                         "de": "Schreibe das Thema. Kurz ist ok, z. B.: Wale, Krypto, Schule.",
                         "ar": "اكتب الموضوع. يمكن أن يكون قصيراً، مثل: الحيتان، العملات الرقمية، المدرسة.",
                         "zh": "请输入主题，可以简短一些，例如：鲸鱼、加密货币、学校。",
                         "es": "Escribe el tema. Puede ser breve, p. ej.: ballenas, cripto, escuela.",
                         "fr": "Écris le sujet. Ça peut être court, ex. : baleines, crypto, école."},
    "msg_own_text_prompt": {"ru": "Пришли свой текст 📝\nМожно черновик — я поправлю ошибки и соберу структуру.",
                            "en": "Send your text 📝\nA draft is fine — I'll fix errors and build the structure.",
                            "de": "Sende deinen Text 📝\nEin Entwurf reicht — ich korrigiere Fehler und erstelle die Struktur.",
                            "ar": "أرسل نصك 📝\nيمكن أن تكون مسودة — سأصحح الأخطاء وأبني الهيكل.",
                            "zh": "请发送你的文本📝\n草稿也可以——我会修正错误并构建结构。",
                            "es": "Envía tu texto 📝\nUn borrador está bien — corregiré errores y armaré la estructura.",
                            "fr": "Envoie ton texte 📝\nUn brouillon suffit — je corrigerai les erreurs et je construirai la structure."},
    "msg_which_style": {"ru": "Выбери стиль оформления:", "en": "Choose a design style:", "de": "Wähle einen Gestaltungsstil:",
                        "ar": "اختر نمط التصميم:", "zh": "选择设计风格：", "es": "Elige un estilo de diseño:", "fr": "Choisis un style de design :"},
    "msg_pick_style": {"ru": "Выбери стиль кнопкой ниже.", "en": "Pick a style using the button below.", "de": "Wähle einen Stil mit dem Button unten.",
                       "ar": "اختر نمطاً من الأزرار أدناه.", "zh": "请用下方按钮选择风格。", "es": "Elige un estilo con el botón de abajo.", "fr": "Choisis un style avec le bouton ci-dessous."},
    "msg_how_many_slides": {"ru": "Сколько слайдов сделать?", "en": "How many slides?", "de": "Wie viele Folien?",
                            "ar": "كم عدد الشرائح؟", "zh": "需要多少张幻灯片？", "es": "¿Cuántas diapositivas?", "fr": "Combien de diapositives ?"},
    "msg_building_sample": {"ru": "Секунду, собираю пробный вариант…", "en": "One moment, building a sample…", "de": "Einen Moment, ich erstelle einen Testentwurf…",
                            "ar": "لحظة، جاري إعداد نسخة تجريبية…", "zh": "请稍等，正在生成样例…", "es": "Un momento, creando una versión de prueba…", "fr": "Un instant, je prépare un aperçu…"},
    "msg_choose_action_below": {"ru": "Выбери действие:", "en": "Choose an action:", "de": "Wähle eine Aktion:",
                                "ar": "اختر إجراءً:", "zh": "选择操作：", "es": "Elige una acción:", "fr": "Choisis une action :"},
    "msg_how_build_doc": {"ru": "Как собираем документ?\n\n✨ Сгенерировать с ИИ — я сам напишу текст.\n📝 Вставить свой текст — пришли материал, я поправлю и оформлю.",
                          "en": "How should we build the document?\n\n✨ Generate with AI — I'll write the text myself.\n📝 Paste your own text — send the material, I'll fix and format it.",
                          "de": "Wie erstellen wir das Dokument?\n\n✨ Mit KI generieren — ich schreibe den Text selbst.\n📝 Eigenen Text einfügen — sende das Material, ich korrigiere und formatiere es.",
                          "ar": "كيف ننشئ المستند؟\n\n✨ إنشاء بالذكاء الاصطناعي — سأكتب النص بنفسي.\n📝 إدراج نصك الخاص — أرسل المادة وسأصححها وأنسقها.",
                          "zh": "如何生成文档？\n\n✨ 用AI生成——由我来撰写文本。\n📝 粘贴自己的文本——发送素材，我会修正并排版。",
                          "es": "¿Cómo creamos el documento?\n\n✨ Generar con IA — yo escribiré el texto.\n📝 Pegar tu propio texto — envía el material, lo corregiré y daré formato.",
                          "fr": "Comment créons-nous le document ?\n\n✨ Générer avec l'IA — j'écrirai le texte moi-même.\n📝 Coller votre propre texte — envoyez le contenu, je le corrigerai et le mettrai en forme."},
    "msg_template_hint_suffix": {"ru": "\n📄 Скачать шаблон — пустой бланк с пропусками для заполнения от руки.",
                                 "en": "\n📄 Download template — a blank form with fields to fill in by hand.",
                                 "de": "\n📄 Vorlage herunterladen — ein leeres Formular mit Feldern zum handschriftlichen Ausfüllen.",
                                 "ar": "\n📄 تنزيل القالب — نموذج فارغ بحقول للتعبئة يدوياً.",
                                 "zh": "\n📄 下载模板——带填空的空白表单，可手动填写。",
                                 "es": "\n📄 Descargar plantilla — un formulario en blanco con espacios para rellenar a mano.",
                                 "fr": "\n📄 Télécharger le modèle — un formulaire vierge à remplir à la main."},
    "msg_too_long_500": {"ru": "Слишком длинно. Максимум 500 символов.", "en": "Too long. Max 500 characters.", "de": "Zu lang. Max. 500 Zeichen.",
                         "ar": "طويل جداً. الحد الأقصى ٥٠٠ حرف.", "zh": "太长了，最多500字。", "es": "Demasiado largo. Máximo 500 caracteres.", "fr": "Trop long. 500 caractères max."},
    "msg_too_long_800": {"ru": "Слишком длинно. Максимум 800 символов.", "en": "Too long. Max 800 characters.", "de": "Zu lang. Max. 800 Zeichen.",
                         "ar": "طويل جداً. الحد الأقصى ٨٠٠ حرف.", "zh": "太长了，最多800字。", "es": "Demasiado largo. Máximo 800 caracteres.", "fr": "Trop long. 800 caractères max."},
    "msg_too_long_4000": {"ru": "Слишком длинно. Сократи до 4000 символов.", "en": "Too long. Shorten to 4000 characters.", "de": "Zu lang. Kürze auf 4000 Zeichen.",
                          "ar": "طويل جداً. اختصره إلى ٤٠٠٠ حرف.", "zh": "太长了，请缩短至4000字以内。", "es": "Demasiado largo. Reduce a 4000 caracteres.", "fr": "Trop long. Réduis à 4000 caractères."},
    "msg_too_little_text": {"ru": "Текста мало. Пришли хотя бы несколько абзацев.", "en": "Not enough text. Send at least a few paragraphs.",
                            "de": "Zu wenig Text. Sende mindestens ein paar Absätze.", "ar": "النص قليل جداً. أرسل بضع فقرات على الأقل.",
                            "zh": "文本太少，请至少发送几段。", "es": "Muy poco texto. Envía al menos varios párrafos.", "fr": "Pas assez de texte. Envoie au moins quelques paragraphes."},
    "msg_too_little_text_doc": {"ru": "Текста мало. Пришли чуть больше деталей.", "en": "Not enough text. Send a bit more detail.",
                                "de": "Zu wenig Text. Sende etwas mehr Details.", "ar": "النص قليل جداً. أرسل مزيداً من التفاصيل.",
                                "zh": "文本太少，请补充更多细节。", "es": "Muy poco texto. Envía un poco más de detalle.", "fr": "Pas assez de texte. Envoie un peu plus de détails."},
    "msg_updating_draft": {"ru": "Обновляю черновик…", "en": "Updating the draft…", "de": "Entwurf wird aktualisiert…",
                           "ar": "جاري تحديث المسودة…", "zh": "正在更新草稿…", "es": "Actualizando el borrador…", "fr": "Mise à jour du brouillon…"},
    "msg_enough_photos": {"ru": "Этого достаточно — дальше можно не присылать, нажми «Готово, собери презентацию».",
                          "en": "That's enough — no need to send more, tap \"Done, build the presentation\".",
                          "de": "Das reicht — du musst nicht mehr senden, tippe auf „Fertig, Präsentation erstellen“.",
                          "ar": "هذا يكفي — لا حاجة لإرسال المزيد، اضغط «تم، أنشئ العرض التقديمي».",
                          "zh": "这些已经够了——无需再发送，点击「完成，生成演示文稿」。",
                          "es": "Con esto es suficiente — no hace falta enviar más, toca «Listo, crea la presentación».",
                          "fr": "C'est suffisant — inutile d'en envoyer plus, appuie sur « Terminé, créez la présentation »."},
    "msg_photo_download_fail": {"ru": "Не получилось скачать это фото, пришли ещё раз или пропусти его.",
                                "en": "Couldn't download that photo, send it again or skip it.",
                                "de": "Dieses Foto konnte nicht heruntergeladen werden, sende es erneut oder überspringe es.",
                                "ar": "تعذر تنزيل هذه الصورة، أرسلها مرة أخرى أو تخطّها.",
                                "zh": "无法下载这张照片，请重新发送或跳过。",
                                "es": "No se pudo descargar esa foto, envíala de nuevo u omítela.",
                                "fr": "Impossible de télécharger cette photo, renvoie-la ou passe-la."},
    "msg_building_pres_photos": {"ru": "Собираю презентацию с картинками. Это займёт 1–2 минуты.",
                                 "en": "Building the presentation with images. This will take 1–2 minutes.",
                                 "de": "Präsentation mit Bildern wird erstellt. Das dauert 1–2 Minuten.",
                                 "ar": "جاري إنشاء العرض التقديمي بالصور. سيستغرق ١-٢ دقيقة.",
                                 "zh": "正在生成带图片的演示文稿，这需要1-2分钟。",
                                 "es": "Creando la presentación con imágenes. Esto tardará 1-2 minutos.",
                                 "fr": "Création de la présentation avec images. Cela prendra 1 à 2 minutes."},
    "msg_building_excel": {"ru": "Собираю Excel…", "en": "Building the Excel file…", "de": "Excel-Datei wird erstellt…",
                           "ar": "جاري إنشاء ملف Excel…", "zh": "正在生成Excel文件…", "es": "Creando el archivo Excel…", "fr": "Création du fichier Excel…"},
    "msg_building_word": {"ru": "Собираю Word. Обычно это быстрее презентации.", "en": "Building the Word file. This is usually faster than a presentation.",
                          "de": "Word-Datei wird erstellt. Das geht meist schneller als eine Präsentation.",
                          "ar": "جاري إنشاء ملف Word. عادة ما يكون أسرع من العرض التقديمي.",
                          "zh": "正在生成Word文件，通常比演示文稿更快。",
                          "es": "Creando el archivo Word. Suele ser más rápido que una presentación.",
                          "fr": "Création du fichier Word. C'est généralement plus rapide qu'une présentation."},
    "msg_history_empty": {"ru": "История пустая.", "en": "History is empty.", "de": "Verlauf ist leer.",
                          "ar": "السجل فارغ.", "zh": "历史记录为空。", "es": "El historial está vacío.", "fr": "L'historique est vide."},
    "msg_history_title": {"ru": "📁 История:\n\n", "en": "📁 History:\n\n", "de": "📁 Verlauf:\n\n",
                          "ar": "📁 السجل:\n\n", "zh": "📁 历史记录：\n\n", "es": "📁 Historial:\n\n", "fr": "📁 Historique :\n\n"},
    "msg_start_pres_again": {"ru": "Сделать презентацию", "en": "Create a presentation", "de": "Präsentation erstellen",
                            "ar": "إنشاء عرض تقديمي", "zh": "制作演示文稿", "es": "Crear una presentación", "fr": "Créer une présentation"},
    "msg_how_build_pres": {"ru": "Как делаем презентацию?\n\n✨ Сгенерировать с ИИ — я сам придумаю текст, стиль и фото.\n📝 Вставить свой текст — пришли материал, я поправлю ошибки и соберу слайды.",
                           "en": "How should we build the presentation?\n\n✨ Generate with AI — I'll come up with the text, style and photos myself.\n📝 Paste your own text — send the material, I'll fix errors and build the slides.",
                           "de": "Wie erstellen wir die Präsentation?\n\n✨ Mit KI generieren — ich denke mir Text, Stil und Fotos selbst aus.\n📝 Eigenen Text einfügen — sende das Material, ich korrigiere Fehler und erstelle die Folien.",
                           "ar": "كيف ننشئ العرض التقديمي؟\n\n✨ إنشاء بالذكاء الاصطناعي — سأبتكر النص والنمط والصور بنفسي.\n📝 إدراج نصك الخاص — أرسل المادة وسأصحح الأخطاء وأبني الشرائح.",
                           "zh": "如何制作演示文稿？\n\n✨ 用AI生成——由我构思文本、风格和图片。\n📝 粘贴自己的文本——发送素材，我会修正错误并制作幻灯片。",
                           "es": "¿Cómo creamos la presentación?\n\n✨ Generar con IA — yo mismo pensaré el texto, el estilo y las fotos.\n📝 Pegar tu propio texto — envía el material, corregiré errores y armaré las diapositivas.",
                           "fr": "Comment créons-nous la présentation ?\n\n✨ Générer avec l'IA — j'imaginerai moi-même le texte, le style et les photos.\n📝 Coller votre propre texte — envoyez le contenu, je corrigerai les erreurs et je construirai les diapositives."},
    "msg_style_fits_topic": {"ru": "🎨 По теме подходит стиль: {style}.\nОставить его или выбрать другой?",
                             "en": "🎨 This style fits the topic: {style}.\nKeep it or choose another one?",
                             "de": "🎨 Zum Thema passt der Stil: {style}.\nBehalten oder einen anderen wählen?",
                             "ar": "🎨 النمط المناسب للموضوع: {style}.\nهل تحتفظ به أم تختار غيره؟",
                             "zh": "🎨 适合该主题的风格：{style}。\n保留还是选择其他？",
                             "es": "🎨 Para este tema encaja el estilo: {style}.\n¿Lo mantenemos o eliges otro?",
                             "fr": "🎨 Ce style convient au sujet : {style}.\nOn le garde ou tu en choisis un autre ?"},
    "msg_style_fits_text": {"ru": "🎨 По тексту подходит стиль: {style}.\nОставить его или выбрать другой?",
                            "en": "🎨 This style fits the text: {style}.\nKeep it or choose another one?",
                            "de": "🎨 Zum Text passt der Stil: {style}.\nBehalten oder einen anderen wählen?",
                            "ar": "🎨 النمط المناسب للنص: {style}.\nهل تحتفظ به أم تختار غيره؟",
                            "zh": "🎨 适合该文本的风格：{style}。\n保留还是选择其他？",
                            "es": "🎨 Para este texto encaja el estilo: {style}.\n¿Lo mantenemos o eliges otro?",
                            "fr": "🎨 Ce style convient au texte : {style}.\nOn le garde ou tu en choisis un autre ?"},
    "msg_grok_error": {"ru": "Не получилось получить ответ от нейросети. Попробуй ещё раз через минуту.",
                       "en": "Couldn't get a response from the AI. Try again in a minute.",
                       "de": "Keine Antwort von der KI erhalten. Versuche es in einer Minute erneut.",
                       "ar": "تعذر الحصول على رد من الذكاء الاصطناعي. حاول مرة أخرى بعد دقيقة.",
                       "zh": "未能获得AI的回复，请一分钟后再试。",
                       "es": "No se pudo obtener respuesta de la IA. Inténtalo de nuevo en un minuto.",
                       "fr": "Impossible d'obtenir une réponse de l'IA. Réessaie dans une minute."},
    "msg_draft_ready_pres": {"ru": "Черновик готов ✅\n\n{sample}\n\nСтиль: {style}\n\nЕсли всё ок — собираем полную версию.",
                             "en": "Draft ready ✅\n\n{sample}\n\nStyle: {style}\n\nIf it looks good, let's build the full version.",
                             "de": "Entwurf fertig ✅\n\n{sample}\n\nStil: {style}\n\nWenn alles passt, erstellen wir die vollständige Version.",
                             "ar": "المسودة جاهزة ✅\n\n{sample}\n\nالنمط: {style}\n\nإذا كان كل شيء جيداً، دعنا ننشئ النسخة الكاملة.",
                             "zh": "草稿已完成 ✅\n\n{sample}\n\n风格：{style}\n\n如果没问题，我们来生成完整版本。",
                             "es": "Borrador listo ✅\n\n{sample}\n\nEstilo: {style}\n\nSi todo está bien, creemos la versión completa.",
                             "fr": "Brouillon prêt ✅\n\n{sample}\n\nStyle : {style}\n\nSi tout va bien, on crée la version complète."},
    "msg_style_changed": {"ru": "Стиль изменён: {style}\n\nЧерновик:\n\n{sample}\n\nВыбери действие:",
                          "en": "Style changed: {style}\n\nDraft:\n\n{sample}\n\nChoose an action:",
                          "de": "Stil geändert: {style}\n\nEntwurf:\n\n{sample}\n\nWähle eine Aktion:",
                          "ar": "تم تغيير النمط: {style}\n\nالمسودة:\n\n{sample}\n\nاختر إجراءً:",
                          "zh": "风格已更改：{style}\n\n草稿：\n\n{sample}\n\n选择操作：",
                          "es": "Estilo cambiado: {style}\n\nBorrador:\n\n{sample}\n\nElige una acción:",
                          "fr": "Style modifié : {style}\n\nBrouillon :\n\n{sample}\n\nChoisis une action :"},
    "msg_draft_updated": {"ru": "Обновлённый черновик ✅\n\n{sample}\n\nВыбери действие:",
                          "en": "Updated draft ✅\n\n{sample}\n\nChoose an action:",
                          "de": "Aktualisierter Entwurf ✅\n\n{sample}\n\nWähle eine Aktion:",
                          "ar": "المسودة المحدّثة ✅\n\n{sample}\n\nاختر إجراءً:",
                          "zh": "更新后的草稿 ✅\n\n{sample}\n\n选择操作：",
                          "es": "Borrador actualizado ✅\n\n{sample}\n\nElige una acción:",
                          "fr": "Brouillon mis à jour ✅\n\n{sample}\n\nChoisis une action :"},
    "msg_draft_update_failed": {"ru": "Не получилось обновить черновик — нейросеть не ответила. Прошлый черновик остался как был.",
                                "en": "Couldn't update the draft — the AI didn't respond. The previous draft stayed as it was.",
                                "de": "Entwurf konnte nicht aktualisiert werden — die KI hat nicht geantwortet. Der vorherige Entwurf blieb unverändert.",
                                "ar": "تعذر تحديث المسودة — لم يرد الذكاء الاصطناعي. بقيت المسودة السابقة كما كانت.",
                                "zh": "未能更新草稿——AI没有响应。之前的草稿保持不变。",
                                "es": "No se pudo actualizar el borrador — la IA no respondió. El borrador anterior se mantuvo igual.",
                                "fr": "Impossible de mettre à jour le brouillon — l'IA n'a pas répondu. Le brouillon précédent est resté tel quel."},
    "msg_photo_or_own": {"ru": "Фото для слайдов сделать через ИИ или использовать свои?\nЕсли фото будет меньше, чем слайдов — оставшиеся бот дорисует сам.",
                         "en": "Should slide photos be AI-generated or your own?\nIf you send fewer photos than slides, the bot will generate the rest.",
                         "de": "Fotos für die Folien per KI erstellen oder eigene verwenden?\nWenn du weniger Fotos als Folien sendest, generiert der Bot den Rest selbst.",
                         "ar": "هل تُنشأ صور الشرائح بالذكاء الاصطناعي أم تستخدم صورك الخاصة؟\nإذا كانت الصور أقل من الشرائح، سيكمل البوت الباقي تلقائياً.",
                         "zh": "幻灯片图片用AI生成还是使用你自己的照片？\n如果照片数量少于幻灯片数，机器人会自动补全其余部分。",
                         "es": "¿Las fotos de las diapositivas se generan con IA o usas las tuyas?\nSi envías menos fotos que diapositivas, el bot generará el resto.",
                         "fr": "Les photos des diapositives doivent-elles être générées par l'IA ou tes propres photos ?\nSi tu envoies moins de photos que de diapositives, le bot complétera le reste."},
    "msg_send_photos_one_by_one": {"ru": "Пришли фото по одному (можно несколько сообщений подряд). Когда закончишь — нажми «Готово, собери презентацию».",
                                   "en": "Send photos one by one (several messages in a row is fine). When done, tap \"Done, build the presentation\".",
                                   "de": "Sende Fotos einzeln (mehrere Nachrichten hintereinander sind ok). Wenn fertig, tippe auf „Fertig, Präsentation erstellen“.",
                                   "ar": "أرسل الصور واحدة تلو الأخرى (يمكن عدة رسائل متتالية). عند الانتهاء، اضغط «تم، أنشئ العرض التقديمي».",
                                   "zh": "请逐张发送照片（可以连续发送多条消息）。完成后点击「完成，生成演示文稿」。",
                                   "es": "Envía las fotos una por una (varios mensajes seguidos está bien). Cuando termines, toca «Listo, crea la presentación».",
                                   "fr": "Envoie les photos une par une (plusieurs messages à la suite, c'est possible). Une fois terminé, appuie sur « Terminé, créez la présentation »."},
    "msg_photo_received": {"ru": "Фото {n} получено 📷 Пришли ещё или нажми «Готово, собери презентацию».",
                           "en": "Photo {n} received 📷 Send more or tap \"Done, build the presentation\".",
                           "de": "Foto {n} erhalten 📷 Sende weitere oder tippe auf „Fertig, Präsentation erstellen“.",
                           "ar": "تم استلام الصورة {n} 📷 أرسل المزيد أو اضغط «تم، أنشئ العرض التقديمي».",
                           "zh": "已收到第{n}张照片📷，可继续发送，或点击「完成，生成演示文稿」。",
                           "es": "Foto {n} recibida 📷 Envía más o toca «Listo, crea la presentación».",
                           "fr": "Photo {n} reçue 📷 Envoie-en d'autres ou appuie sur « Terminé, créez la présentation »."},
    "msg_startup_data_prompt": {
        "ru": "Пришли исходные данные для расчёта одним сообщением:\n\n• стартовые вложения\n• ежемесячные расходы списком (аренда, зарплаты, реклама и т.д. — с суммами)\n• цена за единицу товара/услуги\n• ожидаемое количество продаж в месяц\n• на сколько месяцев считать\n\nНапример: вложения 500 000 ₽, аренда 40 000, зарплаты 90 000, реклама 25 000, цена 350 ₽, продажи 600 шт/мес, считать на 12 месяцев.",
        "en": "Send the source data for the calculation in one message:\n\n• starting investment\n• monthly expenses as a list (rent, salaries, ads, etc. — with amounts)\n• price per unit of product/service\n• expected number of sales per month\n• how many months to calculate for\n\nExample: investment $5,000, rent $400, salaries $900, ads $250, price $3.5, sales 600 units/month, calculate for 12 months.",
        "de": "Sende die Ausgangsdaten für die Berechnung in einer Nachricht:\n\n• Startinvestition\n• monatliche Ausgaben als Liste (Miete, Gehälter, Werbung usw. — mit Beträgen)\n• Preis pro Einheit Produkt/Dienstleistung\n• erwartete Verkaufszahl pro Monat\n• für wie viele Monate gerechnet werden soll\n\nBeispiel: Investition 5.000 €, Miete 400 €, Gehälter 900 €, Werbung 250 €, Preis 3,5 €, Verkäufe 600 Stk/Monat, Berechnung für 12 Monate.",
        "ar": "أرسل البيانات الأصلية للحساب في رسالة واحدة:\n\n• الاستثمار الأولي\n• المصاريف الشهرية كقائمة (الإيجار، الرواتب، الإعلانات، إلخ — مع المبالغ)\n• سعر الوحدة من المنتج/الخدمة\n• عدد المبيعات المتوقع شهرياً\n• عدد الأشهر للحساب\n\nمثال: استثمار 5000$، إيجار 400$، رواتب 900$، إعلانات 250$، سعر 3.5$، مبيعات 600 وحدة/شهر، الحساب لمدة 12 شهراً.",
        "zh": "请一次性发送以下计算所需数据：\n\n• 初始投资\n• 月度支出清单（租金、工资、广告等——附金额）\n• 每单位产品/服务价格\n• 预期月销量\n• 计算周期（月数）\n\n示例：投资5000美元，租金400，工资900，广告250，单价3.5，月销量600件，按12个月计算。",
        "es": "Envía los datos para el cálculo en un solo mensaje:\n\n• inversión inicial\n• gastos mensuales en lista (alquiler, salarios, publicidad, etc. — con montos)\n• precio por unidad de producto/servicio\n• ventas esperadas por mes\n• para cuántos meses calcular\n\nEjemplo: inversión $5000, alquiler $400, salarios $900, publicidad $250, precio $3.5, ventas 600 uds/mes, calcular para 12 meses.",
        "fr": "Envoie les données source pour le calcul en un seul message :\n\n• investissement initial\n• dépenses mensuelles sous forme de liste (loyer, salaires, publicité, etc. — avec montants)\n• prix par unité de produit/service\n• nombre de ventes attendu par mois\n• sur combien de mois calculer\n\nExemple : investissement 5000 €, loyer 400 €, salaires 900 €, publicité 250 €, prix 3,5 €, ventes 600 unités/mois, calculer sur 12 mois."},
    "msg_how_build_table": {
        "ru": "Как собираем таблицу?\n\n✨ Сгенерировать с ИИ — сам придумаю правдоподобные данные по теме.\n📝 Ввести свои данные — пришли реальные цифры, я оформлю их в таблицу с формулами.",
        "en": "How should we build the table?\n\n✨ Generate with AI — I'll come up with plausible data on the topic.\n📝 Enter your own data — send real numbers, I'll format them into a table with formulas.",
        "de": "Wie erstellen wir die Tabelle?\n\n✨ Mit KI generieren — ich denke mir plausible Daten zum Thema aus.\n📝 Eigene Daten eingeben — sende reale Zahlen, ich formatiere sie in eine Tabelle mit Formeln.",
        "ar": "كيف ننشئ الجدول؟\n\n✨ إنشاء بالذكاء الاصطناعي — سأبتكر بيانات معقولة حول الموضوع.\n📝 إدخال بياناتك الخاصة — أرسل أرقاماً حقيقية وسأنسقها في جدول بمعادلات.",
        "zh": "如何生成表格？\n\n✨ 用AI生成——由我构思与主题相符的合理数据。\n📝 输入自己的数据——发送真实数字，我会将其整理成带公式的表格。",
        "es": "¿Cómo creamos la tabla?\n\n✨ Generar con IA — inventaré datos plausibles sobre el tema.\n📝 Introducir tus propios datos — envía cifras reales, las daré formato en una tabla con fórmulas.",
        "fr": "Comment créons-nous le tableau ?\n\n✨ Générer avec l'IA — j'imaginerai des données plausibles sur le sujet.\n📝 Saisir tes propres données — envoie des chiffres réels, je les mettrai en forme dans un tableau avec des formules."},
    "msg_table_ready": {"ru": "Таблица готова ✅", "en": "Table ready ✅", "de": "Tabelle fertig ✅",
                        "ar": "الجدول جاهز ✅", "zh": "表格已完成 ✅", "es": "Tabla lista ✅", "fr": "Tableau prêt ✅"},
    "msg_check_source_numbers": {"ru": "\n\n⚠️ Проверь исходные цифры — модель только оформила то, что ты прислал, ничего не добавляла от себя.",
                                 "en": "\n\n⚠️ Check the source numbers — the model only formatted what you sent, it didn't add anything on its own.",
                                 "de": "\n\n⚠️ Überprüfe die Ausgangszahlen — das Modell hat nur formatiert, was du gesendet hast, ohne etwas hinzuzufügen.",
                                 "ar": "\n\n⚠️ تحقق من الأرقام الأصلية — النموذج قام فقط بتنسيق ما أرسلته دون إضافة أي شيء من عنده.",
                                 "zh": "\n\n⚠️ 请核对原始数字——模型只是整理了你提供的内容，没有自行添加任何数据。",
                                 "es": "\n\n⚠️ Revisa las cifras originales — el modelo solo dio formato a lo que enviaste, sin añadir nada por su cuenta.",
                                 "fr": "\n\n⚠️ Vérifie les chiffres d'origine — le modèle s'est contenté de mettre en forme ce que tu as envoyé, sans rien ajouter de lui-même."},
    "msg_excel_caption": {"ru": "📈 Excel — с формулами, можно редактировать", "en": "📈 Excel — with formulas, editable",
                          "de": "📈 Excel — mit Formeln, bearbeitbar", "ar": "📈 Excel — بمعادلات، قابل للتعديل",
                          "zh": "📈 Excel——带公式，可编辑", "es": "📈 Excel — con fórmulas, editable", "fr": "📈 Excel — avec formules, modifiable"},
    "msg_template_caption": {"ru": "📄 Пустой шаблон — заполни пропуски от руки или в Word и распечатай",
                             "en": "📄 Blank template — fill in the gaps by hand or in Word and print",
                             "de": "📄 Leere Vorlage — fülle die Lücken handschriftlich oder in Word aus und drucke sie",
                             "ar": "📄 قالب فارغ — املأ الفراغات يدوياً أو في Word ثم اطبعه",
                             "zh": "📄 空白模板——手动或在Word中填写空白处后打印",
                             "es": "📄 Plantilla en blanco — rellena los espacios a mano o en Word e imprime",
                             "fr": "📄 Modèle vierge — remplis les champs à la main ou dans Word puis imprime"},
    "msg_draft_ready_doc": {"ru": "Черновик готов ✅\n\n{sample}\n\nЕсли всё ок — собираем файл.",
                            "en": "Draft ready ✅\n\n{sample}\n\nIf it looks good, let's build the file.",
                            "de": "Entwurf fertig ✅\n\n{sample}\n\nWenn alles passt, erstellen wir die Datei.",
                            "ar": "المسودة جاهزة ✅\n\n{sample}\n\nإذا كان كل شيء جيداً، دعنا ننشئ الملف.",
                            "zh": "草稿已完成 ✅\n\n{sample}\n\n如果没问题，我们来生成文件。",
                            "es": "Borrador listo ✅\n\n{sample}\n\nSi todo está bien, creemos el archivo.",
                            "fr": "Brouillon prêt ✅\n\n{sample}\n\nSi tout va bien, on crée le fichier."},
    "msg_writing_more": {"ru": "Черновик получился короче, чем нужно — дописываю подробнее…",
                         "en": "The draft came out shorter than needed — expanding it…",
                         "de": "Der Entwurf ist kürzer als nötig ausgefallen — ich erweitere ihn…",
                         "ar": "المسودة أقصر من اللازم — جاري إطالتها…",
                         "zh": "草稿比预期短——正在补充完善…",
                         "es": "El borrador quedó más corto de lo necesario — ampliándolo…",
                         "fr": "Le brouillon est plus court que nécessaire — je le développe…"},
    "msg_writing_a_bit_more": {"ru": "Ещё чуть-чуть осталось — дописываю…",
                               "en": "Almost there — still expanding…",
                               "de": "Fast fertig — ich schreibe noch etwas…",
                               "ar": "اقتربنا — ما زلت أكتب…",
                               "zh": "快好了——继续补充中…",
                               "es": "Ya casi está — sigo ampliando…",
                               "fr": "Presque terminé — je continue à développer…"},
    "msg_word_caption": {"ru": "📄 Word — этот файл можно править", "en": "📄 Word — this file can be edited",
                         "de": "📄 Word — diese Datei kann bearbeitet werden", "ar": "📄 Word — يمكن تعديل هذا الملف",
                         "zh": "📄 Word——此文件可编辑", "es": "📄 Word — este archivo se puede editar", "fr": "📄 Word — ce fichier est modifiable"},
    "msg_pdf_caption": {"ru": "📄 PDF-копия", "en": "📄 PDF copy", "de": "📄 PDF-Kopie",
                        "ar": "📄 نسخة PDF", "zh": "📄 PDF副本", "es": "📄 Copia en PDF", "fr": "📄 Copie PDF"},
    "msg_doc_ready": {"ru": "Документ готов ✅", "en": "Document ready ✅", "de": "Dokument fertig ✅",
                      "ar": "المستند جاهز ✅", "zh": "文档已完成 ✅", "es": "Documento listo ✅", "fr": "Document prêt ✅"},
    "msg_check_gaps": {"ru": "\n\n⚠️ Проверь пропуски в тексте перед распечаткой.",
                       "en": "\n\n⚠️ Check the blanks in the text before printing.",
                       "de": "\n\n⚠️ Überprüfe die Lücken im Text vor dem Drucken.",
                       "ar": "\n\n⚠️ تحقق من الفراغات في النص قبل الطباعة.",
                       "zh": "\n\n⚠️ 打印前请检查文中的空白处。",
                       "es": "\n\n⚠️ Revisa los espacios en blanco del texto antes de imprimir.",
                       "fr": "\n\n⚠️ Vérifie les champs vides dans le texte avant d'imprimer."},
    "msg_didnt_understand": {"ru": "Не понял. Выбери действие кнопкой ниже, или напиши /cancel, чтобы начать заново.",
                             "en": "Didn't catch that. Choose an action with the button below, or type /cancel to start over.",
                             "de": "Das habe ich nicht verstanden. Wähle eine Aktion mit dem Button unten oder schreibe /cancel, um neu zu beginnen.",
                             "ar": "لم أفهم. اختر إجراءً من الزر أدناه، أو اكتب /cancel للبدء من جديد.",
                             "zh": "没能理解。请用下方按钮选择操作，或输入 /cancel 重新开始。",
                             "es": "No entendí. Elige una acción con el botón de abajo, o escribe /cancel para empezar de nuevo.",
                             "fr": "Je n'ai pas compris. Choisis une action avec le bouton ci-dessous, ou tape /cancel pour recommencer."},
    "msg_plan_info": {"ru": "ℹ️ Генераций: {used} из {limit}\nОсталось: {left}",
                      "en": "ℹ️ Generations: {used} of {limit}\nRemaining: {left}",
                      "de": "ℹ️ Generierungen: {used} von {limit}\nVerbleibend: {left}",
                      "ar": "ℹ️ التوليدات: {used} من {limit}\nالمتبقي: {left}",
                      "zh": "ℹ️ 生成次数：{used} / {limit}\n剩余：{left}",
                      "es": "ℹ️ Generaciones: {used} de {limit}\nRestantes: {left}",
                      "fr": "ℹ️ Générations : {used} sur {limit}\nRestant : {left}"},
}


def tr(key, lang, **kw):
    """Перевод интерфейсной строки по ключу. Если для языка нет перевода - откатывается
    на русский, чтобы бот никогда не показал пустоту/ошибку из-за пропуска в словаре."""
    d = TR.get(key, {})
    s = d.get(lang) or d.get("ru") or key
    return s.format(**kw) if kw else s


# Множества всех переводов конкретных служебных кнопок - чтобы F.text.in_(...) фильтры
# в хендлерах узнавали нажатие кнопки независимо от текущего языка пользователя
# (кнопка "🏠 Главное меню" на любом из 7 языков должна сработать одинаково).
ALL_MAIN_MENU_LABELS = set(TR["btn_main_menu"].values())
ALL_MORE_DOCS_LABELS = set(TR["btn_more_docs"].values())
ALL_BACK_CATEGORIES_LABELS = set(TR["btn_back_categories"].values())
ALL_BTN_PRES_LABELS = set(TR["btn_pres"].values())
ALL_BTN_WORD_LABELS = set(TR["btn_word"].values())
ALL_BTN_EXCEL_LABELS = set(TR["btn_excel"].values())
ALL_BTN_HISTORY_LABELS = set(TR["btn_history"].values())
ALL_BTN_PLAN_LABELS = set(TR["btn_plan"].values())
ALL_BTN_AI_GENERATE_LABELS = set(TR["btn_ai_generate"].values())
ALL_BTN_OWN_TEXT_LABELS = set(TR["btn_own_text"].values())
ALL_BTN_OWN_DATA_LABELS = set(TR["btn_own_data"].values())
ALL_BTN_CHANGE_TOPIC_LABELS = set(TR["btn_change_topic"].values())
ALL_BTN_CHANGE_STYLE_LABELS = set(TR["btn_change_style"].values())
ALL_BTN_ADD_INFO_LABELS = set(TR["btn_add_info"].values())
ALL_BTN_FULL_VERSION_LABELS = set(TR["btn_full_version"].values()) | {"делай", "да", "ок", "yes", "ok"}
ALL_BTN_AI_PHOTOS_LABELS = set(TR["btn_ai_photos"].values())
ALL_BTN_OWN_PHOTOS_LABELS = set(TR["btn_own_photos"].values())
ALL_BTN_PHOTOS_DONE_LABELS = set(TR["btn_photos_done"].values())
ALL_BTN_CHANGE_QUERY_LABELS = set(TR["btn_change_query"].values())
ALL_BTN_BUILD_TABLE_LABELS = set(TR["btn_build_table"].values()) | {"делай", "да", "ок", "yes", "ok"}
ALL_BTN_BUILD_DOC_LABELS = set(TR["btn_build_doc"].values()) | {"делай", "да", "ок", "yes", "ok"}
ALL_BTN_TEMPLATE_LABELS = set(TR["btn_template"].values())


def lang_kb():
    rows = []
    codes = list(LANGS.keys())
    for i in range(0, len(codes), 2):
        pair = codes[i:i + 2]
        rows.append([KeyboardButton(text=f"{LANGS[c]['flag']} {LANGS[c]['name']}") for c in pair])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


LANG_LABEL_TO_CODE = {f"{v['flag']} {v['name']}": k for k, v in LANGS.items()}


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

THEME_LABELS_I18N = {
    "nature": {"ru": "Природа", "en": "Nature", "de": "Natur", "ar": "الطبيعة", "zh": "自然", "es": "Naturaleza", "fr": "Nature"},
    "business": {"ru": "Бизнес", "en": "Business", "de": "Business", "ar": "الأعمال", "zh": "商业", "es": "Negocios", "fr": "Affaires"},
    "tech": {"ru": "Технологии", "en": "Technology", "de": "Technologie", "ar": "التقنية", "zh": "科技", "es": "Tecnología", "fr": "Technologie"},
    "school": {"ru": "Учёба", "en": "Study", "de": "Studium", "ar": "الدراسة", "zh": "学习", "es": "Estudios", "fr": "Études"},
    "fashion": {"ru": "Мода", "en": "Fashion", "de": "Mode", "ar": "الموضة", "zh": "时尚", "es": "Moda", "fr": "Mode"},
    "history": {"ru": "История", "en": "History", "de": "Geschichte", "ar": "التاريخ", "zh": "历史", "es": "Historia", "fr": "Histoire"},
    "science": {"ru": "Наука", "en": "Science", "de": "Wissenschaft", "ar": "العلوم", "zh": "科学", "es": "Ciencia", "fr": "Science"},
    "sport": {"ru": "Спорт", "en": "Sport", "de": "Sport", "ar": "الرياضة", "zh": "体育", "es": "Deporte", "fr": "Sport"},
    "travel": {"ru": "Путешествия", "en": "Travel", "de": "Reisen", "ar": "السفر", "zh": "旅行", "es": "Viajes", "fr": "Voyage"},
    "food": {"ru": "Еда", "en": "Food", "de": "Essen", "ar": "الطعام", "zh": "美食", "es": "Comida", "fr": "Gastronomie"},
    "art": {"ru": "Искусство", "en": "Art", "de": "Kunst", "ar": "الفن", "zh": "艺术", "es": "Arte", "fr": "Art"},
    "eco": {"ru": "Экология", "en": "Ecology", "de": "Ökologie", "ar": "البيئة", "zh": "生态", "es": "Ecología", "fr": "Écologie"},
    "minimal": {"ru": "Минимализм", "en": "Minimalism", "de": "Minimalismus", "ar": "البساطة", "zh": "极简", "es": "Minimalismo", "fr": "Minimalisme"},
    "default": {"ru": "Универсальный", "en": "Universal", "de": "Universell", "ar": "عام", "zh": "通用", "es": "Universal", "fr": "Universel"},
}
THEME_LABELS = {k: v["ru"] for k, v in THEME_LABELS_I18N.items()}  # обратная совместимость

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
        users_db[uid] = {"name": "", "plan": "premium", "generations": 0, "history": [], "busy": False,
                          "lang": "ru", "lang_chosen": False}
    return users_db[uid]


def user_lang(uid):
    return get_user(uid).get("lang", "ru")


def can_generate(uid):
    u = get_user(uid)
    return u["generations"] < PLAN_LIMITS.get(u["plan"], 15)


# --- Защита от флуда и DDoS ------------------------------------------------
# Три уровня, все проверяются в start_job() — единой точке входа перед любой
# тяжёлой генерацией (презентация/Word/Excel/шаблон):
#   1) rate-limit по времени на одного пользователя — не чаще раза в
#      RATE_LIMIT_SECONDS, независимо от того, завершилась ли предыдущая
#      генерация (per-uid busy-флаг это не покрывал);
#   2) глобальный потолок одновременных генераций по всему боту сразу —
#      защищает от атаки с множества разных Telegram-аккаунтов;
#   3) алерт админу при аномальной нагрузке за скользящий час — суточного
#      лимита нет (по запросу), но при подозрительном всплеске бот сам
#      предупредит владельца, чтобы среагировать вручную до того, как
#      выгорит баланс на xAI.
RATE_LIMIT_SECONDS = 30
MAX_CONCURRENT_GENERATIONS = 20
HOURLY_ALERT_THRESHOLD = 50
ALERT_COOLDOWN_SECONDS = 900  # не чаще раза в 15 минут, чтобы не заспамить одним и тем же

_last_request_time = {}          # uid -> monotonic-время последнего успешного старта генерации
_generation_timestamps = deque() # monotonic-метки всех начатых генераций за последний час
_active_generations = 0          # сколько генераций сейчас реально выполняется одновременно
_last_alert_time = float("-inf") # когда в последний раз слали алерт админу.
# ВАЖНО: именно -inf, а не 0.0 — time.monotonic() отсчитывается от произвольной
# точки (не от старта процесса и не от эпохи), и если бы тут был 0.0, проверка
# "now - _last_alert_time > ALERT_COOLDOWN_SECONDS" в первые ~15 минут после
# старта бота была бы ложно False (потому что now ещё меньше 900), и самый
# первый алерт об атаке — то есть ровно тот момент, когда бот только что
# перезапустили и он особенно уязвим — просто не дошёл бы до админа.


def _check_hourly_load():
    """Чистит метки старше часа и, если нагрузка аномально высокая, шлёт
    алерт админу (не чаще раза в ALERT_COOLDOWN_SECONDS)."""
    global _last_alert_time
    now = time.monotonic()
    while _generation_timestamps and now - _generation_timestamps[0] > 3600:
        _generation_timestamps.popleft()
    if len(_generation_timestamps) >= HOURLY_ALERT_THRESHOLD and now - _last_alert_time > ALERT_COOLDOWN_SECONDS:
        _last_alert_time = now
        count = len(_generation_timestamps)
        for admin_id in ADMIN_IDS:
            asyncio.create_task(bot.send_message(
                admin_id,
                f"⚠️ Аномальная нагрузка на бота: {count} генераций за последний час.\n"
                f"Похоже на флуд или атаку — стоит проверить."
            ))


def start_job(uid):
    """Помечает пользователя как занятого дорогой генерацией и проверяет
    защиту от флуда/DDoS (rate-limit + глобальный потолок).
    Возвращает (True, None) при успехе, иначе (False, "текст для пользователя")."""
    global _active_generations
    u = get_user(uid)

    if u.get("busy"):
        return False, "Уже собираю предыдущую версию, подожди немного 🙂"

    now = time.monotonic()
    last = _last_request_time.get(uid, 0)
    if now - last < RATE_LIMIT_SECONDS:
        wait = int(RATE_LIMIT_SECONDS - (now - last)) + 1
        return False, f"Слишком часто 🙂 Подожди ещё {wait} сек. и попробуй снова."

    if _active_generations >= MAX_CONCURRENT_GENERATIONS:
        return False, "Сейчас бот перегружен — очень много запросов одновременно. Попробуй, пожалуйста, через минуту 🙏"

    u["busy"] = True
    u["_counted"] = True
    _last_request_time[uid] = now
    _active_generations += 1
    _generation_timestamps.append(now)
    _check_hourly_load()
    return True, None


def finish_job(uid):
    """Снимает занятость пользователя. Идемпотентна: если по этому uid уже
    финишировали (например, /cancel сработал раньше, чем реально завершилась
    генерация), счётчик _active_generations не уйдёт в минус и не спишется
    дважды за одну и ту же генерацию."""
    global _active_generations
    u = get_user(uid)
    if u.get("_counted"):
        _active_generations = max(0, _active_generations - 1)
        u["_counted"] = False
    u["busy"] = False


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

GROK_LANG_HINT = {
    "ru": "", "en": "\n\nВАЖНО: ответь полностью на английском языке (English).",
    "de": "\n\nВАЖНО: ответь полностью на немецком языке (Deutsch).",
    "ar": "\n\nВАЖНО: ответь полностью на арабском языке (العربية).",
    "zh": "\n\nВАЖНО: ответь полностью на китайском языке (中文).",
    "es": "\n\nВАЖНО: ответь полностью на испанском языке (Español).",
    "fr": "\n\nВАЖНО: ответь полностью на французском языке (Français).",
}


def grok_lang_instruction(lang: str) -> str:
    """Суффикс для промпта Grok, требующий ответ на языке пользователя. Для ru — пусто."""
    return GROK_LANG_HINT.get(lang, "")


GROK_JSON_LANG_HINT = {
    "ru": "",
    "en": "\n\nВАЖНО: весь видимый пользователю текст в JSON (title, content, любые названия, подписи) пиши на английском языке (English). Поля image_prompt оставляй на английском в любом случае (это для генерации картинок).",
    "de": "\n\nВАЖНО: весь видимый пользователю текст в JSON (title, content, любые названия, подписи) пиши на немецком языке (Deutsch). Поля image_prompt оставляй на английском в любом случае (это для генерации картинок).",
    "ar": "\n\nВАЖНО: весь видимый пользователю текст в JSON (title, content, любые названия, подписи) пиши на арабском языке (العربية). Поля image_prompt оставляй на английском в любом случае (это для генерации картинок).",
    "zh": "\n\nВАЖНО: весь видимый пользователю текст в JSON (title, content, любые названия, подписи) пиши на китайском языке (中文). Поля image_prompt оставляй на английском в любом случае (это для генерации картинок).",
    "es": "\n\nВАЖНО: весь видимый пользователю текст в JSON (title, content, любые названия, подписи) пиши на испанском языке (Español). Поля image_prompt оставляй на английском в любом случае (это для генерации картинок).",
    "fr": "\n\nВАЖНО: весь видимый пользователю текст в JSON (title, content, любые названия, подписи) пиши на французском языке (Français). Поля image_prompt оставляй на английском в любом случае (это для генерации картинок).",
}


def grok_json_lang_instruction(lang: str) -> str:
    """Как grok_lang_instruction, но уточняет что image_prompt должен остаться на английском (для генерации картинок)."""
    return GROK_JSON_LANG_HINT.get(lang, "")


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


def add_chart(slide, l, t, w, h, chart_data_dict, colors):
    """Рисует нативный график PowerPoint (bar/line/pie) на слайде вместо фото.
    chart_data_dict: {"chart_type": "bar"/"line"/"pie", "title": "...",
    "categories": [...], "series": [{"name":..., "values":[...]}]}.
    Используется только когда модель сама решила, что тема подразумевает цифры -
    см. инструкцию про ключ "chart" в промпте _build_presentation."""
    ctype = chart_data_dict.get("chart_type", "bar")
    cats = chart_data_dict.get("categories") or []
    series_list = chart_data_dict.get("series") or []
    if not cats or not series_list:
        return None

    cd = CategoryChartData()
    cd.categories = cats
    for s in series_list:
        cd.add_series(s.get("name", ""), s.get("values", []))

    xl_type = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE,
    }.get(ctype, XL_CHART_TYPE.COLUMN_CLUSTERED)

    graphic_frame = slide.shapes.add_chart(xl_type, Inches(l), Inches(t), Inches(w), Inches(h), cd)
    chart = graphic_frame.chart

    # чем уже панель под график - тем мельче шрифты, иначе легенда/подписи не влезают
    compact = w < 5.0
    title_sz, legend_sz, label_sz, tick_sz = (13, 9, 9, 9) if compact else (15, 11, 11, 10)

    palette = [colors["line"], colors["mid"], colors["mute"], colors["ink"]]

    if chart_data_dict.get("title"):
        chart.has_title = True
        chart.chart_title.text_frame.text = chart_data_dict["title"]
        for run in chart.chart_title.text_frame.paragraphs[0].runs:
            run.font.size = Pt(title_sz)
            run.font.bold = True
            run.font.color.rgb = RGBColor(*colors["ink"])
            run.font.name = "Calibri"
    else:
        chart.has_title = False

    is_pie = ctype == "pie"
    multi_series = len(series_list) > 1

    if is_pie or multi_series:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(legend_sz)
        chart.legend.font.color.rgb = RGBColor(*colors["mid"])
        chart.legend.font.name = "Calibri"
    else:
        chart.has_legend = False

    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(label_sz)
    dl.font.color.rgb = RGBColor(*colors["ink"])
    dl.font.name = "Calibri"
    if is_pie:
        dl.number_format = '0%'
        dl.number_format_is_linked = False
        dl.show_percentage = True
        dl.show_value = False
    else:
        dl.show_value = True

    for i, series in enumerate(chart.series):
        color = palette[i % len(palette)]
        if is_pie:
            for j, point in enumerate(series.points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = RGBColor(*palette[j % len(palette)])
        elif ctype == "line":
            series.format.line.color.rgb = RGBColor(*color)
            series.format.line.width = Pt(2.5)
        else:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(*color)

    if not is_pie:
        cat_ax = chart.category_axis
        val_ax = chart.value_axis
        for ax in (cat_ax, val_ax):
            ax.tick_labels.font.size = Pt(tick_sz)
            ax.tick_labels.font.color.rgb = RGBColor(*colors["mute"])
            ax.tick_labels.font.name = "Calibri"
            ax.format.line.color.rgb = RGBColor(*colors["mute"])
        val_ax.has_major_gridlines = False

    return graphic_frame




def main_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_pres", lang))],
        [KeyboardButton(text=tr("btn_word", lang))],
        [KeyboardButton(text=tr("btn_excel", lang))],
        [KeyboardButton(text=tr("btn_history", lang)), KeyboardButton(text=tr("btn_plan", lang))],
        [KeyboardButton(text=tr("btn_language", lang))]
    ], resize_keyboard=True)


def cancel_kb(lang="ru"):
    """Клавиатура-заглушка с одной кнопкой возврата в меню - используется в шагах,
    где пользователь должен прислать свой текст (тему, данные, правки), но всё равно
    должен иметь возможность выйти в главное меню в один тап, а не только через /cancel."""
    return ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text=tr("btn_main_menu", lang))]], resize_keyboard=True)


def mode_kb(show_template=False, lang="ru"):
    rows = [
        [KeyboardButton(text=tr("btn_ai_generate", lang))],
        [KeyboardButton(text=tr("btn_own_text", lang))],
    ]
    if show_template:
        rows.append([KeyboardButton(text=tr("btn_template", lang))])
    rows.append([KeyboardButton(text=tr("btn_main_menu", lang))])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def slides_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_5slides", lang)), KeyboardButton(text=tr("btn_8slides", lang))],
        [KeyboardButton(text=tr("btn_10slides", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
    ], resize_keyboard=True)


def confirm_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_full_version", lang))],
        [KeyboardButton(text=tr("btn_add_info", lang))],
        [KeyboardButton(text=tr("btn_change_style", lang))],
        [KeyboardButton(text=tr("btn_change_topic", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
    ], resize_keyboard=True)


def photo_source_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_ai_photos", lang))],
        [KeyboardButton(text=tr("btn_own_photos", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
    ], resize_keyboard=True)


def photos_done_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_photos_done", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
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

# Минимальный суммарный объём (в словах) для проверки после генерации - см.
# комментарий у WORD_MIN_TOTAL_WORDS.get(...) в word_build. Соответствует нижней
# границе из length_hint, но отдельной константой, чтобы не парсить текст промпта.
WORD_MIN_TOTAL_WORDS = {
    ("coursework", "long"): 6000,
    ("coursework", "short"): 2000,
    ("referat", "long"): 3000,
    ("referat", "short"): 1200,
    ("report", "long"): 1500,
    ("report", "short"): 600,
    ("essay", "long"): 1500,
    ("essay", "short"): 600,
}

# Минимум слов на один содержательный раздел - используется в довыворота-промпте,
# чтобы указывать модели точную недостачу по каждому разделу, а не только по сумме.
WORD_SECTION_MIN_WORDS = {
    ("coursework", "long"): 600,
    ("coursework", "short"): 200,
    ("referat", "long"): 400,
    ("referat", "short"): 150,
    ("report", "long"): 250,
    ("report", "short"): 100,
    ("essay", "long"): 250,
    ("essay", "short"): 100,
}
# Заголовки, которые не нужно искусственно раздувать при довыворота (оглавление,
# список источников - у них естественно фиксированный, а не текстовый объём).
_WORD_SECTION_SKIP_EXPAND = ("содержание", "список литератур", "список источник", "титульный лист", "задание")

WORD_CATEGORY_TITLES = {
    "physical": {"ru": "👤 Для физлиц", "en": "👤 For individuals", "de": "👤 Für Privatpersonen",
                 "ar": "👤 للأفراد", "zh": "👤 个人用户", "es": "👤 Para particulares", "fr": "👤 Pour particuliers"},
    "entity": {"ru": "🏢 Для юрлиц / ИП", "en": "🏢 For businesses / sole traders", "de": "🏢 Für Unternehmen / Selbstständige",
               "ar": "🏢 للشركات / أصحاب الأعمال", "zh": "🏢 企业/个体户", "es": "🏢 Para empresas / autónomos", "fr": "🏢 Pour entreprises / indépendants"},
    "study": {"ru": "🎓 Для учёбы", "en": "🎓 For studies", "de": "🎓 Für Studium/Schule",
              "ar": "🎓 للدراسة", "zh": "🎓 学习用", "es": "🎓 Para estudios", "fr": "🎓 Pour les études"},
}

WORD_CATEGORIES = {
    "physical": {
        "title": WORD_CATEGORY_TITLES["physical"]["ru"],
        "main": ["dkp", "rent", "proxy", "statement"],
        "more": ["act", "loan", "claim", "consent", "marriage_contract", "gift", "lawsuit", "alimony"],
    },
    "entity": {
        "title": WORD_CATEGORY_TITLES["entity"]["ru"],
        "main": ["offer", "services", "employment", "work_act"],
        "more": ["supply", "agency", "joint_activity", "nonresidential_rent", "cession", "nda",
                 "self_employed", "warranty_letter"],
    },
    "study": {
        "title": WORD_CATEGORY_TITLES["study"]["ru"],
        "main": ["doc", "referat", "report", "essay", "notes", "coursework"],
        "more": [],
    },
}

# Название вида документа на каждом из 7 языков - используется и в кнопках выбора,
# и как fallback-заголовок сгенерированного файла (если модель не даст лучше).
WORD_KIND_LABELS = {
    "doc": {"ru": "📄 Обычный документ", "en": "📄 Generic document", "de": "📄 Allgemeines Dokument",
            "ar": "📄 مستند عام", "zh": "📄 普通文档", "es": "📄 Documento genérico", "fr": "📄 Document générique"},
    "referat": {"ru": "🎓 Реферат", "en": "🎓 Research paper", "de": "🎓 Referat",
                "ar": "🎓 بحث", "zh": "🎓 学术论文", "es": "🎓 Trabajo de investigación", "fr": "🎓 Dossier de recherche"},
    "report": {"ru": "🎤 Доклад", "en": "🎤 Report", "de": "🎤 Vortrag",
               "ar": "🎤 تقرير", "zh": "🎤 报告", "es": "🎤 Exposición", "fr": "🎤 Exposé"},
    "essay": {"ru": "✍️ Эссе", "en": "✍️ Essay", "de": "✍️ Essay",
              "ar": "✍️ مقال", "zh": "✍️ 论文/短文", "es": "✍️ Ensayo", "fr": "✍️ Essai"},
    "notes": {"ru": "📓 Конспект", "en": "📓 Study notes", "de": "📓 Konspekt",
              "ar": "📓 ملخص دراسي", "zh": "📓 课堂笔记", "es": "📓 Apuntes", "fr": "📓 Notes de cours"},
    "coursework": {"ru": "📚 Курсовая работа", "en": "📚 Term paper", "de": "📚 Hausarbeit",
                   "ar": "📚 بحث فصلي", "zh": "📚 学期论文", "es": "📚 Trabajo de curso", "fr": "📚 Dossier de cours"},
    "dkp": {"ru": "📝 Договор купли-продажи", "en": "📝 Sale and purchase agreement", "de": "📝 Kaufvertrag",
            "ar": "📝 عقد بيع", "zh": "📝 买卖合同", "es": "📝 Contrato de compraventa", "fr": "📝 Contrat de vente"},
    "rent": {"ru": "🏠 Договор аренды", "en": "🏠 Rental agreement", "de": "🏠 Mietvertrag",
             "ar": "🏠 عقد إيجار", "zh": "🏠 租赁合同", "es": "🏠 Contrato de alquiler", "fr": "🏠 Contrat de location"},
    "proxy": {"ru": "🧾 Доверенность", "en": "🧾 Power of attorney", "de": "🧾 Vollmacht",
              "ar": "🧾 توكيل", "zh": "🧾 授权委托书", "es": "🧾 Poder notarial", "fr": "🧾 Procuration"},
    "statement": {"ru": "📋 Заявление", "en": "📋 Application", "de": "📋 Antrag",
                  "ar": "📋 طلب", "zh": "📋 申请书", "es": "📋 Solicitud", "fr": "📋 Demande"},
    "act": {"ru": "📦 Акт приёма-передачи", "en": "📦 Handover certificate", "de": "📦 Übergabeprotokoll",
            "ar": "📦 محضر تسليم واستلام", "zh": "📦 交接单", "es": "📦 Acta de entrega", "fr": "📦 Procès-verbal de remise"},
    "loan": {"ru": "💰 Расписка / договор займа", "en": "💰 Promissory note / loan agreement", "de": "💰 Schuldschein / Darlehensvertrag",
             "ar": "💰 سند دين / عقد قرض", "zh": "💰 借条/借款合同", "es": "💰 Pagaré / contrato de préstamo", "fr": "💰 Reconnaissance de dette / contrat de prêt"},
    "claim": {"ru": "⚠️ Претензия", "en": "⚠️ Complaint letter", "de": "⚠️ Mängelrüge",
              "ar": "⚠️ مطالبة", "zh": "⚠️ 索赔函", "es": "⚠️ Reclamación", "fr": "⚠️ Réclamation"},
    "consent": {"ru": "✈️ Согласие на выезд ребёнка", "en": "✈️ Child travel consent", "de": "✈️ Zustimmung zur Ausreise des Kindes",
                "ar": "✈️ موافقة على سفر الطفل", "zh": "✈️ 儿童出境同意书", "es": "✈️ Autorización de viaje para menores", "fr": "✈️ Autorisation de voyage pour mineur"},
    "marriage_contract": {"ru": "💍 Брачный договор", "en": "💍 Prenuptial agreement", "de": "💍 Ehevertrag",
                          "ar": "💍 عقد زواج (اتفاقية)", "zh": "💍 婚前协议", "es": "💍 Capitulaciones matrimoniales", "fr": "💍 Contrat de mariage"},
    "gift": {"ru": "🎁 Договор дарения", "en": "🎁 Gift deed", "de": "🎁 Schenkungsvertrag",
             "ar": "🎁 عقد هبة", "zh": "🎁 赠与合同", "es": "🎁 Contrato de donación", "fr": "🎁 Acte de donation"},
    "lawsuit": {"ru": "⚖️ Исковое заявление", "en": "⚖️ Statement of claim", "de": "⚖️ Klageschrift",
                "ar": "⚖️ لائحة دعوى", "zh": "⚖️ 起诉状", "es": "⚖️ Demanda judicial", "fr": "⚖️ Assignation en justice"},
    "alimony": {"ru": "👶 Соглашение об алиментах", "en": "👶 Child support agreement", "de": "👶 Unterhaltsvereinbarung",
                "ar": "👶 اتفاقية نفقة", "zh": "👶 抚养费协议", "es": "👶 Acuerdo de manutención", "fr": "👶 Convention de pension alimentaire"},
    "offer": {"ru": "💼 Коммерческое предложение", "en": "💼 Commercial proposal", "de": "💼 Geschäftsangebot",
              "ar": "💼 عرض تجاري", "zh": "💼 商业提案", "es": "💼 Propuesta comercial", "fr": "💼 Proposition commerciale"},
    "services": {"ru": "🤝 Договор оказания услуг", "en": "🤝 Service agreement", "de": "🤝 Dienstleistungsvertrag",
                 "ar": "🤝 عقد تقديم خدمات", "zh": "🤝 服务合同", "es": "🤝 Contrato de prestación de servicios", "fr": "🤝 Contrat de prestation de services"},
    "employment": {"ru": "👔 Трудовой договор", "en": "👔 Employment contract", "de": "👔 Arbeitsvertrag",
                   "ar": "👔 عقد عمل", "zh": "👔 劳动合同", "es": "👔 Contrato de trabajo", "fr": "👔 Contrat de travail"},
    "work_act": {"ru": "✅ Акт выполненных работ", "en": "✅ Certificate of completed work", "de": "✅ Abnahmeprotokoll",
                 "ar": "✅ محضر إنجاز الأعمال", "zh": "✅ 工作完成验收单", "es": "✅ Acta de obra terminada", "fr": "✅ Procès-verbal de réception des travaux"},
    "supply": {"ru": "🚚 Договор поставки", "en": "🚚 Supply agreement", "de": "🚚 Liefervertrag",
               "ar": "🚚 عقد توريد", "zh": "🚚 供货合同", "es": "🚚 Contrato de suministro", "fr": "🚚 Contrat de fourniture"},
    "agency": {"ru": "🕴️ Агентский договор", "en": "🕴️ Agency agreement", "de": "🕴️ Agenturvertrag",
               "ar": "🕴️ عقد وكالة", "zh": "🕴️ 代理合同", "es": "🕴️ Contrato de agencia", "fr": "🕴️ Contrat d'agence"},
    "joint_activity": {"ru": "🤝 Договор о совместной деятельности", "en": "🤝 Joint activity agreement", "de": "🤝 Vertrag über gemeinsame Tätigkeit",
                       "ar": "🤝 عقد نشاط مشترك", "zh": "🤝 合作协议", "es": "🤝 Contrato de actividad conjunta", "fr": "🤝 Contrat d'activité commune"},
    "nonresidential_rent": {"ru": "🏢 Аренда нежилого помещения", "en": "🏢 Commercial premises lease", "de": "🏢 Gewerbemietvertrag",
                            "ar": "🏢 عقد إيجار عقار غير سكني", "zh": "🏢 非居住用房租赁合同", "es": "🏢 Alquiler de local comercial", "fr": "🏢 Bail commercial"},
    "cession": {"ru": "🔄 Договор цессии", "en": "🔄 Assignment of claim agreement", "de": "🔄 Abtretungsvertrag",
                "ar": "🔄 عقد حوالة حق", "zh": "🔄 债权转让合同", "es": "🔄 Contrato de cesión de derechos", "fr": "🔄 Contrat de cession de créance"},
    "nda": {"ru": "🔒 Соглашение о неразглашении (NDA)", "en": "🔒 Non-disclosure agreement (NDA)", "de": "🔒 Vertraulichkeitsvereinbarung (NDA)",
            "ar": "🔒 اتفاقية عدم إفصاح", "zh": "🔒 保密协议", "es": "🔒 Acuerdo de confidencialidad (NDA)", "fr": "🔒 Accord de confidentialité (NDA)"},
    "self_employed": {"ru": "🧑‍💻 Договор с самозанятым", "en": "🧑‍💻 Contract with a self-employed contractor", "de": "🧑‍💻 Vertrag mit Selbstständigem",
                      "ar": "🧑‍💻 عقد مع عامل مستقل", "zh": "🧑‍💻 与自由职业者合同", "es": "🧑‍💻 Contrato con autónomo", "fr": "🧑‍💻 Contrat avec un travailleur indépendant"},
    "warranty_letter": {"ru": "✉️ Гарантийное письмо", "en": "✉️ Letter of guarantee", "de": "✉️ Garantieschreiben",
                        "ar": "✉️ خطاب ضمان", "zh": "✉️ 保证函", "es": "✉️ Carta de garantía", "fr": "✉️ Lettre de garantie"},
}
KIND_LABELS = {k: v["ru"] for k, v in WORD_KIND_LABELS.items()}  # обратная совместимость (заголовки по умолчанию и т.п.)
# Обратный словарь label->kind собирается по ВСЕМ языкам сразу, чтобы кнопка, нажатая
# на любом из 7 языков, узнавалась одинаково, независимо от текущего языка пользователя.
LABEL_TO_KIND = {label: k for k, variants in WORD_KIND_LABELS.items() for label in variants.values()}


def word_category_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=WORD_CATEGORY_TITLES["physical"][lang if lang in WORD_CATEGORY_TITLES["physical"] else "ru"])],
        [KeyboardButton(text=WORD_CATEGORY_TITLES["entity"][lang if lang in WORD_CATEGORY_TITLES["entity"] else "ru"])],
        [KeyboardButton(text=WORD_CATEGORY_TITLES["study"][lang if lang in WORD_CATEGORY_TITLES["study"] else "ru"])],
        [KeyboardButton(text=tr("btn_main_menu", lang))],
    ], resize_keyboard=True)


WORD_TITLE_TO_CATEGORY = {label: cat for cat, variants in WORD_CATEGORY_TITLES.items() for label in variants.values()}


def word_kind_kb(category, more=False, lang="ru"):
    cat = WORD_CATEGORIES.get(category, WORD_CATEGORIES["physical"])
    ids = list(cat["main"]) + (list(cat["more"]) if more else [])
    rows = [[KeyboardButton(text=WORD_KIND_LABELS[k].get(lang, WORD_KIND_LABELS[k]["ru"]))] for k in ids]
    if cat["more"] and not more:
        rows.append([KeyboardButton(text=tr("btn_more_docs", lang))])
    rows.append([KeyboardButton(text=tr("btn_back_categories", lang))])
    rows.append([KeyboardButton(text=tr("btn_main_menu", lang))])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def word_size_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_shallow", lang))],
        [KeyboardButton(text=tr("btn_deep", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
    ], resize_keyboard=True)


def word_confirm_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_build_doc", lang))],
        [KeyboardButton(text=tr("btn_add_info", lang))],
        [KeyboardButton(text=tr("btn_change_query", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
    ], resize_keyboard=True)


EXCEL_CATEGORIES = {
    "physical": {
        "title": WORD_CATEGORY_TITLES["physical"]["ru"],
        "kinds": ["expense_estimate", "family_budget", "startup_model"],
    },
    "entity": {
        "title": WORD_CATEGORY_TITLES["entity"]["ru"],
        "kinds": ["project_budget", "price_list"],
    },
    "study": {
        "title": WORD_CATEGORY_TITLES["study"]["ru"],
        "kinds": ["calc_table"],
    },
}

EXCEL_KIND_LABELS_I18N = {
    "expense_estimate": {"ru": "🧾 Смета расходов", "en": "🧾 Expense estimate", "de": "🧾 Kostenvoranschlag",
                         "ar": "🧾 تقدير المصروفات", "zh": "🧾 支出预算表", "es": "🧾 Presupuesto de gastos", "fr": "🧾 Devis de dépenses"},
    "family_budget": {"ru": "💰 Семейный бюджет", "en": "💰 Family budget", "de": "💰 Familienbudget",
                      "ar": "💰 ميزانية الأسرة", "zh": "💰 家庭预算", "es": "💰 Presupuesto familiar", "fr": "💰 Budget familial"},
    "startup_model": {"ru": "🚀 Финмодель стартапа", "en": "🚀 Startup financial model", "de": "🚀 Startup-Finanzmodell",
                      "ar": "🚀 نموذج مالي لشركة ناشئة", "zh": "🚀 创业财务模型", "es": "🚀 Modelo financiero de startup", "fr": "🚀 Modèle financier de startup"},
    "project_budget": {"ru": "📊 Смета проекта / бизнес-план", "en": "📊 Project budget / business plan", "de": "📊 Projektbudget / Businessplan",
                       "ar": "📊 ميزانية مشروع / خطة عمل", "zh": "📊 项目预算/商业计划", "es": "📊 Presupuesto de proyecto / plan de negocio", "fr": "📊 Budget de projet / plan d'affaires"},
    "price_list": {"ru": "🏷️ Прайс-лист", "en": "🏷️ Price list", "de": "🏷️ Preisliste",
                   "ar": "🏷️ قائمة الأسعار", "zh": "🏷️ 价目表", "es": "🏷️ Lista de precios", "fr": "🏷️ Liste de prix"},
    "calc_table": {"ru": "📐 Расчётная таблица к работе", "en": "📐 Calculation table for coursework", "de": "📐 Berechnungstabelle für die Arbeit",
                   "ar": "📐 جدول حسابي للبحث", "zh": "📐 作业计算表", "es": "📐 Tabla de cálculo para el trabajo", "fr": "📐 Tableau de calcul pour le devoir"},
}
EXCEL_KIND_LABELS = {k: v["ru"] for k, v in EXCEL_KIND_LABELS_I18N.items()}  # обратная совместимость
EXCEL_LABEL_TO_KIND = {label: k for k, variants in EXCEL_KIND_LABELS_I18N.items() for label in variants.values()}

# У финмодели стартапа данные всегда реальные, от пользователя - модель их не придумывает,
# только оформляет реальными формулами Excel (см. решение в этой сессии). Остальные виды
# работают как таблицы в презентациях/курсовых: можно как сгенерировать иллюстративные цифры
# по теме, так и вставить свои реальные данные - и в том, и в другом случае формулы настоящие.
EXCEL_KIND_HINTS_I18N = {
    "expense_estimate": {
        "ai": {"ru": "На что смета? Например: смета на ремонт кухни, смета на свадьбу на 40 человек.",
               "en": "What is the estimate for? E.g. kitchen renovation, a wedding for 40 people.",
               "de": "Wofür ist der Kostenvoranschlag? Z. B. Küchenrenovierung, Hochzeit für 40 Personen.",
               "ar": "لماذا التقدير؟ مثلاً: تجديد المطبخ، حفل زفاف لـ٤٠ شخصاً.",
               "zh": "预算用于什么？例如：厨房装修、40人婚礼。",
               "es": "¿Para qué es el presupuesto? Por ejemplo: reforma de cocina, boda para 40 personas.",
               "fr": "Devis pour quoi ? Ex. : rénovation de cuisine, mariage pour 40 personnes."},
        "user": {"ru": "Пришли список статей расходов с количеством и ценой (например: линолеум - 25 м² - 800 ₽) - я оформлю в таблицу с формулами.",
                 "en": "Send a list of expense items with quantity and price (e.g. linoleum - 25 m² - $80) - I'll format it into a table with formulas.",
                 "de": "Sende eine Liste der Ausgabenposten mit Menge und Preis (z. B. Linoleum - 25 m² - 80 €) - ich erstelle eine Tabelle mit Formeln.",
                 "ar": "أرسل قائمة ببنود المصروفات مع الكمية والسعر (مثال: أرضية فينيل - ٢٥ م² - ٨٠ دولار) - سأنسقها في جدول بمعادلات.",
                 "zh": "请发送费用项目清单（含数量和单价），例如：地板革-25平方米-80美元，我会做成带公式的表格。",
                 "es": "Envía la lista de partidas de gasto con cantidad y precio (ej.: linóleo - 25 m² - 80 €) - lo pondré en una tabla con fórmulas.",
                 "fr": "Envoie la liste des postes de dépense avec quantité et prix (ex. : linoléum - 25 m² - 80 €) - je mettrai ça en tableau avec formules."},
    },
    "family_budget": {
        "ai": {"ru": "Опиши примерный бюджет, который нужно смоделировать (например: бюджет семьи из 3 человек на месяц).",
               "en": "Describe the budget to model (e.g. monthly budget for a family of 3).",
               "de": "Beschreibe das Budget, das modelliert werden soll (z. B. Monatsbudget für eine 3-köpfige Familie).",
               "ar": "صف الميزانية المطلوب نمذجتها (مثال: ميزانية شهرية لأسرة من ٣ أفراد).",
               "zh": "描述需要建模的预算（例如：三口之家的月度预算）。",
               "es": "Describe el presupuesto a modelar (ej.: presupuesto mensual para una familia de 3).",
               "fr": "Décris le budget à modéliser (ex. : budget mensuel pour une famille de 3 personnes)."},
        "user": {"ru": "Пришли свои статьи доходов и расходов с суммами - я оформлю в таблицу с формулами (итого доходы/расходы/остаток).",
                 "en": "Send your income and expense items with amounts - I'll format a table with formulas (total income/expenses/balance).",
                 "de": "Sende deine Einnahmen- und Ausgabenposten mit Beträgen - ich erstelle eine Tabelle mit Formeln (Summe Einnahmen/Ausgaben/Rest).",
                 "ar": "أرسل بنود دخلك ومصروفاتك مع المبالغ - سأنسقها في جدول بمعادلات (إجمالي الدخل/المصروفات/المتبقي).",
                 "zh": "请发送你的收支项目及金额，我会做成带公式的表格（总收入/支出/结余）。",
                 "es": "Envía tus partidas de ingresos y gastos con importes - lo pondré en una tabla con fórmulas (total ingresos/gastos/saldo).",
                 "fr": "Envoie tes postes de revenus et dépenses avec les montants - je ferai un tableau avec formules (total revenus/dépenses/solde)."},
    },
    "project_budget": {
        "ai": {"ru": "На какой проект нужна смета/бизнес-план? Опиши в двух словах.",
               "en": "Which project needs the budget/business plan? Describe briefly.",
               "de": "Für welches Projekt wird der Kostenvoranschlag/Businessplan benötigt? Kurz beschreiben.",
               "ar": "لأي مشروع تحتاج الميزانية/خطة العمل؟ صف بإيجاز.",
               "zh": "哪个项目需要预算/商业计划？请简要描述。",
               "es": "¿Para qué proyecto es el presupuesto/plan de negocio? Descríbelo brevemente.",
               "fr": "Pour quel projet faut-il le budget/plan d'affaires ? Décris brièvement."},
        "user": {"ru": "Пришли статьи доходов и расходов проекта с суммами - оформлю в таблицу с формулами (итого доходы/расходы/прибыль).",
                 "en": "Send the project's income and expense items with amounts - I'll format a table with formulas (total income/expenses/profit).",
                 "de": "Sende die Einnahmen- und Ausgabenposten des Projekts mit Beträgen - ich erstelle eine Tabelle mit Formeln (Summe Einnahmen/Ausgaben/Gewinn).",
                 "ar": "أرسل بنود دخل ومصروفات المشروع مع المبالغ - سأنسقها في جدول بمعادلات (إجمالي الدخل/المصروفات/الربح).",
                 "zh": "请发送项目的收支项目及金额，我会做成带公式的表格（总收入/支出/利润）。",
                 "es": "Envía las partidas de ingresos y gastos del proyecto con importes - lo pondré en una tabla con fórmulas (total ingresos/gastos/beneficio).",
                 "fr": "Envoie les postes de revenus et dépenses du projet avec les montants - je ferai un tableau avec formules (total revenus/dépenses/bénéfice)."},
    },
    "price_list": {
        "ai": {"ru": "Прайс на что? Например: прайс-лист кофейни, прайс на клининговые услуги.",
               "en": "Price list for what? E.g. a coffee shop menu, cleaning services.",
               "de": "Preisliste wofür? Z. B. Café-Preisliste, Reinigungsdienstleistungen.",
               "ar": "قائمة أسعار لماذا؟ مثال: قائمة أسعار مقهى، خدمات تنظيف.",
               "zh": "价目表用于什么？例如：咖啡店价目表、清洁服务。",
               "es": "¿Lista de precios de qué? Ej.: carta de una cafetería, servicios de limpieza.",
               "fr": "Liste de prix pour quoi ? Ex. : carte d'un café, services de nettoyage."},
        "user": {"ru": "Пришли позиции прайса: наименование, единица, количество, цена, скидка (если есть) - оформлю с формулами сумм.",
                 "en": "Send price list items: name, unit, quantity, price, discount (if any) - I'll format it with sum formulas.",
                 "de": "Sende die Positionen: Bezeichnung, Einheit, Menge, Preis, Rabatt (falls vorhanden) - ich erstelle Summenformeln.",
                 "ar": "أرسل بنود القائمة: الاسم، الوحدة، الكمية، السعر، الخصم (إن وجد) - سأضيف معادلات الإجمالي.",
                 "zh": "请发送价目表条目：名称、单位、数量、单价、折扣（如有），我会加上求和公式。",
                 "es": "Envía las partidas: nombre, unidad, cantidad, precio, descuento (si lo hay) - añadiré fórmulas de suma.",
                 "fr": "Envoie les postes : nom, unité, quantité, prix, remise (le cas échéant) - j'ajouterai les formules de somme."},
    },
    "calc_table": {
        "ai": {"ru": "По какой теме нужна расчётная таблица для работы? Например: результаты опроса по теме курсовой.",
               "en": "What topic is the calculation table for? E.g. survey results for a term paper.",
               "de": "Zu welchem Thema wird die Berechnungstabelle benötigt? Z. B. Umfrageergebnisse für die Hausarbeit.",
               "ar": "لأي موضوع يحتاج الجدول الحسابي؟ مثال: نتائج استبيان لبحث دراسي.",
               "zh": "计算表用于什么主题？例如：学期论文的问卷调查结果。",
               "es": "¿Para qué tema es la tabla de cálculo? Ej.: resultados de una encuesta para un trabajo de curso.",
               "fr": "Pour quel sujet faut-il le tableau de calcul ? Ex. : résultats d'une enquête pour un dossier de cours."},
        "user": {"ru": "Пришли показатели и значения - оформлю в таблицу с формулой доли в % и итогом.",
                 "en": "Send indicators and values - I'll format a table with a % share formula and total.",
                 "de": "Sende Kennzahlen und Werte - ich erstelle eine Tabelle mit %-Anteilsformel und Summe.",
                 "ar": "أرسل المؤشرات والقيم - سأنسقها في جدول بمعادلة النسبة المئوية والإجمالي.",
                 "zh": "请发送指标及数值，我会做成带百分比公式和合计的表格。",
                 "es": "Envía los indicadores y valores - lo pondré en una tabla con fórmula de porcentaje y total.",
                 "fr": "Envoie les indicateurs et valeurs - je ferai un tableau avec formule de pourcentage et total."},
    },
}
EXCEL_KIND_HINTS = {k: {"ai": v["ai"]["ru"], "user": v["user"]["ru"]} for k, v in EXCEL_KIND_HINTS_I18N.items()}  # обратная совместимость


def excel_kind_hint(kind, mode_key, lang):
    variants = EXCEL_KIND_HINTS_I18N.get(kind, EXCEL_KIND_HINTS_I18N["expense_estimate"])[mode_key]
    return variants.get(lang, variants["ru"])


def excel_category_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=WORD_CATEGORY_TITLES["physical"].get(lang, WORD_CATEGORY_TITLES["physical"]["ru"]))],
        [KeyboardButton(text=WORD_CATEGORY_TITLES["entity"].get(lang, WORD_CATEGORY_TITLES["entity"]["ru"]))],
        [KeyboardButton(text=WORD_CATEGORY_TITLES["study"].get(lang, WORD_CATEGORY_TITLES["study"]["ru"]))],
        [KeyboardButton(text=tr("btn_main_menu", lang))],
    ], resize_keyboard=True)


def excel_kind_kb(category, lang="ru"):
    cat = EXCEL_CATEGORIES.get(category, EXCEL_CATEGORIES["physical"])
    rows = [[KeyboardButton(text=EXCEL_KIND_LABELS_I18N[k].get(lang, EXCEL_KIND_LABELS_I18N[k]["ru"]))] for k in cat["kinds"]]
    rows.append([KeyboardButton(text=tr("btn_back_categories", lang))])
    rows.append([KeyboardButton(text=tr("btn_main_menu", lang))])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def excel_confirm_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_build_table", lang))],
        [KeyboardButton(text=tr("btn_add_info", lang))],
        [KeyboardButton(text=tr("btn_change_query", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))]
    ], resize_keyboard=True)


def style_kb(include_keep=False, lang="ru"):
    rows = []
    if include_keep:
        rows.append([KeyboardButton(text=tr("btn_keep_style", lang))])
    theme_order = ["nature", "business", "tech", "school", "history", "science", "sport", "travel",
                   "food", "art", "eco", "minimal", "fashion", "default"]
    for i in range(0, len(theme_order), 2):
        pair = theme_order[i:i + 2]
        rows.append([KeyboardButton(text=THEME_LABELS_I18N[k].get(lang, THEME_LABELS_I18N[k]["ru"])) for k in pair])
    rows.append([KeyboardButton(text=tr("btn_main_menu", lang))])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


STYLE_BY_LABEL = {label: k for k, variants in THEME_LABELS_I18N.items() for label in variants.values()}


@dp.message(Command("start"))
async def cmd_start(m: Message, state: FSMContext):
    u = get_user(m.from_user.id)
    u["name"] = m.from_user.first_name or "друг"
    await state.clear()
    if u.get("lang_chosen"):
        lang = user_lang(m.from_user.id)
        await m.answer(tr("msg_welcome", lang, name=u["name"]), reply_markup=main_kb(lang))
        return
    await state.set_state(Form.waiting_language)
    intro = " / ".join(TR["msg_choose_lang"][c] for c in ("ru", "en", "ar", "zh"))
    await m.answer(intro, reply_markup=lang_kb())


@dp.message(Form.waiting_language)
async def set_language(m: Message, state: FSMContext):
    code = LANG_LABEL_TO_CODE.get((m.text or "").strip())
    if not code:
        await m.answer(" / ".join(TR["msg_choose_lang"][c] for c in ("ru", "en", "ar", "zh")), reply_markup=lang_kb())
        return
    u = get_user(m.from_user.id)
    u["lang"] = code
    u["lang_chosen"] = True
    await state.clear()
    await m.answer(tr("msg_welcome", code, name=u.get("name") or "🙂"), reply_markup=main_kb(code))


@dp.message(F.text.in_(TR["btn_language"].values()))
async def change_language(m: Message, state: FSMContext):
    await state.set_state(Form.waiting_language)
    await m.answer(tr("msg_choose_lang", user_lang(m.from_user.id)), reply_markup=lang_kb())


@dp.message(Command("cancel"))
async def cmd_cancel(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.clear()
    finish_job(m.from_user.id)
    await m.answer(tr("msg_cancelled", lang), reply_markup=main_kb(lang))


@dp.message(Command("menu"))
@dp.message(F.text.in_(ALL_MAIN_MENU_LABELS))
async def to_main_menu(m: Message, state: FSMContext):
    """Кнопка/команда возврата в главное меню - работает на любом шаге любого процесса
    (презентация/документ/таблица), а не только в конце. Регистрируется здесь, до всех
    хендлеров конкретных состояний (Form.waiting_*), которые иначе перехватили бы этот
    текст как обычный пользовательский ввод (тему, данные и т.п.) - aiogram проверяет
    хендлеры в порядке регистрации и берёт первый подошедший."""
    lang = user_lang(m.from_user.id)
    await state.clear()
    finish_job(m.from_user.id)
    await m.answer(tr("msg_main_menu", lang), reply_markup=main_kb(lang))


@dp.message(F.text.in_(ALL_BTN_PRES_LABELS))
async def start_pres(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    if not can_generate(m.from_user.id):
        await m.answer(tr("msg_limit", lang))
        return
    await m.answer(tr("msg_how_build_pres", lang), reply_markup=mode_kb(lang=lang))
    await state.set_state(Form.waiting_mode)


@dp.message(Form.waiting_mode, F.text.in_(ALL_BTN_AI_GENERATE_LABELS))
async def mode_ai(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(mode="ai", user_text="")
    await m.answer(tr("msg_topic_prompt", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_topic)


@dp.message(Form.waiting_mode, F.text.in_(ALL_BTN_OWN_TEXT_LABELS))
async def mode_user(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(mode="user")
    await m.answer(tr("msg_own_text_prompt", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_user_text)


@dp.message(Form.waiting_topic)
async def process_topic(m: Message, state: FSMContext):
    text = m.text or ""
    lang = user_lang(m.from_user.id)
    if len(text) > 500:
        await m.answer(tr("msg_too_long_500", lang))
        return
    name, _ = pick_theme(text)
    await state.update_data(topic=text, extra="", extra_used=0, theme_name=name)
    style_label = THEME_LABELS_I18N.get(name, {}).get(lang, THEME_LABELS.get(name, name))
    await m.answer(
        tr("msg_style_fits_topic", lang, style=style_label),
        reply_markup=style_kb(include_keep=True, lang=lang)
    )
    await state.set_state(Form.waiting_theme)


@dp.message(Form.waiting_user_text)
async def process_user_text(m: Message, state: FSMContext):
    text = m.text or ""
    lang = user_lang(m.from_user.id)
    if len(text) < 40:
        await m.answer(tr("msg_too_little_text", lang))
        return
    if len(text) > 4000:
        await m.answer(tr("msg_too_long_4000", lang))
        return
    name, _ = pick_theme(text)
    topic = text[:80].replace("\n", " ")
    await state.update_data(user_text=text, topic=topic, extra="", extra_used=0, theme_name=name)
    style_label = THEME_LABELS_I18N.get(name, {}).get(lang, THEME_LABELS.get(name, name))
    await m.answer(
        tr("msg_style_fits_text", lang, style=style_label),
        reply_markup=style_kb(include_keep=True, lang=lang)
    )
    await state.set_state(Form.waiting_theme)


@dp.message(Form.waiting_theme)
async def process_theme(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    label = m.text or ""
    if label not in TR["btn_keep_style"].values():
        name = STYLE_BY_LABEL.get(label)
        if not name:
            await m.answer(tr("msg_pick_style", lang), reply_markup=style_kb(include_keep=True, lang=lang))
            return
        await state.update_data(theme_name=name)
    await m.answer(tr("msg_how_many_slides", lang), reply_markup=slides_kb(lang))
    await state.set_state(Form.waiting_slides)


@dp.message(Form.waiting_slides)
async def process_slides(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    slides = 8
    t = m.text or ""
    if "5" in t:
        slides = 5
    elif "10" in t:
        slides = 10
    data = await state.get_data()
    angle = random.choice(ANGLES)
    await state.update_data(slides=slides, angle=angle)
    await m.answer(tr("msg_building_sample", lang), reply_markup=cancel_kb(lang))
    lang_instr = grok_lang_instruction(lang)

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
Без JSON, без вводных фраз, без нумерации "Слайд 1".{lang_instr}"""
    else:
        prompt = f"""Ты — редактор презентаций высокого уровня, который умеет находить неочевидный
и интересный угол в любой теме, избегая шаблонных заголовков вроде "Введение" или "Что это такое".
Тема: {data.get('topic')}
Слайдов: {slides}
Доп: {data.get('extra')}
Угол подачи: {angle}
Дай короткий уникальный план текстом: цепляющее конкретное название и 3 пункта — каждый должен быть
содержательным тезисом (что именно будет рассказано), а не общей категорией.
Без JSON, без вводных фраз, без нумерации "Слайд 1".{lang_instr}"""

    sample = await ask_grok(prompt)
    if grok_failed(sample):
        await m.answer(
            tr("msg_grok_error", lang),
            reply_markup=slides_kb(lang)
        )
        return
    theme_name = data.get("theme_name", "default")
    style_label = THEME_LABELS_I18N.get(theme_name, {}).get(lang, THEME_LABELS.get(theme_name, theme_name))
    await state.update_data(sample=sample)
    await send_draft(
        m,
        tr("msg_draft_ready_pres", lang, sample=sample, style=style_label),
        title=data.get("topic", tr("msg_start_pres_again", lang)),
        reply_markup=confirm_kb(lang)
    )
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(ALL_BTN_CHANGE_TOPIC_LABELS))
async def change_topic(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    if data.get("mode") == "user":
        await m.answer(tr("msg_new_text", lang), reply_markup=cancel_kb(lang))
        await state.set_state(Form.waiting_user_text)
    else:
        await m.answer(tr("msg_new_topic", lang), reply_markup=cancel_kb(lang))
        await state.set_state(Form.waiting_topic)


@dp.message(Form.waiting_confirm, F.text.in_(ALL_BTN_CHANGE_STYLE_LABELS))
async def change_style(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(tr("msg_which_style", lang), reply_markup=style_kb(lang=lang))
    await state.set_state(Form.waiting_style)


@dp.message(Form.waiting_style)
async def process_style(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    label = m.text or ""
    name = STYLE_BY_LABEL.get(label)
    if not name:
        await m.answer(tr("msg_pick_style", lang), reply_markup=style_kb(lang=lang))
        return
    data = await state.get_data()
    await state.update_data(theme_name=name)
    await m.answer(
        tr("msg_style_changed", lang, style=label, sample=data.get('sample')),
        reply_markup=confirm_kb(lang)
    )
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(ALL_BTN_ADD_INFO_LABELS))
async def add_extra(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    if (await state.get_data()).get("extra_used", 0) >= 3:
        await m.answer(tr("msg_extra_limit", lang), reply_markup=confirm_kb(lang))
        return
    await m.answer(tr("msg_extra_prompt", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_extra)


@dp.message(Form.waiting_extra)
async def process_extra(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    text = m.text or ""
    if len(text) > 800:
        await m.answer(tr("msg_too_long_800", lang))
        return
    extra = ((data.get("extra") or "") + "\n" + text).strip()
    angle = random.choice(ANGLES)
    await state.update_data(extra=extra, extra_used=data.get("extra_used", 0) + 1, angle=angle)
    await m.answer(tr("msg_updating_draft", lang))
    lang_instr = grok_lang_instruction(lang)
    if data.get("mode") == "user":
        prompt = f"""Исправь и обнови структуру.
Исходный текст:
{data.get('user_text')}
Дополнительно:
{extra}
Слайдов: {data.get('slides')}
Угол: {angle}
Короткий план: название и 3 пункта. Без JSON.{lang_instr}"""
    else:
        prompt = f"""Тема: {data.get('topic')}
Доп: {extra}
Угол: {angle}
Новый короткий план, 3 пункта. Без JSON.{lang_instr}"""
    sample = await ask_grok(prompt)
    if grok_failed(sample):
        await m.answer(
            tr("msg_draft_update_failed", lang),
            reply_markup=confirm_kb(lang)
        )
        await state.set_state(Form.waiting_confirm)
        return
    await state.update_data(sample=sample)
    await send_draft(m, tr("msg_draft_updated", lang, sample=sample), title=data.get("topic", tr("msg_start_pres_again", lang)), reply_markup=confirm_kb(lang))
    await state.set_state(Form.waiting_confirm)


@dp.message(Form.waiting_confirm, F.text.in_(ALL_BTN_FULL_VERSION_LABELS))
async def ask_photo_source(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(user_photos=[])
    await m.answer(
        tr("msg_photo_or_own", lang),
        reply_markup=photo_source_kb(lang)
    )
    await state.set_state(Form.waiting_pres_photo_choice)


@dp.message(Form.waiting_pres_photo_choice, F.text.in_(ALL_BTN_AI_PHOTOS_LABELS))
async def photo_source_ai(m: Message, state: FSMContext):
    await _build_presentation(m, state)


@dp.message(Form.waiting_pres_photo_choice, F.text.in_(ALL_BTN_OWN_PHOTOS_LABELS))
async def photo_source_own(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(
        tr("msg_send_photos_one_by_one", lang),
        reply_markup=photos_done_kb(lang)
    )
    await state.set_state(Form.waiting_pres_photos)


@dp.message(Form.waiting_pres_photos, F.photo)
async def collect_pres_photo(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    photos = data.get("user_photos") or []
    if len(photos) >= 20:
        await m.answer(tr("msg_enough_photos", lang))
        return
    uid = m.from_user.id
    path = f"/tmp/{uid}_userphoto_{len(photos)}_{random.randint(1000, 9999)}.jpg"
    try:
        file = await bot.get_file(m.photo[-1].file_id)
        await bot.download_file(file.file_path, path)
    except Exception as e:
        print("Photo download error:", e)
        await m.answer(tr("msg_photo_download_fail", lang))
        return
    photos.append(path)
    await state.update_data(user_photos=photos)
    await m.answer(tr("msg_photo_received", lang, n=len(photos)))


@dp.message(Form.waiting_pres_photos, F.text.in_(ALL_BTN_PHOTOS_DONE_LABELS))
async def finish_pres_photos(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    if not (data.get("user_photos") or []):
        await m.answer(tr("msg_photo_wait", lang), reply_markup=photos_done_kb(lang))
        return
    await _build_presentation(m, state)


@dp.message(Form.waiting_pres_photos)
async def pres_photos_fallback(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(tr("msg_send_photos", lang), reply_markup=photos_done_kb(lang))


async def _build_presentation(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    ok, reason = start_job(uid)
    if not ok:
        await m.answer(reason)
        return
    await m.answer(tr("msg_building_pres_photos", lang))
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
    - Фото: описание живого кадра (реальная сцена, человек, объект, место), не стоковый шаблон и не абстракция.
    - График (ключ "chart", НЕОБЯЗАТЕЛЬНЫЙ): добавляй его ТОЛЬКО если тема или конкретный слайд подразумевает
      цифры, статистику, доли, сравнение или динамику по времени - и не больше 1-2 слайдов с графиком на всю
      презентацию. Если тема не про цифры (например: питомцы, отношения, психология, искусство, литература,
      рецепты, путешествия как впечатление) - НЕ добавляй chart вообще, у слайда остаётся только фото.
      Формат: "chart": {{"chart_type": "bar" | "line" | "pie", "title": "короткий заголовок графика",
      "categories": ["...", "..."], "series": [{{"name": "...", "values": [12, 34, ...]}}]}} - придумай
      правдоподобные цифры по теме. У слайда с графиком image_prompt всё равно укажи (на случай, если график
      не соберётся), но использоваться будет либо график, либо фото, не оба сразу."""
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
    Только JSON (ключ "chart" добавляй лишь на 1-2 слайдах и только если тема того требует, иначе не пиши его вовсе):
    {{"title":"...","slides":[{{"title":"...","content":"абзац1\n\nабзац2","image_prompt":"unique cinematic scene","chart":null}}]}}{grok_json_lang_instruction(lang)}""")
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
    - Фото: описание живого кадра (реальная сцена, человек, объект, место), не стоковый шаблон и не абстракция.
    - График (ключ "chart", НЕОБЯЗАТЕЛЬНЫЙ): добавляй его ТОЛЬКО если тема или конкретный слайд подразумевает
      цифры, статистику, доли, сравнение или динамику по времени - и не больше 1-2 слайдов с графиком на всю
      презентацию. Если тема не про цифры (например: питомцы, отношения, психология, искусство, литература,
      рецепты, путешествия как впечатление) - НЕ добавляй chart вообще, у слайда остаётся только фото.
      Формат: "chart": {{"chart_type": "bar" | "line" | "pie", "title": "короткий заголовок графика",
      "categories": ["...", "..."], "series": [{{"name": "...", "values": [12, 34, ...]}}]}} - придумай
      правдоподобные цифры по теме. У слайда с графиком image_prompt всё равно укажи (на случай, если график
      не соберётся), но использоваться будет либо график, либо фото, не оба сразу."""
            raw = await ask_grok(f"""Собери уникальную презентацию уровня лучшего журнала на эту тему,
    найди неочевидный и интересный угол, избегай банальностей.
    Тема: {data.get('topic')}
    Слайдов: {data.get('slides')}
    Доп: {data.get('extra')}
    Угол: {angle}
    Стиль: {theme_name}
    {common_rules}
    Только JSON (ключ "chart" добавляй лишь на 1-2 слайдах и только если тема того требует, иначе не пиши его вовсе):
    {{"title":"...","slides":[{{"title":"...","content":"абзац1\n\nабзац2","image_prompt":"unique cinematic scene","chart":null}}]}}{grok_json_lang_instruction(lang)}""")

        try:
            content = extract_json(raw)
            if not isinstance(content.get("slides"), list) or not content["slides"]:
                raise ValueError("В ответе модели нет слайдов")
        except Exception as e:
            print("Presentation JSON parse error:", e)
            await m.answer(tr("msg_no_text", lang, btn=tr("msg_start_pres_again", lang)), reply_markup=main_kb(lang))
            await state.clear()
            return

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        slides_data = content.get("slides", [])
        n = len(slides_data)

        # Свои фото пользователя (если он их прислал) идут в очередь: первое - на обложку,
        # остальные - по слайдам по порядку; на то, что не хватило, генерируем через ИИ как раньше.
        user_photos = list(data.get("user_photos") or [])

        cover_img = None
        cover_src = f"/tmp/{uid}_cover.png"
        cover_own = user_photos.pop(0) if user_photos else None
        cover_ok = bool(cover_own) or await generate_image(
            f"{content.get('title')}, wide cinematic opening shot, {colors['photo']}", cover_src
        )
        if cover_ok:
            cover_wide = f"/tmp/{uid}_cover_w.png"
            cover(cover_own or cover_src, cover_wide, 1920, 1080)
            cover_img = cover_wide

        # Если модель дала валидный chart для слайда - фото для него не генерируем вообще
        # (график займёт то же место в раскладке, экономим время и токены на генерацию картинки).
        MAX_CHARTS = 2
        charts = []
        chart_count = 0
        for s in slides_data:
            c = s.get("chart")
            valid = (
                isinstance(c, dict) and c.get("categories") and c.get("series")
                and chart_count < MAX_CHARTS
            )
            if valid:
                chart_count += 1
                charts.append(c)
            else:
                charts.append(None)

        images = []
        raw_sources = []
        for i, s in enumerate(slides_data):
            if charts[i]:
                images.append(None)
                continue
            own = user_photos.pop(0) if user_photos else None
            if own:
                src, ok = own, True
            else:
                src = f"/tmp/{uid}_{i}_{random.randint(1000, 9999)}.png"
                prompt = f"{s.get('image_prompt') or s.get('title')}, {colors['photo']}, unique composition"
                ok = await generate_image(prompt, src)
            if ok:
                raw_sources.append(src)
                wide, tall = f"/tmp/{uid}_{i}_{random.randint(1000, 9999)}_w.png", f"/tmp/{uid}_{i}_{random.randint(1000, 9999)}_t.png"
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
            chart_data = charts[idx] if idx < len(charts) else None
            if layout == 0:
                if chart_data:
                    add_chart(slide, 0.5, 0.6, 5.4, 6.3, chart_data, colors)
                elif img:
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
                if chart_data:
                    add_chart(slide, 0.7, 0.35, 11.9, 4.0, chart_data, colors)
                elif img:
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
                if chart_data:
                    add_chart(slide, 8.5, 1.3, 4.6, 4.7, chart_data, colors)
                elif img:
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
        for p in (cover_src, cover_own, cover_img, pptx_path, pdf_path, *raw_sources, *[f for pair in images if pair for f in pair]):
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


def _tabbed_line(doc, parts, tabs_cm=(9, 13), size=12, before=0, after=0):
    """Строка с элементами по табуляции - для строк вида 'подпись ___ И.О. Фамилия'.
    parts: список (текст, bold)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = DocxPt(before)
    p.paragraph_format.space_after = DocxPt(after)
    for cm in tabs_cm:
        p.paragraph_format.tab_stops.add_tab_stop(Cm(cm))
    for i, (text, bold) in enumerate(parts):
        if i > 0:
            p.add_run("\t")
        _run(p, text, size, bold)
    return p


def _student_title_page(doc, kind_label, title, meta):
    """Титульный лист по образцу техникума/ссуза: жирные лейблы ВЫПОЛНИЛ/ПРОВЕРИЛ/ОЦЕНКА
    слева, линия подписи с ФИО справа от неё, мелкие подписи (подпись)/(ФИО) под линией."""
    _p(doc, meta.get("org") or "Министерство образования", 13, align=WD_ALIGN_PARAGRAPH.CENTER, after=0)
    _p(doc, meta.get("school") or "[Название учебного заведения]", 13, align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(5):
        doc.add_paragraph()
    _p(doc, title, 16, True, WD_ALIGN_PARAGRAPH.CENTER, after=10)
    _p(doc, kind_label, 14, True, WD_ALIGN_PARAGRAPH.CENTER, after=10)
    if meta.get("discipline"):
        _p(doc, f"по предмету/дисциплине: {meta.get('discipline')}", 12, True, WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(5):
        doc.add_paragraph()
    _p(doc, "ВЫПОЛНИЛ", 12, True, after=0)
    _tabbed_line(
        doc,
        [(f"студент {meta.get('group') or '[группа]'} группы", False), ("__________", False), (meta.get("author") or "И.О. Фамилия", True)],
        after=0,
    )
    _tabbed_line(doc, [("", False), ("(подпись)", False), ("(ФИО)", False)], size=9, after=8)
    _p(doc, "«___»____________ 20___ г.", 12, before=0, after=16)
    _p(doc, "ПРОВЕРИЛ", 12, True, after=0)
    _tabbed_line(
        doc,
        [(meta.get("teacher_position") or "", False), ("__________", False), (meta.get("teacher") or "И.О. Фамилия", True)],
        after=0,
    )
    _tabbed_line(doc, [("", False), ("(подпись)", False), ("(ФИО)", False)], size=9, after=8)
    doc.add_paragraph()
    _p(doc, "ОЦЕНКА", 12, True, after=0)
    _tabbed_line(doc, [("", False), ("__________________________", False)], after=16)
    _p(doc, "«___»____________ 20___ г.", 12, after=30)
    _p(doc, f"{meta.get('city') or '[Город]'} {meta.get('year') or '2026'}", 12, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()


def _set_cell_borders(cell):
    """Тонкая чёрная рамка у ячейки таблицы - python-docx не даёт это напрямую, только через XML."""
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tcPr.append(borders)


def _add_table(doc, table_data, size=11):
    """table_data: {"headers": [...], "rows": [[...], ...]}. Рисует таблицу с рамками
    и жирной шапкой - для перечней товаров/имущества в договорах и актах, статистики
    и практических данных в курсовых/рефератах."""
    headers = table_data.get("headers") or []
    rows = table_data.get("rows") or []
    if not headers and not rows:
        return
    ncols = len(headers) if headers else max((len(r) for r in rows), default=0)
    if ncols == 0:
        return
    t = doc.add_table(rows=0, cols=ncols)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if headers:
        row = t.add_row()
        for i, htext in enumerate(headers):
            cell = row.cells[i]
            cell.text = ""
            _run(cell.paragraphs[0], str(htext), size, True)
            _set_cell_borders(cell)
    for r in rows:
        row = t.add_row()
        for i in range(ncols):
            cell = row.cells[i]
            cell.text = ""
            val = str(r[i]) if i < len(r) else ""
            _run(cell.paragraphs[0], val, size, False)
            _set_cell_borders(cell)
    doc.add_paragraph()  # отступ после таблицы


def _parse_markdown_table(lines):
    """Модель иногда пишет таблицу markdown-синтаксисом прямо в тексте раздела
    (| ячейка | ячейка |) вместо структурированного ключа "table" - на такой случай
    разбираем это защитным парсером и рендерим как настоящую таблицу Word.
    Возвращает {"headers": [...], "rows": [[...], ...]} или None, если это не таблица."""
    rows = []
    for ln in lines:
        ln = ln.strip()
        if not (ln.startswith("|") and ln.endswith("|") and ln.count("|") >= 3):
            return None
        rows.append([c.strip() for c in ln.strip("|").split("|")])
    if len(rows) < 2:
        return None
    sep = rows[1]
    if sep and all(set(c) <= set("-: ") for c in sep):
        rows.pop(1)
    if not rows:
        return None
    return {"headers": rows[0], "rows": rows[1:]}


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
        raw_paras = [x.strip() for x in body.split("\n") if x.strip()]
        i = 0
        while i < len(raw_paras):
            if raw_paras[i].startswith("|") and raw_paras[i].endswith("|"):
                j = i
                while j < len(raw_paras) and raw_paras[j].startswith("|") and raw_paras[j].endswith("|"):
                    j += 1
                md_table = _parse_markdown_table(raw_paras[i:j])
                if md_table:
                    _add_table(doc, md_table, size=size - 1 if size > 9 else size)
                    i = j
                    continue
            _p(doc, raw_paras[i], size, first=first if indent else None, after=6, line=line)
            i += 1
        table_data = block.get("table")
        if isinstance(table_data, dict):
            _add_table(doc, table_data, size=size - 1 if size > 9 else size)


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
    "referat": '{"org":"Министерство образования","school":"[учебное заведение]","discipline":"[предмет/дисциплина/МДК]","group":"[номер группы]","author":"[ФИО студента]","teacher":"[должность и/или ФИО преподавателя]","city":"[город]","year":"2026"}',
    "coursework": '{"org":"Министерство образования","school":"[учебное заведение]","specialty":"[специальность]","author":"[ФИО студента]","teacher":"[ФИО научного руководителя]","city":"[город]","year":"2026"}',
    "report": '{"school":"[учебное заведение]","discipline":"[предмет/дисциплина]","group":"[номер группы]","author":"[ФИО]","teacher":"[должность и/или ФИО преподавателя]"}',
    "essay": '{"school":"[учебное заведение]","discipline":"[предмет/дисциплина]","group":"[номер группы]","author":"[ФИО]","teacher":"[должность и/или ФИО преподавателя]"}',
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
        _student_title_page(doc, "РЕФЕРАТ", title, meta)
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
        sec.left_margin = Cm(2)
        sec.right_margin = Cm(2)
        _student_title_page(doc, "ДОКЛАД", title, meta)
        _body(doc, sections, indent=True, head_center=False)

    elif kind == "essay":
        sec.left_margin = Cm(2)
        sec.right_margin = Cm(2)
        _student_title_page(doc, "ЭССЕ", title, meta)
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


# ==================== EXCEL ====================

def _xl_hex(color_tuple):
    return "%02X%02X%02X" % color_tuple


def _xl_num(v, default=0.0):
    # Модель/пользователь должны прислать чистое число, но это не гарантировано -
    # иногда проскакивает "500 000", "500 000 ₽" или "12 месяцев" вместо 500000/12.
    # Вместо падения с ValueError на float()/int() достаём первое число из строки,
    # иначе одна нечисловая ячейка ломает формулу умножения/суммы на всём листе
    # (Excel даёт #VALUE! на любую арифметику с текстом).
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\xa0", " ").replace(" ", "").replace(",", ".")
    m = re.search(r"-?\d+(\.\d+)?", s)
    return float(m.group()) if m else default


# Приблизительный курс к рублю на момент написания (конец августа 2026) - используется
# только для финмодели стартапа, если пользователь называет суммы не в рублях, а всё
# остальное в таблице считается в ₽. Курс плавает, поэтому в итоговый файл всегда
# добавляется примечание с датой и курсом, чтобы пользователь мог свериться и при
# необходимости попросить пересчитать по актуальному курсу.
FX_TO_RUB = {"RUB": 1.0, "РУБ": 1.0, "₽": 1.0, "USD": 85.0, "$": 85.0, "EUR": 95.0, "€": 95.0}
FX_NOTE_DATE = "конец августа 2026"


def _xl_fx_rate(currency):
    return FX_TO_RUB.get((currency or "RUB").strip().upper(), 1.0)


def _xl_style_sheet(ws, colors, n_cols):
    """Базовое оформление листа под тему презентаций/документов: акцентная шапка,
    тонкие границы у таблицы, читаемая ширина колонок, альбомная печать по ширине
    (чтобы при печати/экспорте в PDF таблица не резалась на колонки по границе листа)."""
    ws.sheet_view.showGridLines = False
    for i in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 22 if i == 1 else 16
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = False


def _xl_header_row(ws, row_idx, headers, colors):
    fill = PatternFill(start_color=_xl_hex(colors["line"]), end_color=_xl_hex(colors["line"]), fill_type="solid")
    font = XlFont(bold=True, color=_xl_hex((255, 255, 255) if sum(colors["line"]) < 400 else (20, 20, 20)), size=11)
    thin = Side(style="thin", color=_xl_hex(colors["mute"]))
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row_idx, column=i, value=h)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def _xl_body_row(ws, row_idx, values, colors, number_cols=None, money_cols=None, percent_cols=None):
    thin = Side(style="thin", color=_xl_hex(colors["mute"]))
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row_idx, column=i, value=v)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")
        if money_cols and i in money_cols:
            c.number_format = '#,##0.00 ₽'
        elif percent_cols and i in percent_cols:
            c.number_format = '0.0%'
        elif number_cols and i in number_cols:
            c.number_format = '#,##0.##'


def _xl_total_row(ws, row_idx, label, values, colors, n_cols, money_cols=None, percent_cols=None):
    fill = PatternFill(start_color=_xl_hex(colors["bg"] if sum(colors["bg"]) < 400 else (235, 235, 235)),
                        end_color=_xl_hex(colors["bg"] if sum(colors["bg"]) < 400 else (235, 235, 235)), fill_type="solid")
    thin = Side(style="thin", color=_xl_hex(colors["mute"]))
    ws.cell(row=row_idx, column=1, value=label).font = XlFont(bold=True)
    for i in range(1, n_cols + 1):
        c = ws.cell(row=row_idx, column=i)
        c.border = Border(left=thin, right=thin, top=Side(style="double", color=_xl_hex(colors["mute"])), bottom=thin)
        c.font = XlFont(bold=True)
        if i > 1 and i - 2 < len(values) and values[i - 2] is not None:
            c.value = values[i - 2]
            if money_cols and i in money_cols:
                c.number_format = '#,##0.00 ₽'
            elif percent_cols and i in percent_cols:
                c.number_format = '0.0%'


def _xl_title(ws, title, colors, n_cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = XlFont(bold=True, size=14, color=_xl_hex(colors["ink"] if sum(colors["bg"]) > 400 else (30, 30, 30)))
    ws.row_dimensions[1].height = 28


def _xl_inject_cached_values(path, cache):
    """openpyxl пишет в ячейку с формулой только сам текст формулы ("=SUM(...)"),
    без результата - настоящий Excel/Numbers/Google Таблицы пересчитывают всё сами
    при открытии, но быстрый просмотр файла (например Quick Look на iPhone) формулы
    не считает и показывает пустую ячейку или 0. Чтобы число было видно в любом
    просмотрщике, а не только в "настоящем" Excel, здесь вручную дописывается
    закешированный результат прямо в XML уже сохранённого файла - формула остаётся
    рабочей и редактируемой, просто рядом с ней лежит готовое число на первый показ.
    cache: {"E8": 4500, "B13": 59900.0, ...} - адрес ячейки -> посчитанное код."""
    import zipfile
    import shutil

    sheet_path = "xl/worksheets/sheet1.xml"
    tmp_path = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin:
        data = {n: zin.read(n) for n in zin.namelist()}
    xml = data[sheet_path].decode("utf-8")
    for ref, val in cache.items():
        if val is None:
            continue
        # openpyxl уже пишет пустой <v></v> (или самозакрывающийся <v/>) сразу после
        # </f> - подставляем число внутрь него, а не добавляем новый тег.
        pattern = re.compile(r'(<c r="%s"[^>]*>(?:(?!</c>).)*?</f>)<v\s*/?>(?:</v>)?(</c>)' % re.escape(ref), re.DOTALL)
        xml, n = pattern.subn(lambda m: f"{m.group(1)}<v>{val}</v>{m.group(2)}", xml, count=1)
        if not n:
            print(f"_xl_inject_cached_values: не нашёл ячейку {ref} с формулой (не критично)")
    data[sheet_path] = xml.encode("utf-8")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)
    shutil.move(tmp_path, path)


# Фиксированные схемы колонок под каждый вид - модель (или сам пользователь) поставляет
# только содержимое строк (rows), формулы и оформление всегда собираются кодом, а не ИИ,
# чтобы суммы/проценты/итоги были настоящими формулами Excel, а не текстом.
EXCEL_ROW_SCHEMA = {
    "expense_estimate": {"fields": ["name", "qty", "price"], "example": '{"name":"...", "qty": 1, "price": 1000}'},
    "family_budget": {"fields": ["name", "type", "amount"], "example": '{"name":"...", "type": "доход"|"расход", "amount": 1000}'},
    "project_budget": {"fields": ["name", "type", "amount"], "example": '{"name":"...", "type": "доход"|"расход", "amount": 1000}'},
    "price_list": {"fields": ["name", "unit", "qty", "price", "discount"], "example": '{"name":"...", "unit":"шт", "qty": 1, "price": 1000, "discount": 0}'},
    "calc_table": {"fields": ["name", "value"], "example": '{"name":"...", "value": 100}'},
}


def build_excel_items(path, title, kind, colors, rows):
    """Собирает смету/бюджет/прайс/расчётную таблицу с реальными формулами Excel.
    rows - список dict по схеме EXCEL_ROW_SCHEMA[kind]."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Таблица"
    cache = {}  # адрес ячейки -> посчитанное код значение, для _xl_inject_cached_values

    if kind == "expense_estimate":
        headers = ["№", "Статья расходов", "Кол-во", "Цена, ₽", "Сумма, ₽"]
        _xl_style_sheet(ws, colors, len(headers))
        _xl_title(ws, title, colors, len(headers))
        _xl_header_row(ws, 3, headers, colors)
        r = 4
        total = 0.0
        for i, row in enumerate(rows, 1):
            qty, price = _xl_num(row.get("qty"), 1), _xl_num(row.get("price"))
            line_sum = qty * price
            total += line_sum
            _xl_body_row(ws, r, [i, row.get("name", ""), qty, price, f"=C{r}*D{r}"],
                         colors, number_cols={3}, money_cols={4, 5})
            cache[f"E{r}"] = line_sum
            r += 1
        _xl_total_row(ws, r, "ИТОГО", [None, None, None, f"=SUM(E4:E{r - 1})"], colors, len(headers), money_cols={5})
        cache[f"E{r}"] = total

    elif kind in ("family_budget", "project_budget"):
        headers = ["№", "Статья", "Тип", "Сумма, ₽"]
        _xl_style_sheet(ws, colors, len(headers))
        _xl_title(ws, title, colors, len(headers))
        _xl_header_row(ws, 3, headers, colors)
        income_label, expense_label = "Доход", "Расход"
        r = 4
        total_inc, total_exp = 0.0, 0.0
        for i, row in enumerate(rows, 1):
            # Нормализуем "Тип" в самом коде, а не полагаемся на то, что модель/пользователь
            # напишет ровно "доход"/"расход" - иначе SUMIF ниже просто не найдёт совпадение
            # (он сравнивает строки целиком, а не по смыслу) и итог молча окажется нулевым.
            raw_type = (row.get("type") or "").strip().lower()
            norm_type = expense_label if "расход" in raw_type else income_label
            amount = _xl_num(row.get("amount"))
            if norm_type == income_label:
                total_inc += amount
            else:
                total_exp += amount
            _xl_body_row(ws, r, [i, row.get("name", ""), norm_type, amount],
                         colors, money_cols={4})
            r += 1
        last = r - 1
        r_inc, r_exp = r, r + 1
        ws.cell(row=r_inc, column=1, value="").border = Border()
        _xl_total_row(ws, r_inc, "Итого доходы", [None, None, f'=SUMIF(C4:C{last},"{income_label}",D4:D{last})'], colors, len(headers), money_cols={4})
        cache[f"D{r_inc}"] = total_inc
        _xl_total_row(ws, r_exp, "Итого расходы", [None, None, f'=SUMIF(C4:C{last},"{expense_label}",D4:D{last})'], colors, len(headers), money_cols={4})
        cache[f"D{r_exp}"] = total_exp
        r_res = r_exp + 1
        res_label = "Остаток" if kind == "family_budget" else "Прибыль"
        _xl_total_row(ws, r_res, res_label, [None, None, f"=D{r_inc}-D{r_exp}"], colors, len(headers), money_cols={4})
        cache[f"D{r_res}"] = total_inc - total_exp

    elif kind == "price_list":
        headers = ["№", "Наименование", "Ед.", "Кол-во", "Цена, ₽", "Скидка, %", "Сумма, ₽"]
        _xl_style_sheet(ws, colors, len(headers))
        _xl_title(ws, title, colors, len(headers))
        _xl_header_row(ws, 3, headers, colors)
        r = 4
        total = 0.0
        for i, row in enumerate(rows, 1):
            disc = _xl_num(row.get("discount"))
            qty, price = _xl_num(row.get("qty"), 1), _xl_num(row.get("price"))
            disc_frac = disc / 100 if disc else 0
            line_sum = qty * price * (1 - disc_frac)
            total += line_sum
            _xl_body_row(ws, r, [i, row.get("name", ""), row.get("unit", "шт"), qty,
                                  price, disc_frac, f"=D{r}*E{r}*(1-F{r})"],
                         colors, number_cols={4}, money_cols={5, 7}, percent_cols={6})
            cache[f"G{r}"] = line_sum
            r += 1
        _xl_total_row(ws, r, "ИТОГО", [None, None, None, None, None, f"=SUM(G4:G{r - 1})"], colors, len(headers), money_cols={7})
        cache[f"G{r}"] = total

    elif kind == "calc_table":
        headers = ["№", "Показатель", "Значение", "Доля, %"]
        _xl_style_sheet(ws, colors, len(headers))
        _xl_title(ws, title, colors, len(headers))
        _xl_header_row(ws, 3, headers, colors)
        r = 4
        first_r = r
        values = []
        for i, row in enumerate(rows, 1):
            v = _xl_num(row.get("value"))
            values.append(v)
            _xl_body_row(ws, r, [i, row.get("name", ""), v, None], colors, number_cols={3}, percent_cols={4})
            r += 1
        last_r = r - 1
        total_value = sum(values)
        # IFERROR - если все значения окажутся нулевыми (или их сумма равна нулю), формула
        # деления вернёт #DIV/0! в каждой строке; выводим 0% вместо ошибки на весь лист.
        for rr, v in zip(range(first_r, r), values):
            ws.cell(row=rr, column=4, value=f"=IFERROR(C{rr}/SUM($C${first_r}:$C${last_r}),0)")
            ws.cell(row=rr, column=4).number_format = '0.0%'
            cache[f"D{rr}"] = (v / total_value) if total_value else 0.0
        total_pct_formula = f"=IFERROR(SUM(C{first_r}:C{last_r})/SUM(C{first_r}:C{last_r}),0)"
        _xl_total_row(ws, r, "ИТОГО", [None, f"=SUM(C{first_r}:C{last_r})", total_pct_formula], colors, len(headers), percent_cols={4})
        cache[f"C{r}"] = total_value
        cache[f"D{r}"] = 1.0 if total_value else 0.0

    wb.save(path)
    if cache:
        _xl_inject_cached_values(path, cache)


def build_excel_startup(path, title, colors, data):
    """Финмодель стартапа - данные (инвестиции, расходы, цена, продажи, срок) реальные,
    от пользователя, никогда не придумываются моделью. Всё, что видно в файле кроме
    самих исходных чисел - настоящие формулы Excel (выручка/прибыль/накопительный итог)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Финмодель"

    horizon = max(1, min(int(_xl_num(data.get("horizon_months"), 12)), 36))
    investment = _xl_num(data.get("investment"))
    price = _xl_num(data.get("price"))
    monthly_sales = _xl_num(data.get("monthly_sales"))
    expenses = [{"name": (e or {}).get("name", ""), "amount": _xl_num((e or {}).get("amount"))}
                for e in (data.get("expenses") or [])]

    n_cols = 5
    _xl_style_sheet(ws, colors, n_cols)
    _xl_title(ws, title, colors, n_cols)
    cache = {}  # адрес ячейки -> посчитанное код значение, для _xl_inject_cached_values

    ws.cell(row=2, column=1, value=f"Стартовые вложения: {investment:,.0f} ₽".replace(",", " ")).font = XlFont(italic=True, size=10)

    # блок постоянных ежемесячных расходов - отдельной таблицей, чтобы сумма расходов
    # в помесячном расчёте тоже была формулой (SUM), а не готовым числом
    r = 4
    ws.cell(row=r, column=1, value="Ежемесячные расходы").font = XlFont(bold=True)
    r += 1
    exp_first = r
    total_exp_value = sum((e.get("amount") or 0) for e in expenses)
    for e in expenses:
        _xl_body_row(ws, r, [None, e.get("name", ""), None, None, e.get("amount", 0)], colors, money_cols={5})
        r += 1
    exp_last = r - 1 if expenses else exp_first
    if not expenses:
        _xl_body_row(ws, r, [None, "—", None, None, 0], colors, money_cols={5})
        exp_last = r
        r += 1
    total_exp_row = r
    _xl_total_row(ws, total_exp_row, "Итого расходов в месяц", [None, None, None, f"=SUM(E{exp_first}:E{exp_last})"], colors, n_cols, money_cols={5})
    cache[f"E{total_exp_row}"] = total_exp_value
    r = total_exp_row + 1

    fx_note = data.get("_fx_note")
    if fx_note:
        ws.cell(row=r, column=1, value=f"💱 {fx_note}").font = XlFont(italic=True, size=9, color=_xl_hex(colors["mute"]))
        r += 1

    r += 1

    ws.cell(row=r, column=1, value=f"Цена за единицу: {price:,.0f} ₽ · Продажи в месяц: {monthly_sales:,.0f} шт.".replace(",", " ")).font = XlFont(italic=True, size=10)
    r += 2

    headers = ["Месяц", "Выручка, ₽", "Расходы, ₽", "Прибыль, ₽", "Накопительно, ₽"]
    _xl_header_row(ws, r, headers, colors)
    table_start = r + 1
    r = table_start

    # cum_values считаем заранее (не только для "срока окупаемости", но и чтобы
    # сразу знать, какое число закешировать в каждую ячейку E - формула та же математика)
    rev_value = price * monthly_sales
    profit_value = rev_value - total_exp_value
    cum_values = []
    running = -investment
    for month in range(1, horizon + 1):
        running += profit_value
        cum_values.append(running)

    for month in range(1, horizon + 1):
        rev_formula = f"={price}*{monthly_sales}"
        exp_formula = f"=E{total_exp_row}"
        profit_formula = f"=B{r}-C{r}"
        if month == 1:
            cum_formula = f"=D{r}-{investment}"
        else:
            cum_formula = f"=D{r}+E{r - 1}"
        _xl_body_row(ws, r, [f"Месяц {month}", rev_formula, exp_formula, profit_formula, cum_formula], colors, money_cols={2, 3, 4, 5})
        cache[f"B{r}"] = rev_value
        cache[f"C{r}"] = total_exp_value
        cache[f"D{r}"] = profit_value
        cache[f"E{r}"] = cum_values[month - 1]
        r += 1
    table_end = r - 1

    # срок окупаемости - первый месяц, где накопительный итог формулой стал неотрицательным;
    # считается по-настоящему из реальных чисел пользователя (это математика, а не выдумка),
    # но записывается как обычная сводная ячейка, а не как сложная формула поиска,
    # чтобы не зависеть от array-формул, которые по-разному ведут себя в разных версиях Excel.
    payback = next((i + 1 for i, v in enumerate(cum_values) if v >= 0), None)

    r += 1
    ws.cell(row=r, column=1, value="Срок окупаемости").font = XlFont(bold=True)
    ws.cell(row=r, column=2, value=(f"{payback} мес." if payback else f"не достигается за {horizon} мес."))
    r += 1
    ws.cell(row=r, column=1, value="Точка безубыточности (продаж/мес)").font = XlFont(bold=True)
    be_cell = ws.cell(row=r, column=2, value=f"=ROUNDUP(E{total_exp_row}/{price},0)" if price else "—")
    if price:
        be_cell.number_format = '#,##0 "шт."'
        import math
        cache[f"B{r}"] = math.ceil(total_exp_value / price)

    for i in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
    ws.column_dimensions["A"].width = 34

    wb.save(path)
    _xl_inject_cached_values(path, cache)


EXCEL_KIND_DESC = {
    "expense_estimate": "смета расходов",
    "family_budget": "семейный бюджет",
    "project_budget": "смета проекта / бизнес-план",
    "price_list": "прайс-лист",
    "calc_table": "расчётная таблица к учебной работе",
    "startup_model": "финансовая модель стартапа",
}


def excel_mode_kb(lang="ru"):
    return ReplyKeyboardMarkup(keyboard=[
        [KeyboardButton(text=tr("btn_ai_generate", lang))],
        [KeyboardButton(text=tr("btn_own_data", lang))],
        [KeyboardButton(text=tr("btn_main_menu", lang))],
    ], resize_keyboard=True)


@dp.message(F.text.in_(ALL_BTN_EXCEL_LABELS))
async def start_excel(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    if not can_generate(m.from_user.id):
        await m.answer(tr("msg_limit", lang))
        return
    await m.answer(tr("msg_for_whom_table", lang), reply_markup=excel_category_kb(lang))
    await state.set_state(Form.waiting_excel_category)


@dp.message(Form.waiting_excel_category)
async def excel_category(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    cat = WORD_TITLE_TO_CATEGORY.get((m.text or "").strip())
    if not cat:
        await m.answer(tr("msg_pick_category", lang), reply_markup=excel_category_kb(lang))
        return
    await state.update_data(excel_category=cat)
    await m.answer(tr("msg_which_table", lang), reply_markup=excel_kind_kb(cat, lang))
    await state.set_state(Form.waiting_excel_kind)


@dp.message(Form.waiting_excel_kind, F.text.in_(ALL_BACK_CATEGORIES_LABELS))
async def excel_kind_back(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(tr("msg_for_whom_table", lang), reply_markup=excel_category_kb(lang))
    await state.set_state(Form.waiting_excel_category)


@dp.message(Form.waiting_excel_kind)
async def excel_kind(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    kind = EXCEL_LABEL_TO_KIND.get((m.text or "").strip())
    if not kind:
        data = await state.get_data()
        cat = data.get("excel_category", "physical")
        await m.answer(tr("msg_pick_table", lang), reply_markup=excel_kind_kb(cat, lang))
        return
    await state.update_data(excel_kind=kind, extra="")
    if kind == "startup_model":
        # Финмодель - данные всегда реальные от пользователя, режим "сгенерировать" тут не предлагаем.
        await m.answer(
            tr("msg_startup_data_prompt", lang),
            reply_markup=cancel_kb(lang)
        )
        await state.set_state(Form.waiting_excel_startup_data)
        return
    await m.answer(
        tr("msg_how_build_table", lang),
        reply_markup=excel_mode_kb(lang)
    )
    await state.set_state(Form.waiting_excel_mode)


@dp.message(Form.waiting_excel_mode, F.text.in_(ALL_BTN_AI_GENERATE_LABELS))
async def excel_mode_ai(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    kind = data.get("excel_kind", "expense_estimate")
    await state.update_data(excel_mode="ai")
    await m.answer(excel_kind_hint(kind, "ai", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_excel_topic)


@dp.message(Form.waiting_excel_mode, F.text.in_(ALL_BTN_OWN_DATA_LABELS))
async def excel_mode_user(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    kind = data.get("excel_kind", "expense_estimate")
    await state.update_data(excel_mode="user")
    await m.answer(excel_kind_hint(kind, "user", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_excel_data)


@dp.message(Form.waiting_excel_topic)
async def excel_topic(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(excel_topic=m.text or "")
    await m.answer(tr("msg_build_table_q", lang), reply_markup=excel_confirm_kb(lang))
    await state.set_state(Form.waiting_excel_confirm)


@dp.message(Form.waiting_excel_data)
async def excel_data(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(excel_topic=m.text or "")
    await m.answer(tr("msg_build_table_q", lang), reply_markup=excel_confirm_kb(lang))
    await state.set_state(Form.waiting_excel_confirm)


@dp.message(Form.waiting_excel_startup_data)
async def excel_startup_data(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(excel_topic=m.text or "")
    await m.answer(tr("msg_build_model_q", lang), reply_markup=excel_confirm_kb(lang))
    await state.set_state(Form.waiting_excel_confirm)


@dp.message(Form.waiting_excel_confirm, F.text.in_(ALL_BTN_CHANGE_QUERY_LABELS))
async def excel_change_request(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    kind = data.get("excel_kind", "expense_estimate")
    if kind == "startup_model":
        await m.answer(tr("msg_resend_data", lang), reply_markup=cancel_kb(lang))
        await state.set_state(Form.waiting_excel_startup_data)
    else:
        mode = data.get("excel_mode", "ai")
        hint_key = "ai" if mode == "ai" else "user"
        await m.answer(excel_kind_hint(kind, hint_key, lang), reply_markup=cancel_kb(lang))
        await state.set_state(Form.waiting_excel_topic if mode == "ai" else Form.waiting_excel_data)


@dp.message(Form.waiting_excel_confirm, F.text.in_(ALL_BTN_ADD_INFO_LABELS))
async def excel_add_extra(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(tr("msg_what_add", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_excel_extra)


@dp.message(Form.waiting_excel_extra)
async def excel_extra(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    extra = (data.get("extra", "") + "\n" + (m.text or "")).strip()
    await state.update_data(extra=extra)
    await m.answer(tr("msg_build_table_q", lang), reply_markup=excel_confirm_kb(lang))
    await state.set_state(Form.waiting_excel_confirm)


@dp.message(Form.waiting_excel_confirm, F.text.in_(ALL_BTN_BUILD_TABLE_LABELS))
async def excel_build(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    ok, reason = start_job(uid)
    if not ok:
        await m.answer(reason)
        return
    await m.answer(tr("msg_building_excel", lang))
    kind = data.get("excel_kind", "expense_estimate")
    topic = data.get("excel_topic", "")
    extra = data.get("extra", "")
    theme_name, colors = pick_theme(topic)
    lang_instr = grok_json_lang_instruction(lang)

    try:
        if kind == "startup_model":
            # Только извлечение реальных чисел, которые пользователь уже прислал -
            # модель не имеет права ничего досочинять или менять числа.
            raw = await ask_grok(f"""Извлеки структурированные данные для финансовой модели стартапа
из сообщения пользователя. НИЧЕГО не выдумывай и не досчитывай за пользователя - если какого-то
числа нет в тексте, ставь 0 (кроме horizon_months - если не указано, ставь 12).
Для КАЖДОЙ суммы (стартовые вложения, каждая статья расходов, цена за единицу) отдельно определи,
в какой валюте её назвал пользователь - по явному указанию ("долларов", "$", "евро", "€") или по
контексту. Если валюта явно не указана - ставь "RUB". НИКОГДА не конвертируй суммы сам и не пересчитывай
их в рубли - оставляй исходное число ровно как в сообщении пользователя, конвертация делается отдельно
кодом по курсу, а не тобой.
Сообщение пользователя:
{topic}
Доп. уточнения: {extra}
Только JSON:
{{"title":"короткое название проекта по смыслу сообщения","investment":0,"investment_currency":"RUB",
"expenses":[{{"name":"...","amount":0,"currency":"RUB"}}],"price":0,"price_currency":"RUB",
"monthly_sales":0,"horizon_months":12}}{lang_instr}""")
            parsed = extract_json(raw)

            # Конвертация в рубли делается здесь кодом по фиксированному курсу (FX_TO_RUB),
            # а не доверяется модели - иначе есть риск, что она либо забудет сконвертировать
            # (как было раньше: доллары просто вставлялись как есть вместо рублей), либо
            # применит собственный устаревший/придуманный курс. Собираем список реальных
            # конвертаций, чтобы честно показать пользователю, что и по какому курсу пересчитано.
            fx_notes = []

            def _conv(amount, currency, label):
                rate = _xl_fx_rate(currency)
                amount = _xl_num(amount)
                if rate != 1.0 and amount:
                    converted = amount * rate
                    fx_notes.append(f"{label}: {amount:,.0f} {currency} → {converted:,.0f} ₽".replace(",", " "))
                    return converted
                return amount

            parsed["investment"] = _conv(parsed.get("investment"), parsed.get("investment_currency"), "Стартовые вложения")
            parsed["price"] = _conv(parsed.get("price"), parsed.get("price_currency"), "Цена за единицу")
            for e in (parsed.get("expenses") or []):
                e["amount"] = _conv(e.get("amount"), e.get("currency"), e.get("name") or "Расход")
            if fx_notes:
                parsed["_fx_note"] = (
                    f"Курс на {FX_NOTE_DATE} (примерный, уточни актуальный при необходимости): "
                    + "; ".join(fx_notes)
                )

            xlsx_path = f"/tmp/xl_{uid}.xlsx"
            build_excel_startup(xlsx_path, parsed.get("title") or "Финансовая модель", colors, parsed)
            title_for_name = parsed.get("title") or "Финмодель"
        else:
            schema = EXCEL_ROW_SCHEMA[kind]
            if data.get("excel_mode") == "user":
                task_line = (
                    "Ниже данные прислал сам пользователь - структурируй их в JSON СТРОГО как есть, "
                    "ничего не досочиняя и не меняя числа. Если каких-то полей не хватает - оставь "
                    "разумные значения по умолчанию (0 или пустая строка), но не выдумывай новые позиции."
                )
            else:
                task_line = (
                    "Пользователь не прислал реальные данные - подбери сам правдоподобные иллюстративные "
                    "данные по теме (как в примерах для учебных работ), 5-8 строк, разумные по масштабу цифры."
                )
            raw = await ask_grok(f"""Собери содержимое для Excel-таблицы: {EXCEL_KIND_DESC.get(kind)}.
{task_line}
Тема/данные: {topic}
Доп: {extra}
Каждая строка - объект с полями {schema['fields']}, например {schema['example']}.
Только JSON:
{{"title":"короткое название таблицы по теме","rows":[{schema['example']}]}}{lang_instr}""")
            parsed = extract_json(raw)
            rows = parsed.get("rows") or []
            if not rows:
                raise ValueError("Модель не вернула строки таблицы")
            xlsx_path = f"/tmp/xl_{uid}.xlsx"
            build_excel_items(xlsx_path, parsed.get("title") or EXCEL_KIND_DESC.get(kind, "Таблица"), kind, colors, rows)
            title_for_name = parsed.get("title") or EXCEL_KIND_DESC.get(kind)

        fname = safe_filename(title_for_name, fallback=EXCEL_KIND_DESC.get(kind, "Таблица"))
        await m.answer_document(FSInputFile(xlsx_path, filename=f"{fname}.xlsx"), caption=tr("msg_excel_caption", lang))
        u["generations"] += 1
        u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {title_for_name}")
        try:
            os.remove(xlsx_path)
        except Exception:
            pass
        note = tr("msg_table_ready", lang)
        if kind == "startup_model":
            note += tr("msg_check_source_numbers", lang)
            if parsed.get("_fx_note"):
                note += f"\n\n💱 {parsed['_fx_note']}"
        await m.answer(note, reply_markup=main_kb(lang))
        await state.clear()
    except Exception as e:
        print("Excel build error:", e)
        await m.answer(tr("msg_build_error_table", lang), reply_markup=main_kb(lang))
        await state.clear()
    finally:
        finish_job(uid)


@dp.message(Form.waiting_excel_confirm)
async def excel_confirm_fallback(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(tr("msg_choose_action", lang), reply_markup=excel_confirm_kb(lang))


@dp.message(F.text.in_(ALL_BTN_WORD_LABELS))
async def start_word(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    if not can_generate(m.from_user.id):
        await m.answer(tr("msg_limit", lang))
        return
    await m.answer(tr("msg_for_whom_doc", lang), reply_markup=word_category_kb(lang))
    await state.set_state(Form.waiting_word_category)


@dp.message(Form.waiting_word_category)
async def word_category(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    cat = WORD_TITLE_TO_CATEGORY.get((m.text or "").strip())
    if not cat:
        await m.answer(tr("msg_pick_category", lang), reply_markup=word_category_kb(lang))
        return
    await state.update_data(word_category=cat)
    await m.answer(tr("msg_which_doc", lang), reply_markup=word_kind_kb(cat, lang=lang))
    await state.set_state(Form.waiting_word_kind)


@dp.message(Form.waiting_word_kind, F.text.in_(ALL_MORE_DOCS_LABELS))
async def word_kind_more(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    cat = data.get("word_category", "physical")
    await m.answer(tr("msg_more_docs", lang), reply_markup=word_kind_kb(cat, more=True, lang=lang))


@dp.message(Form.waiting_word_kind, F.text.in_(ALL_BACK_CATEGORIES_LABELS))
async def word_kind_back(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(tr("msg_for_whom_doc", lang), reply_markup=word_category_kb(lang))
    await state.set_state(Form.waiting_word_category)


@dp.message(Form.waiting_word_kind)
async def word_kind(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    kind = LABEL_TO_KIND.get((m.text or "").strip())
    if not kind:
        data = await state.get_data()
        cat = data.get("word_category", "physical")
        await m.answer(tr("msg_pick_doc", lang), reply_markup=word_kind_kb(cat, lang=lang))
        return
    await state.update_data(word_kind=kind, extra="", extra_used=0)
    show_template = kind not in STUDY_KINDS
    hint = tr("msg_how_build_doc", lang)
    if show_template:
        hint += tr("msg_template_hint_suffix", lang)
    await m.answer(hint, reply_markup=mode_kb(show_template, lang=lang))
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


@dp.message(Form.waiting_word_mode, F.text.in_(ALL_BTN_AI_GENERATE_LABELS))
async def word_mode_ai(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(mode="ai", user_text="")
    data = await state.get_data()
    await m.answer(word_hint(data.get("word_kind", "doc"), "ai"), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_word_topic)


@dp.message(Form.waiting_word_mode, F.text.in_(ALL_BTN_OWN_TEXT_LABELS))
async def word_mode_user(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await state.update_data(mode="user")
    data = await state.get_data()
    await m.answer(word_hint(data.get("word_kind", "doc"), "user"), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_word_text)


@dp.message(Form.waiting_word_mode, F.text.in_(ALL_BTN_TEMPLATE_LABELS))
async def word_mode_template(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    kind = data.get("word_kind", "doc")
    uid = m.from_user.id
    if kind in STUDY_KINDS:
        await m.answer(tr("msg_no_template_study", lang), reply_markup=mode_kb(False, lang=lang))
        return
    ok, reason = start_job(uid)
    if not ok:
        await m.answer(reason)
        return
    await m.answer(tr("msg_building_template", lang), reply_markup=cancel_kb(lang))
    try:
        u = get_user(uid)
        tpl_path = f"/tmp/tpl_{uid}.docx"
        try:
            meta = json.loads(META_SCHEMAS.get(kind, META_SCHEMAS["doc"]))
        except Exception:
            meta = {}
        build_word(tpl_path, "Документ", template_sections_for(kind), kind, meta)
        await m.answer_document(FSInputFile(tpl_path), caption=tr("msg_template_caption", lang))
        try:
            os.remove(tpl_path)
        except Exception:
            pass
        u["generations"] += 1
        u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — шаблон: {KIND_LABELS.get(kind, kind)}")
        await m.answer(tr("msg_ready", lang), reply_markup=main_kb(lang))
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
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    kind = data.get("word_kind", "doc")
    if kind in WORD_SIZE_KINDS:
        await m.answer(tr("msg_which_size", lang), reply_markup=word_size_kb(lang))
        await state.set_state(Form.waiting_word_size)
    else:
        await word_build_draft(m, state, data, "short")


@dp.message(Form.waiting_word_topic)
async def word_topic(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    text = m.text or ""
    if len(text) > 500:
        await m.answer(tr("msg_too_long_500", lang))
        return
    await state.update_data(topic=text)
    await word_after_input(m, state)


@dp.message(Form.waiting_word_text)
async def word_user_text(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    text = m.text or ""
    if len(text) < 30:
        await m.answer(tr("msg_too_little_text_doc", lang))
        return
    if len(text) > 4000:
        await m.answer(tr("msg_too_long_4000", lang))
        return
    await state.update_data(user_text=text, topic=text[:80].replace("\n", " "))
    await word_after_input(m, state)


async def word_build_draft(m: Message, state: FSMContext, data: dict, size: str):
    lang = user_lang(m.from_user.id)
    await state.update_data(word_size=size)
    await m.answer(tr("msg_building_draft", lang), reply_markup=cancel_kb(lang))
    size_map = {"short": "поверхностное раскрытие темы — только суть и ключевые моменты, без глубокого разбора деталей и подпунктов, но по-настоящему содержательно, объём определяй по теме, не режь искусственно", "long": "полное раскрытие темы — подробно, с деталями, подпунктами и глубоким разбором, объём определяй по теме"}
    kind = data.get("word_kind", "doc")
    kind_name = WORD_KIND_DESC.get(kind, "документ")
    lang_instr = grok_lang_instruction(lang)
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
содержания). Без JSON.{lang_instr}"""
    sample = await ask_grok(prompt)
    if grok_failed(sample):
        await m.answer(
            tr("msg_grok_error", lang),
            reply_markup=word_size_kb(lang) if kind in WORD_SIZE_KINDS else word_confirm_kb(lang)
        )
        return
    await state.update_data(sample=sample)
    await send_draft(m, tr("msg_draft_ready_doc", lang, sample=sample), title=data.get("topic") or kind_name, reply_markup=word_confirm_kb(lang))
    await state.set_state(Form.waiting_word_confirm)


@dp.message(Form.waiting_word_size)
async def word_size(m: Message, state: FSMContext):
    t = (m.text or "").lower()
    size = "long" if "полн" in t or "full" in t or "voll" in t or "完整" in t or "الكامل" in t or "completo" in t or "complet" in t else "short"
    data = await state.get_data()
    await word_build_draft(m, state, data, size)


@dp.message(Form.waiting_word_confirm, F.text.in_(ALL_BTN_CHANGE_QUERY_LABELS))
async def word_change(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    if data.get("mode") == "user":
        await m.answer(tr("msg_new_text", lang), reply_markup=cancel_kb(lang))
        await state.set_state(Form.waiting_word_text)
    else:
        await m.answer(tr("msg_new_topic", lang), reply_markup=cancel_kb(lang))
        await state.set_state(Form.waiting_word_topic)


@dp.message(Form.waiting_word_confirm, F.text.in_(ALL_BTN_ADD_INFO_LABELS))
async def word_add(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    if (await state.get_data()).get("extra_used", 0) >= 3:
        await m.answer(tr("msg_extra_limit", lang), reply_markup=word_confirm_kb(lang))
        return
    await m.answer(tr("msg_extra_prompt", lang), reply_markup=cancel_kb(lang))
    await state.set_state(Form.waiting_word_extra)


@dp.message(Form.waiting_word_extra)
async def word_extra(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    text = m.text or ""
    if len(text) > 800:
        await m.answer(tr("msg_too_long_800", lang))
        return
    extra = ((data.get("extra") or "") + "\n" + text).strip()
    await state.update_data(extra=extra, extra_used=data.get("extra_used", 0) + 1)
    await m.answer(tr("msg_updating_draft", lang))
    lang_instr = grok_lang_instruction(lang)
    sample = await ask_grok(
        f"Обнови черновик документа.\nТема: {data.get('topic')}\nТекст: {data.get('user_text')}\nДоп: {extra}\nКороткий план. Без JSON.{lang_instr}"
    )
    if grok_failed(sample):
        await m.answer(
            tr("msg_draft_update_failed", lang),
            reply_markup=word_confirm_kb(lang)
        )
        await state.set_state(Form.waiting_word_confirm)
        return
    await state.update_data(sample=sample)
    await send_draft(m, tr("msg_draft_updated", lang, sample=sample), title=data.get("topic") or WORD_KIND_DESC.get(data.get("word_kind", "doc"), "документ"), reply_markup=word_confirm_kb(lang))
    await state.set_state(Form.waiting_word_confirm)


@dp.message(Form.waiting_word_confirm, F.text.in_(ALL_BTN_BUILD_DOC_LABELS))
async def word_build(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    data = await state.get_data()
    uid = m.from_user.id
    u = get_user(uid)
    ok, reason = start_job(uid)
    if not ok:
        await m.answer(reason)
        return
    await m.answer(tr("msg_building_word", lang))
    kind = data.get("word_kind", "doc")
    size = data.get("word_size", "short")
    lang_instr = grok_json_lang_instruction(lang)
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
            "\nПервым разделом в sections обязательно должен идти раздел «Содержание» - "
            "перечисли в его content все остальные разделы и подразделы документа по одному на строке "
            "(без номеров страниц, они не нужны в черновике). Не пропускай этот раздел.\n"
            "Заключение пиши как самостоятельные выводы по каждой задаче из введения "
            "и рекомендации по теме - не пересказывай содержание глав.\n"
            "Раздел со списком литературы должен содержать не менее 5 источников "
            "(автор/название/год/издательство или ссылка), оформленных по ГОСТ.\n"
            "ОБЯЗАТЕЛЬНО добавь хотя бы одну таблицу (ключ \"table\", см. схему ниже) в практическую/аналитическую "
            "главу - это часть требований к курсовой/реферату, а не опция. Собери в неё конкретные цифры по теме: "
            "результаты опроса, статистику, сравнение показателей до/после, данные по годам - придумай "
            "правдоподобные иллюстративные значения, если реальных нет под рукой (как и с остальным текстом "
            "черновика). Это требование не выполняется текстовым описанием цифр вместо таблицы."
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
Не создавай в sections отдельный раздел «Титульный лист» - обложка (организация, учебное заведение, ФИО
автора и руководителя, город, год) формируется отдельно из title и meta и так уже попадёт в документ;
первым разделом в sections должно идти «Введение» (или «Содержание», если оно есть в структуре документа).
Если в разделе уместна таблица (перечень товаров/имущества с ценой и количеством в договоре/акте, статистика
или практические данные в курсовой/реферате) - добавь в объект этого раздела ключ "table":
{{"headers":["...","..."],"rows":[["...","..."],["...","..."]]}}. Не выдумывай реальные официальные цифры
(даты, номера, суммы) для юридических документов - только то, что дал пользователь; для курсовых/рефератов
данные в таблице могут быть иллюстративными для черновика, как и остальной текст. Таблицы уместны не в каждом
разделе - добавляй только там, где это реально яснее текста, не в каждый раздел подряд.
Только JSON:
{{"title":"...","meta":{meta_schema},"sections":[{{"title":"Введение","content":"абзац1\n\nабзац2"}},{{"title":"...","content":"...","table":{{"headers":["...","..."],"rows":[["...","..."]]}}}}]}}{lang_instr}""", max_tokens=gen_max_tokens)
    try:
        content = extract_json(raw)
        if not isinstance(content.get("sections"), list) or not content["sections"]:
            raise ValueError("В ответе модели нет разделов документа")
    except Exception as e:
        print("Word JSON parse error:", e)
        finish_job(uid)
        await m.answer(tr("msg_no_text", lang, btn=WORD_KIND_LABELS.get(kind, {}).get(lang) or tr("msg_which_doc", lang)), reply_markup=main_kb(lang))
        await state.clear()
        return

    # Числовые ориентиры по объёму в промпте модель выполняет непоследовательно -
    # может уложиться в цель по одному разделу и заметно недобрать по другим,
    # так что итоговый документ всё равно выходит в разы короче нужного. Вместо
    # того чтобы полагаться на то, что промпт сработает с первого раза, меряем
    # фактический объём и, если он далеко от цели, просим модель дописать черновик
    # подробнее - раздел за разделом, с явной цифрой по каждому разделу отдельно
    # (общая просьба "дописать подробнее" на практике не даёт нужного прироста).
    min_total = WORD_MIN_TOTAL_WORDS.get((kind, size))
    section_min = WORD_SECTION_MIN_WORDS.get((kind, size))
    if min_total:
        actual_words = sum(len((b.get("content") or "").split()) for b in content.get("sections", []))
        attempts = 0
        while actual_words < min_total * 0.7 and attempts < 2:
            attempts += 1
            msg = tr("msg_writing_more", lang) if attempts == 1 else tr("msg_writing_a_bit_more", lang)
            await m.answer(msg)
            orig_sections = content.get("sections", [])
            targets_text = "\n".join(
                f"{i}. «{(b.get('title') or '').strip()}» — сейчас {len((b.get('content') or '').split())} слов, "
                + ("оставь как есть." if any(k in (b.get('title') or '').lower() for k in _WORD_SECTION_SKIP_EXPAND)
                   else f"нужно не менее {section_min} слов.")
                for i, b in enumerate(orig_sections, 1)
            )
            expand_raw = await ask_grok(f"""Вот черновик документа в JSON, он слишком короткий: сейчас примерно {actual_words} слов, а нужно суммарно не менее {min_total}.
Раздел за разделом, вот что нужно дописать:
{targets_text}
В документе должно остаться РОВНО {len(orig_sections)} разделов, с теми же заголовками, в том же порядке -
не удаляй, не объединяй, не переименовывай и не убирай ни один раздел (включая "Содержание" и заголовки глав
вроде "Глава 1"), только доработай содержимое (content) там, где указана недостача.
Добавляй конкретные факты, примеры, аргументы, анализ, детали по теме - без воды и без повторов одной мысли разными словами.
Это критично: не пересказывай своими словами то, что уже написано в разделе, и не добавляй общие рассуждения
не по существу - это одновременно снижает уникальность текста (антиплагиат видит шаблонные обороты как
совпадения с другими работами) и делает текст более узнаваемым как сгенерированный. Расширяй раздел только
за счёт нового содержания: дополнительные примеры, цифры, источники, конкретные детали по теме, которых
раньше не было.
{style_rule}
Верни ТОЛЬКО JSON той же структуры, без пояснений:
{{"title":"...","meta":{meta_schema},"sections":[{{"title":"...","content":"..."}}]}}{lang_instr}
Черновик:
{json.dumps(content, ensure_ascii=False)}""", max_tokens=gen_max_tokens)
            try:
                expanded = extract_json(expand_raw)
                expanded_sections = expanded.get("sections")
                expanded_words = sum(len((b.get("content") or "").split()) for b in expanded_sections or [])
                # Принимаем результат только если он не только длиннее, но и не потерял
                # разделы (модель иногда "экономит" объём, просто слив несколько
                # разделов в один или выкинув служебные заголовки при переписывании).
                if (isinstance(expanded_sections, list) and expanded_sections
                        and len(expanded_sections) >= len(orig_sections)
                        and expanded_words > actual_words):
                    content = expanded
                    actual_words = expanded_words
                else:
                    break  # результат хуже или структурно испорчен - не продолжаем попытки
            except Exception as e:
                print("Word expand parse error:", e)
                break

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

        fname = safe_filename(content.get("title"), fallback=WORD_KIND_DESC.get(kind, "Документ"), author=(content.get("meta") or {}).get("author"))
        await m.answer_document(FSInputFile(docx_path, filename=f"{fname}.docx"), caption=tr("msg_word_caption", lang))
        await m.answer_document(FSInputFile(pdf_path, filename=f"{fname}.pdf"), caption=tr("msg_pdf_caption", lang))
        u["generations"] += 1
        u["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title')}")
        for p in (docx_path, pdf_path):
            try:
                os.remove(p)
            except Exception:
                pass
        final_msg = tr("msg_doc_ready", lang)
        if kind not in STUDY_KINDS:
            final_msg += tr("msg_check_gaps", lang)
        await m.answer(final_msg, reply_markup=main_kb(lang))
        await state.clear()
    finally:
        finish_job(uid)


@dp.message(Form.waiting_word_confirm)
async def waiting_word_confirm_fallback(m: Message, state: FSMContext):
    lang = user_lang(m.from_user.id)
    await m.answer(
        tr("msg_didnt_understand", lang),
        reply_markup=word_confirm_kb(lang)
    )


@dp.message(F.text.in_(ALL_BTN_HISTORY_LABELS))
async def history(m: Message):
    lang = user_lang(m.from_user.id)
    u = get_user(m.from_user.id)
    await m.answer(tr("msg_history_empty", lang) if not u["history"] else tr("msg_history_title", lang) + "\n".join(u["history"][-10:]))


@dp.message(F.text.in_(ALL_BTN_PLAN_LABELS))
async def my_plan(m: Message):
    lang = user_lang(m.from_user.id)
    u = get_user(m.from_user.id)
    limit = PLAN_LIMITS.get(u["plan"], 15)
    await m.answer(tr("msg_plan_info", lang, used=u['generations'], limit=limit, left=max(0, limit - u['generations'])))


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
