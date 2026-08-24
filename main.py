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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ====================== НАСТРОЙКИ ======================
BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
ADMIN_IDS = [123456789]  # ← СВОЙ TELEGRAM ID

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
    if plan in ["premium"]:
        level = """
Ты — топовый презентационный дизайнер и стратег.
Сделай презентацию профессионального уровня:
- Сильные, цепляющие заголовки
- ёмкий и полезный текст
- Чёткая логика и структура
- Каждый слайд должен нести ценность
"""
    elif plan in ["standard"]:
        level = """
Сделай качественную презентацию хорошего уровня:
- Понятные и сильные заголовки
- Достаточно подробный текст
- Логичная структура
