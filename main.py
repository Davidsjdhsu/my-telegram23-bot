import os
import asyncio
import json
from datetime import datetime, timedelta
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from openai import AsyncOpenAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ====================== НАСТРОЙКИ ======================
BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")
ADMIN_IDS = [123456789]  # ← СЮДА ВПИШИ СВОЙ TELEGRAM ID (узнать можно у @userinfobot)

client = AsyncOpenAI(api_key=XAI_API_KEY, base_url="https://api.x.ai/v1")
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)

# Хранилище пользователей (в тестовом режиме — в памяти)
users_db = {}  # {user_id: {"name": "", "category": "", "plan": "free", "generations": 0, "history": []}}

# Лимиты тарифов
PLAN_LIMITS = {
    "free": 3,
    "basic": 10,      # 199 ₽
    "standard": 40,   # 499 ₽
    "premium": 100    # 999 ₽
}

# ====================== СОСТОЯНИЯ ======================
class Form(StatesGroup):
    waiting_category = State()
    waiting_confirm = State()
    waiting_document_data = State()

# ====================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ======================
async def ask_grok(prompt: str) -> str:
    try:
        response = await client.chat.completions.create(
            model="grok-3",
            messages=[
                {"role": "system", "content": "Ты профессиональный помощник по созданию презентаций и документов на русском языке. Отвечай чётко и структурированно."},
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
    limit = PLAN_LIMITS.get(user["plan"], 3)
    return user["generations"] < limit

# ====================== КЛАВИАТУРЫ ======================
def category_keyboard():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="👨‍🎓 Студент")],
            [KeyboardButton(text="💼 Предприниматель")],
            [KeyboardButton(text="👤 Обычный человек")]
        ],
        resize_keyboard=True
    )

def main_menu_keyboard(category: str):
    buttons = [
        [KeyboardButton(text="📊 Сделать презентацию")],
        [KeyboardButton(text="📄 Сделать документ")],
        [KeyboardButton(text="📋 Готовые шаблоны")],
        [KeyboardButton(text="📁 Моя история")],
        [KeyboardButton(text="ℹ️ Мой тариф")]
    ]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)

# ====================== /START ======================
@dp.message(Command("start"))
async def start_command(message: Message, state: FSMContext):
    user_id = message.from_user.id
    user = get_user(user_id)
    
    name = message.from_user.first_name or "друг"
    user["name"] = name

    await message.answer(
        f"Привет, {name}! 👋\n\n"
        "Я бот, который помогает создавать презентации, документы и заполнять готовые шаблоны.\n\n"
        "Чтобы я мог лучше подсказать, кем ты себя считаешь?",
        reply_markup=category_keyboard()
    )
    await state.set_state(Form.waiting_category)

# ====================== ВЫБОР КАТЕГОРИИ ======================
@dp.message(Form.waiting_category)
async def process_category(message: Message, state: FSMContext):
    user_id = message.from_user.id
    user = get_user(user_id)
    text = message.text

    if "Студент" in text:
        user["category"] = "student"
        desc = "Отлично! Я помогу с презентациями, рефератами и учебными документами."
    elif "Предприниматель" in text:
        user["category"] = "business"
        desc = "Отлично! Я помогу с презентациями, коммерческими предложениями, карточками компании и договорами."
    elif "Обычный" in text:
        user["category"] = "regular"
        desc = "Отлично! Я помогу с презентациями, договорами и готовыми документами (европротокол, купля-продажа и т.д.)."
    else:
        await message.answer("Пожалуйста, выбери один из вариантов кнопками.")
        return

    await message.answer(
        f"{desc}\n\nЧто хочешь сделать?",
        reply_markup=main_menu_keyboard(user["category"])
    )
    await state.clear()

# ====================== ГЛАВНОЕ МЕНЮ ======================
@dp.message(F.text == "📊 Сделать презентацию")
async def make_presentation(message: Message, state: FSMContext):
    if not can_generate(message.from_user.id):
        await message.answer("Лимит генераций на этом тарифе закончился. Напиши админу.")
        return

    await message.answer(
        "Напиши тему и детали презентации.\n\n"
        "Пример:\n"
        "«Сделай презентацию про хомяков на 10 слайдов в строгом тёмно-сером стиле»"
    )
    await state.set_state(Form.waiting_confirm)
    await state.update_data(action="presentation")

@dp.message(F.text == "📄 Сделать документ")
async def make_document(message: Message, state: FSMContext):
    if not can_generate(message.from_user.id):
        await message.answer("Лимит генераций закончился.")
        return

    await message.answer(
        "Напиши, какой документ нужен.\n\n"
        "Пример:\n"
        "«Сделай коммерческое предложение для IT-компании»"
    )
    await state.set_state(Form.waiting_confirm)
    await state.update_data(action="document")

@dp.message(F.text == "📋 Готовые шаблоны")
async def templates_menu(message: Message):
    await message.answer(
        "Доступные шаблоны (тестовый режим):\n\n"
        "1. Европротокол (ДТП)\n"
        "2. Договор купли-продажи автомобиля\n"
        "3. Карточка предприятия\n"
        "4. Акт выполненных работ\n\n"
        "Пока это демо. Напиши название шаблона, который хочешь заполнить."
    )

@dp.message(F.text == "📁 Моя история")
async def show_history(message: Message):
    user = get_user(message.from_user.id)
    if not user["history"]:
        await message.answer("История пока пустая.")
        return
    text = "Твоя история:\n\n"
    for i, item in enumerate(user["history"][-10:], 1):
        text += f"{i}. {item}\n"
    await message.answer(text)

@dp.message(F.text == "ℹ️ Мой тариф")
async def my_plan(message: Message):
    user = get_user(message.from_user.id)
    plan_names = {
        "free": "Бесплатный",
        "basic": "Базовый (199 ₽)",
        "standard": "Стандарт (499 ₽)",
        "premium": "Премиум (999 ₽)"
    }
    limit = PLAN_LIMITS.get(user["plan"], 3)
    await message.answer(
        f"Твой тариф: {plan_names.get(user['plan'])}\n"
        f"Использовано генераций: {user['generations']} из {limit}"
    )

# ====================== ОБРАБОТКА ЗАПРОСА ======================
@dp.message(Form.waiting_confirm)
async def process_request(message: Message, state: FSMContext):
    data = await state.get_data()
    action = data.get("action")
    user_id = message.from_user.id
    user = get_user(user_id)

    await message.answer("Делаю образец... Подожди 10–20 секунд.")

    if action == "presentation":
        prompt = f"""
        Пользователь попросил: {message.text}
        
        Сделай краткий образец структуры презентации:
        - Название
        - Список слайдов (номер + заголовок)
        - Рекомендуемый стиль
        
        Ответ должен быть коротким и понятным.
        """
        sample = await ask_grok(prompt)
        
        await state.update_data(original_request=message.text, sample=sample, action="presentation")
        await message.answer(
            f"Вот образец:\n\n{sample}\n\n"
            "Если подходит — напиши «делай»\n"
            "Если нет — напиши новый запрос."
        )
    else:
        # Для документов пока упрощённо
        await message.answer("Генерирую документ...")
        # Здесь можно добавить генерацию docx
        user["generations"] += 1
        await message.answer("Документ будет готов в следующей версии. Пока используй презентации.")
        await state.clear()

# ====================== ПОДТВЕРЖДЕНИЕ И ГЕНЕРАЦИЯ ======================
@dp.message(F.text.lower().in_(["делай", "да", "подтверждаю", "ок", "хорошо"]))
async def confirm_and_generate(message: Message, state: FSMContext):
    data = await state.get_data()
    if not data.get("original_request"):
        await message.answer("Сначала напиши тему презентации.")
        return

    user_id = message.from_user.id
    user = get_user(user_id)

    if not can_generate(user_id):
        await message.answer("Лимит генераций закончился.")
        return

    await message.answer("Делаю полную версию... Это займёт 20–40 секунд.")

    # Генерация полного контента
    prompt = f"""
    Создай полную структуру презентации по запросу:
    {data['original_request']}

    Верни ответ строго в JSON формате:
    {{
        "title": "Название презентации",
        "slides": [
            {{"title": "Заголовок", "content": "Текст слайда"}}
        ]
    }}
    """
    raw = await ask_grok(prompt)

    try:
        start = raw.find("{")
        end = raw.rfind("}") + 1
        content = json.loads(raw[start:end])
    except:
        await message.answer("Не удалось разобрать ответ. Попробуй ещё раз.")
        return

    # Создаём PPTX
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Титульный
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RgbColor(40, 40, 40)

    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.8), Inches(11.3), Inches(1.5))
    p = title_box.text_frame.paragraphs[0]
    p.text = content.get("title", "Презентация")
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RgbColor(230, 230, 230)
    p.alignment = PP_ALIGN.CENTER

    # Остальные слайды
    for s in content.get("slides", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RgbColor(45, 45, 45)

        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(12), Inches(1))
        p = title_box.text_frame.paragraphs[0]
        p.text = s.get("title", "")
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RgbColor(230, 230, 230)

        content_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(12), Inches(5))
        tf = content_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = s.get("content", "")
        p.font.size = Pt(18)
        p.font.color.rgb = RgbColor(200, 200, 200)

    pptx_path = f"pres_{user_id}.pptx"
    prs.save(pptx_path)

    # PDF
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
        y -= 20
        c.setFont("Helvetica", 10)
        text = s.get("content", "")[:250]
        c.drawString(40, y, text[:95])
        y -= 15
        if len(text) > 95:
            c.drawString(40, y, text[95:190])
            y -= 15
        y -= 15

    c.save()

    # Отправка
    await message.answer_document(FSInputFile(pptx_path), caption="Редактируемая версия (PPTX)")
    await message.answer_document(FSInputFile(pdf_path), caption="Версия PDF")

    user["generations"] += 1
    user["history"].append(f"{datetime.now().strftime('%d.%m %H:%M')} — {content.get('title', 'Презентация')}")

    await message.answer("Готово!", reply_markup=main_menu_keyboard(user.get("category")))
    await state.clear()

# ====================== АДМИНКА ======================
@dp.message(Command("grant"))
async def admin_grant(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    try:
        parts = message.text.split()
        target_id = int(parts[1])
        plan = parts[2] if len(parts) > 2 else "premium"
        user = get_user(target_id)
        user["plan"] = plan
        await message.answer(f"Пользователю {target_id} выдан тариф: {plan}")
    except:
        await message.answer("Использование: /grant user_id [basic/standard/premium]")

@dp.message(Command("revoke"))
async def admin_revoke(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    try:
        target_id = int(message.text.split()[1])
        user = get_user(target_id)
        user["plan"] = "free"
        await message.answer(f"Доступ пользователя {target_id} сброшен на free")
    except:
        await message.answer("Использование: /revoke user_id")

@dp.message(Command("users"))
async def admin_users(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    text = "Пользователи:\n\n"
    for uid, data in list(users_db.items())[:30]:
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
