import os
import asyncio
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile
from openai import AsyncOpenAI
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")

client = AsyncOpenAI(
    api_key=XAI_API_KEY,
    base_url="https://api.x.ai/v1"
)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

@dp.message(Command("start"))
async def start_command(message: Message):
    await message.answer(
        "Привет! Я бот для создания документов и презентаций.\n\n"
        "Просто напиши, что тебе нужно, например:\n"
        "• Сделай презентацию про маркетинг\n"
        "• Сделай документ: договор аренды\n\n"
        "Я создам файл и пришлю его тебе."
    )

@dp.message()
async def handle_message(message: Message):
    text = message.text.lower()
    
    await message.answer("Генерирую... Подожди 10–30 секунд.")

    try:
        if "презентац" in text:
            # Создаём презентацию
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            title = slide.shapes.title
            title.text = "Презентация"
            
            # Здесь потом добавим генерацию через Grok
            prs.save("presentation.pptx")
            await message.answer_document(FSInputFile("presentation.pptx"))
            
        else:
            # Создаём документ
            doc = Document()
            doc.add_heading("Документ", 0)
            doc.add_paragraph("Здесь будет сгенерированный текст.")
            doc.save("document.docx")
            await message.answer_document(FSInputFile("document.docx"))
            
    except Exception as e:
        await message.answer(f"Ошибка: {e}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
