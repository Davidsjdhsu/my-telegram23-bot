import os
import asyncio
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile
from openai import AsyncOpenAI
from docx import Document
from pptx import Presentation

BOT_TOKEN = os.getenv("BOT_TOKEN")
XAI_API_KEY = os.getenv("XAI_API_KEY")

if not BOT_TOKEN:
    raise ValueError("Не найден BOT_TOKEN")

if not XAI_API_KEY:
    raise ValueError("Не найден XAI_API_KEY")

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
        "• Сделай документ: договор аренды"
    )

@dp.message()
async def handle_message(message: Message):
    await message.answer("Генерирую... Подожди немного.")

    try:
        text = message.text.lower()

        if "презентац" in text:
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[0])
            slide.shapes.title.text = "Презентация"
            prs.save("presentation.pptx")
            await message.answer_document(FSInputFile("presentation.pptx"))
        else:
            doc = Document()
            doc.add_heading("Документ", 0)
            doc.add_paragraph("Здесь будет текст.")
            doc.save("document.docx")
            await message.answer_document(FSInputFile("document.docx"))

    except Exception as e:
        await message.answer(f"Ошибка: {e}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
