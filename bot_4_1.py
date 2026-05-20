import asyncio
import base64
import io
import logging
import os
import urllib.parse

import aiohttp
import openpyxl
from anthropic import AsyncAnthropic
from aiogram import Bot, Dispatcher, F
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, BufferedInputFile
from pptx import Presentation
from pptx.util import Inches, Pt

# ─── Настройки ────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "YOUR_ANTHROPIC_API_KEY")
PROXY_URL = os.getenv("PROXY_URL", "")

# ─── Инициализация ────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

if PROXY_URL:
    session = AiohttpSession(proxy=PROXY_URL)
    bot = Bot(token=TELEGRAM_TOKEN, session=session)
    logger.info(f"Используется прокси: {PROXY_URL}")
else:
    bot = Bot(token=TELEGRAM_TOKEN)

dp = Dispatcher()
anthropic = AsyncAnthropic(api_key=ANTHROPIC_API_KEY)

conversation_history: dict[int, list[dict]] = {}
MAX_HISTORY = 20

SYSTEM_PROMPT = """Ты умный AI-ассистент в Telegram с расширенными возможностями.

ВАЖНО: Ты УМЕЕШЬ генерировать изображения через встроенный инструмент! 
Когда тебя спрашивают умеешь ли ты рисовать/генерировать картинки — отвечай ДА.

Твои возможности:
- Отвечать на вопросы
- Генерировать изображения по описанию (команда: "нарисуй...")  
- Распознавать и анализировать картинки
- Создавать Excel таблицы
- Создавать презентации PowerPoint

Если пользователь пишет на русском — отвечай на русском.
если на английском — на английском.

Когда пользователь просит создать Excel таблицу — отвечай ТОЛЬКО в формате:
EXCEL_TABLE:
Заголовок1|Заголовок2|Заголовок3
Данные1|Данные2|Данные3

Когда пользователь просит создать презентацию — отвечай ТОЛЬКО в формате:
PRESENTATION:
TITLE:Название презентации
SLIDE:Заголовок слайда|Текст содержимого слайда
SLIDE:Заголовок 2|Текст 2
"""


# ─── Вспомогательные функции ──────────────────────────────────────────────────
def is_image_request(text: str) -> bool:
    keywords = ["нарисуй", "сгенерируй картинку", "создай изображение",
                "generate image", "draw", "создай картинку", "нарисуй мне",
                "картинку с", "изображение с"]
    return any(kw in text.lower() for kw in keywords)


def is_excel_request(text: str) -> bool:
    keywords = ["excel", "таблицу", "xlsx", "создай таблицу", "сделай таблицу"]
    return any(kw in text.lower() for kw in keywords)


def is_presentation_request(text: str) -> bool:
    keywords = ["презентацию", "powerpoint", "pptx", "слайды", "создай презентацию"]
    return any(kw in text.lower() for kw in keywords)


async def generate_image(prompt: str, retries: int = 3) -> bytes | None:
    encoded = urllib.parse.quote(prompt, safe='')
    url = f"https://image.pollinations.ai/prompt/{encoded}?width=1024&height=1024&nologo=true"

    for attempt in range(1, retries + 1):
        try:
            async with aiohttp.ClientSession() as s:
                async with s.get(url, timeout=aiohttp.ClientTimeout(total=90)) as resp:
                    if resp.status == 200:
                        data = await resp.read()
                        content_type = resp.headers.get("Content-Type", "")
                        if data and content_type.startswith("image"):
                            return data
                        logger.warning(f"Попытка {attempt}: не картинка (Content-Type={content_type})")
                    else:
                        logger.warning(f"Попытка {attempt}: статус {resp.status}")
        except asyncio.TimeoutError:
            logger.warning(f"Попытка {attempt}: таймаут")
        except Exception as e:
            logger.error(f"Попытка {attempt}: {type(e).__name__}: {e}")

        if attempt < retries:
            await asyncio.sleep(2 * attempt)

    return None


def create_excel(text: str) -> bytes | None:
    try:
        lines = []
        in_table = False
        for line in text.split("\n"):
            if "EXCEL_TABLE:" in line:
                in_table = True
                continue
            if in_table and "|" in line:
                lines.append(line.strip())

        if not lines:
            return None

        wb = openpyxl.Workbook()
        ws = wb.active
        for i, line in enumerate(lines, 1):
            cells = line.split("|")
            for j, cell in enumerate(cells, 1):
                ws.cell(row=i, column=j, value=cell.strip())
                if i == 1:
                    ws.cell(row=i, column=j).font = openpyxl.styles.Font(bold=True)

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_len + 4

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()
    except Exception as e:
        logger.error(f"Ошибка создания Excel: {e}")
        return None


def create_presentation(text: str) -> bytes | None:
    try:
        title = "Презентация"
        slides_data = []

        for line in text.split("\n"):
            line = line.strip()
            if line.startswith("TITLE:"):
                title = line[6:]
            elif line.startswith("SLIDE:"):
                parts = line[6:].split("|", 1)
                slide_title = parts[0] if parts else ""
                slide_content = parts[1] if len(parts) > 1 else ""
                slides_data.append((slide_title, slide_content))

        if not slides_data:
            return None

        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[0])
        slide.shapes.title.text = title
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = "Создано AI-ассистентом"

        for slide_title, slide_content in slides_data:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = slide_title
            if len(slide.placeholders) > 1:
                slide.placeholders[1].text = slide_content

        buf = io.BytesIO()
        prs.save(buf)
        return buf.getvalue()
    except Exception as e:
        logger.error(f"Ошибка создания презентации: {e}")
        return None


async def ask_claude(user_id: int, content) -> str:
    if user_id not in conversation_history:
        conversation_history[user_id] = []

    conversation_history[user_id].append({"role": "user", "content": content})

    if len(conversation_history[user_id]) > MAX_HISTORY:
        conversation_history[user_id] = conversation_history[user_id][-MAX_HISTORY:]

    response = await anthropic.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=conversation_history[user_id]
    )

    reply = response.content[0].text
    conversation_history[user_id].append({"role": "assistant", "content": reply})
    return reply


# ─── Хэндлеры ─────────────────────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: Message):
    conversation_history[message.from_user.id] = []
    await message.answer(
        "👋 Привет! Я AI-ассистент на базе Claude.\n\n"
        "🔧 Я умею:\n"
        "• 💬 Отвечать на вопросы\n"
        "• 🖼 Распознавать картинки и решать задачи с фото\n"
        "• 🎨 Генерировать изображения по описанию\n"
        "• 📊 Создавать Excel таблицы\n"
        "• 📋 Создавать презентации PowerPoint\n\n"
        "📌 Команды:\n"
        "/start — начать заново\n"
        "/clear — очистить историю\n"
        "/help — примеры запросов"
    )


@dp.message(Command("clear"))
async def cmd_clear(message: Message):
    conversation_history[message.from_user.id] = []
    await message.answer("🗑️ История очищена!")


@dp.message(Command("help"))
async def cmd_help(message: Message):
    await message.answer(
        "🤖 Примеры запросов:\n\n"
        "🎨 Генерация картинки:\n"
        "«Нарисуй закат над морем»\n\n"
        "📊 Excel таблица:\n"
        "«Создай таблицу с расходами за месяц»\n\n"
        "📋 Презентация:\n"
        "«Создай презентацию о космосе на 5 слайдов»\n\n"
        "🖼 Анализ картинки:\n"
        "Просто отправь фото с подписью или без!"
    )


@dp.message(F.photo)
async def handle_photo(message: Message):
    await bot.send_chat_action(message.chat.id, "typing")
    try:
        photo = message.photo[-1]
        file = await bot.get_file(photo.file_id)
        buf = io.BytesIO()
        await bot.download_file(file.file_path, buf)
        photo_bytes = buf.getvalue()

        photo_b64 = base64.standard_b64encode(photo_bytes).decode()
        caption = message.caption or "Что на этом изображении? Опиши подробно. Если это задача или уравнение — реши его."

        content = [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": photo_b64}},
            {"type": "text", "text": caption}
        ]

        reply = await ask_claude(message.from_user.id, content)
        await message.answer(reply)

    except Exception as e:
        logger.error(f"Ошибка обработки фото: {e}")
        await message.answer("❌ Не удалось обработать изображение. Попробуй ещё раз.")


@dp.message(F.text)
async def handle_text(message: Message):
    user_id = message.from_user.id
    text = message.text

    await bot.send_chat_action(message.chat.id, "typing")

    try:
        if is_image_request(text):
            await message.answer("🎨 Генерирую изображение, подожди 10-20 секунд...")
            await bot.send_chat_action(message.chat.id, "upload_photo")
            image_bytes = await generate_image(text)
            if image_bytes:
                await message.answer_photo(
                    BufferedInputFile(image_bytes, filename="image.jpg"),
                    caption="✅ Готово!"
                )
            else:
                await message.answer("❌ Не удалось сгенерировать изображение. Попробуй ещё раз.")
            return

        if is_excel_request(text) or is_presentation_request(text):
            reply = await ask_claude(user_id, text)

            if "EXCEL_TABLE:" in reply:
                excel_bytes = create_excel(reply)
                if excel_bytes:
                    await message.answer_document(
                        BufferedInputFile(excel_bytes, filename="таблица.xlsx"),
                        caption="📊 Вот твоя Excel таблица!"
                    )
                    return

            if "PRESENTATION:" in reply:
                pptx_bytes = create_presentation(reply)
                if pptx_bytes:
                    await message.answer_document(
                        BufferedInputFile(pptx_bytes, filename="презентация.pptx"),
                        caption="📋 Вот твоя презентация!"
                    )
                    return

            await message.answer(reply)
            return

        reply = await ask_claude(user_id, text)
        await message.answer(reply)

    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await message.answer("❌ Произошла ошибка. Попробуй ещё раз.")


# ─── Запуск ───────────────────────────────────────────────────────────────────
async def main():
    logger.info("Бот запускается...")
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
