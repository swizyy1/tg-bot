import asyncio
import base64
import io
import logging
import os
import sqlite3
import threading
import urllib.parse
from datetime import datetime, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler

import aiohttp
import openpyxl
from anthropic import AsyncAnthropic
from aiogram import Bot, Dispatcher, F
from aiogram.client.session.aiohttp import AiohttpSession
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    Message, BufferedInputFile,
    LabeledPrice, PreCheckoutQuery,
    InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton
)
from pptx import Presentation

# ─── Настройки ────────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "YOUR_TELEGRAM_BOT_TOKEN")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "YOUR_ANTHROPIC_API_KEY")
PROXY_URL = os.getenv("PROXY_URL", "")
WOLFRAM_API_KEY = os.getenv("WOLFRAM_API_KEY", "")

# Цена подписки в Telegram Stars (1 Star ≈ 0.013$, 250 Stars ≈ ~3$)
SUBSCRIPTION_PRICE_STARS = 250
SUBSCRIPTION_DAYS = 30
FREE_MESSAGES_PER_DAY = 10
TRIAL_DAYS = 3
DAILY_BONUS_MESSAGES = 3
MAX_BONUS_MESSAGES = 15  # максимум накопленных бонусных сообщений
MAX_HISTORY = 20
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "0").split(",") if x and x != "0"]

# ─── Инициализация ────────────────────────────────────────────────────────────
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

if PROXY_URL:
    session = AiohttpSession(proxy=PROXY_URL)
    bot = Bot(token=TELEGRAM_TOKEN, session=session)
else:
    bot = Bot(token=TELEGRAM_TOKEN)

dp = Dispatcher()
anthropic = AsyncAnthropic(api_key=ANTHROPIC_API_KEY)

# ─── Health-сервер для Render ─────────────────────────────────────────────────
class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"OK")
    def log_message(self, *args):
        pass  # Отключаем логи

def run_health_server():
    server = HTTPServer(("0.0.0.0", 8080), HealthHandler)
    server.serve_forever()


SYSTEM_PROMPT_BASE = """Ты умный AI-ассистент в Telegram с расширенными возможностями.

ВАЖНО: Ты УМЕЕШЬ генерировать изображения через встроенный инструмент! 
Когда тебя спрашивают умеешь ли ты рисовать/генерировать картинки — отвечай ДА.

Твои возможности:
- Отвечать на вопросы
- Генерировать изображения по описанию (команда: "нарисуй...")  
- Распознавать и анализировать картинки
- Создавать Excel таблицы
- Создавать презентации PowerPoint
- Точно решать математические задачи

При решении математических и геометрических задач:
1. Всегда решай строго пошагово
2. Записывай все формулы которые используешь
3. Подставляй числа явно на каждом шаге
4. В конце проверяй ответ подстановкой
5. Если задача геометрическая — сначала опиши фигуру и все известные элементы

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

STYLE_PROMPTS = {
    "formal": "СТИЛЬ ОБЩЕНИЯ: Отвечай официально и вежливо. Используй грамотный литературный язык. Обращайся на «вы».",
    "friend": "СТИЛЬ ОБЩЕНИЯ: Общайся как близкий друг — неформально, живо, с юмором и эмодзи. Говори просто, можешь шутить, используй разговорный язык. Обращайся на «ты».",
    "short": "СТИЛЬ ОБЩЕНИЯ: Отвечай максимально кратко и по делу. Только суть, никакой воды. Минимум слов — максимум смысла.",
}

STYLE_NAMES = {
    "formal": "🎩 Формальный",
    "friend": "😊 Как с другом",
    "short": "⚡ Краткий",
}


def get_system_prompt(user_id: int) -> str:
    style = get_user_style(user_id)
    style_text = STYLE_PROMPTS.get(style, STYLE_PROMPTS["friend"])
    return SYSTEM_PROMPT_BASE + "\n" + style_text


# ─── База данных ──────────────────────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()

    # Пользователи и подписки
    c.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            username TEXT,
            created_at TEXT,
            subscription_until TEXT,
            messages_today INTEGER DEFAULT 0,
            last_message_date TEXT,
            referred_by INTEGER DEFAULT NULL,
            referral_count INTEGER DEFAULT 0,
            notified_expiry INTEGER DEFAULT 0,
            trial_used INTEGER DEFAULT 0,
            bonus_messages INTEGER DEFAULT 0,
            last_bonus_date TEXT
        )
    """)

    # Добавляем колонки если их нет (для существующих БД)
    for col in ["referred_by INTEGER DEFAULT NULL",
                "referral_count INTEGER DEFAULT 0",
                "notified_expiry INTEGER DEFAULT 0",
                "trial_used INTEGER DEFAULT 0",
                "bonus_messages INTEGER DEFAULT 0",
                "last_bonus_date TEXT",
                "speech_style TEXT DEFAULT 'friend'"]:
        try:
            c.execute(f"ALTER TABLE users ADD COLUMN {col}")
        except Exception:
            pass

    # История сообщений (сохраняется между сессиями)
    c.execute("""
        CREATE TABLE IF NOT EXISTS message_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            role TEXT,
            content TEXT,
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()


def get_user(user_id: int) -> dict | None:
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return {
            "user_id": row[0],
            "username": row[1],
            "created_at": row[2],
            "subscription_until": row[3],
            "messages_today": row[4],
            "last_message_date": row[5],
            "referred_by": row[6] if len(row) > 6 else None,
            "referral_count": row[7] if len(row) > 7 else 0,
            "notified_expiry": row[8] if len(row) > 8 else 0,
            "trial_used": row[9] if len(row) > 9 else 0,
            "bonus_messages": row[10] if len(row) > 10 else 0,
            "last_bonus_date": row[11] if len(row) > 11 else None,
        }
    return None


def create_user(user_id: int, username: str, referred_by: int = None):
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    now = datetime.now().isoformat()
    c.execute(
        "INSERT OR IGNORE INTO users (user_id, username, created_at, messages_today, last_message_date, referred_by) VALUES (?, ?, ?, 0, ?, ?)",
        (user_id, username, now, datetime.now().date().isoformat(), referred_by)
    )
    # Если новый пользователь пришёл по реферальной ссылке — увеличиваем счётчик рефералов
    if referred_by and c.lastrowid:
        c.execute(
            "UPDATE users SET referral_count = referral_count + 1 WHERE user_id = ?",
            (referred_by,)
        )
        # Проверяем нужно ли давать бесплатный месяц (каждые 5 рефералов)
        c.execute("SELECT referral_count FROM users WHERE user_id = ?", (referred_by,))
        row = c.fetchone()
        if row and row[0] % 5 == 0:
            conn.commit()
            conn.close()
            # Даём бесплатный месяц рефереру
            activate_subscription(referred_by)
            return True  # сигнал что реферер получил награду
    conn.commit()
    conn.close()
    return False


def get_referral_count(user_id: int) -> int:
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("SELECT referral_count FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else 0


def get_user_style(user_id: int) -> str:
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("SELECT speech_style FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    conn.close()
    return row[0] if row and row[0] else "friend"


def set_user_style(user_id: int, style: str):
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("UPDATE users SET speech_style = ? WHERE user_id = ?", (style, user_id))
    conn.commit()
    conn.close()


def activate_trial(user_id: int):
    """Активируем пробный период 3 дня."""
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    trial_until = (datetime.now() + timedelta(days=TRIAL_DAYS)).isoformat()
    c.execute(
        "UPDATE users SET subscription_until = ?, trial_used = 1 WHERE user_id = ? AND trial_used = 0",
        (trial_until, user_id)
    )
    affected = c.rowcount
    conn.commit()
    conn.close()
    return affected > 0


def is_trial_used(user_id: int) -> bool:
    user = get_user(user_id)
    if not user:
        return False
    return bool(user.get("trial_used", 0))


def claim_daily_bonus(user_id: int) -> bool:
    """Даём ежедневный бонус +3 сообщения. Возвращает True если бонус выдан."""
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    today = datetime.now().date().isoformat()
    c.execute("SELECT last_bonus_date, bonus_messages FROM users WHERE user_id = ?", (user_id,))
    row = c.fetchone()
    if not row or row[0] == today:
        conn.close()
        return False
    current_bonus = row[1] or 0
    if current_bonus >= MAX_BONUS_MESSAGES:
        # Бонус накоплен до максимума — не добавляем
        c.execute("UPDATE users SET last_bonus_date = ? WHERE user_id = ?", (today, user_id))
        conn.commit()
        conn.close()
        return False
    new_bonus = min(current_bonus + DAILY_BONUS_MESSAGES, MAX_BONUS_MESSAGES)
    c.execute(
        "UPDATE users SET bonus_messages = ?, last_bonus_date = ? WHERE user_id = ?",
        (new_bonus, today, user_id)
    )
    conn.commit()
    conn.close()
    return True


def get_users_expiring_soon() -> list:
    """Пользователи чья подписка истекает через 3 дня."""
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    now = datetime.now()
    in_3_days = (now + timedelta(days=3)).isoformat()
    tomorrow = (now + timedelta(days=1)).isoformat()
    c.execute("""
        SELECT user_id, username, subscription_until FROM users
        WHERE subscription_until BETWEEN ? AND ?
        AND notified_expiry = 0
    """, (tomorrow, in_3_days))
    rows = c.fetchall()
    conn.close()
    return rows


def mark_notified(user_id: int):
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("UPDATE users SET notified_expiry = 1 WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()


def reset_notified_expired():
    """Сбрасываем флаг уведомления для истёкших подписок."""
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    now = datetime.now().isoformat()
    c.execute("UPDATE users SET notified_expiry = 0 WHERE subscription_until < ?", (now,))
    conn.commit()
    conn.close()


def is_subscribed(user_id: int) -> bool:
    user = get_user(user_id)
    if not user or not user["subscription_until"]:
        return False
    until = datetime.fromisoformat(user["subscription_until"])
    return until > datetime.now()


def get_subscription_until(user_id: int) -> str | None:
    user = get_user(user_id)
    if not user or not user["subscription_until"]:
        return None
    until = datetime.fromisoformat(user["subscription_until"])
    if until > datetime.now():
        return until.strftime("%d.%m.%Y")
    return None


def activate_subscription(user_id: int):
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    user = get_user(user_id)
    if user and user["subscription_until"]:
        current = datetime.fromisoformat(user["subscription_until"])
        if current > datetime.now():
            new_until = current + timedelta(days=SUBSCRIPTION_DAYS)
        else:
            new_until = datetime.now() + timedelta(days=SUBSCRIPTION_DAYS)
    else:
        new_until = datetime.now() + timedelta(days=SUBSCRIPTION_DAYS)
    c.execute(
        "UPDATE users SET subscription_until = ? WHERE user_id = ?",
        (new_until.isoformat(), user_id)
    )
    conn.commit()
    conn.close()


def can_send_message(user_id: int) -> tuple[bool, int]:
    """Возвращает (может ли отправить, осталось сообщений)"""
    if is_subscribed(user_id):
        return True, -1  # -1 = безлимит

    user = get_user(user_id)
    if not user:
        return False, 0

    today = datetime.now().date().isoformat()
    bonus = user.get("bonus_messages", 0) or 0
    total_limit = FREE_MESSAGES_PER_DAY + bonus

    if user["last_message_date"] != today:
        conn = sqlite3.connect("bot.db")
        c = conn.cursor()
        c.execute(
            "UPDATE users SET messages_today = 0, last_message_date = ? WHERE user_id = ?",
            (today, user_id)
        )
        conn.commit()
        conn.close()
        return True, total_limit

    remaining = total_limit - user["messages_today"]
    return remaining > 0, max(0, remaining)


def increment_message_count(user_id: int):
    today = datetime.now().date().isoformat()
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute(
        "UPDATE users SET messages_today = messages_today + 1, last_message_date = ? WHERE user_id = ?",
        (today, user_id)
    )
    conn.commit()
    conn.close()


# ─── История сообщений (персистентная) ───────────────────────────────────────
def load_history(user_id: int) -> list[dict]:
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute(
        "SELECT role, content FROM message_history WHERE user_id = ? ORDER BY id DESC LIMIT ?",
        (user_id, MAX_HISTORY)
    )
    rows = c.fetchall()
    conn.close()
    # Возвращаем в правильном порядке (старые сначала)
    return [{"role": r[0], "content": r[1]} for r in reversed(rows)]


def save_message(user_id: int, role: str, content: str):
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    now = datetime.now().isoformat()
    c.execute(
        "INSERT INTO message_history (user_id, role, content, created_at) VALUES (?, ?, ?, ?)",
        (user_id, role, content if isinstance(content, str) else str(content), now)
    )
    # Удаляем старые сообщения, оставляем только MAX_HISTORY*2
    c.execute("""
        DELETE FROM message_history WHERE id IN (
            SELECT id FROM message_history WHERE user_id = ?
            ORDER BY id DESC LIMIT -1 OFFSET ?
        )
    """, (user_id, MAX_HISTORY * 2))
    conn.commit()
    conn.close()


def clear_history(user_id: int):
    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("DELETE FROM message_history WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()


def bottom_menu() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="📋 Меню")]],
        resize_keyboard=True,
        persistent=True
    )


def main_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="💳 Подписка", callback_data="menu_subscribe"),
            InlineKeyboardButton(text="📊 Статус", callback_data="menu_status"),
        ],
        [
            InlineKeyboardButton(text="👥 Пригласить друга", callback_data="menu_referral"),
            InlineKeyboardButton(text="🗑️ Очистить историю", callback_data="menu_clear"),
        ],
        [
            InlineKeyboardButton(text="🎙️ Настройки речи", callback_data="menu_speech"),
            InlineKeyboardButton(text="❓ Помощь", callback_data="menu_help"),
        ]
    ])


def speech_style_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🎩 Формальный", callback_data="style_formal")],
        [InlineKeyboardButton(text="😊 Как с другом", callback_data="style_friend")],
        [InlineKeyboardButton(text="⚡ Краткий", callback_data="style_short")],
    ])


def subscription_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(
            text="💳 Купить подписку — 299₽ / месяц",
            callback_data="buy_subscription"
        )
    ]])


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


def is_math_request(text: str) -> bool:
    keywords = ["вычисли", "посчитай", "сколько будет", "реши уравнение", "решить уравнение",
                "найди корни", "интеграл", "производная", "логарифм", "sin", "cos", "sqrt",
                "факториал", "calculate", "solve", "integral", "derivative", "=/", "^2", "^3"]
    return any(kw in text.lower() for kw in keywords)


async def wolfram_calculate(query: str) -> str | None:
    """Точное вычисление через WolframAlpha."""
    if not WOLFRAM_API_KEY:
        return None
    try:
        # Переводим запрос на английский через Claude
        translation_response = await anthropic.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=200,
            messages=[{
                "role": "user",
                "content": f"Переведи этот математический запрос на английский язык для WolframAlpha. Верни ТОЛЬКО перевод без пояснений: {query}"
            }]
        )
        english_query = translation_response.content[0].text.strip()
        logger.info(f"WolframAlpha запрос: {english_query}")

        encoded = urllib.parse.quote(english_query)
        url = f"http://api.wolframalpha.com/v1/result?appid={WOLFRAM_API_KEY}&i={encoded}&units=metric"
        async with aiohttp.ClientSession() as s:
            async with s.get(url, timeout=aiohttp.ClientTimeout(total=15)) as resp:
                if resp.status == 200:
                    result = await resp.text()
                    return result
                else:
                    logger.warning(f"WolframAlpha статус: {resp.status}")
    except Exception as e:
        logger.error(f"WolframAlpha ошибка: {e}")
    return None


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
    history = load_history(user_id)
    content_str = content if isinstance(content, str) else "[фото]"
    history.append({"role": "user", "content": content})

    response = await anthropic.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=2048,
        system=get_system_prompt(user_id),
        messages=history
    )

    reply = response.content[0].text

    # Сохраняем в БД (только текстовые сообщения)
    save_message(user_id, "user", content_str)
    save_message(user_id, "assistant", reply)

    return reply


# ─── Хэндлеры ─────────────────────────────────────────────────────────────────
@dp.message(CommandStart())
async def cmd_start(message: Message):
    user_id = message.from_user.id
    username = message.from_user.username or message.from_user.first_name

    # Проверяем реферальную ссылку
    referred_by = None
    args = message.text.split()
    if len(args) > 1:
        try:
            referred_by = int(args[1])
            if referred_by == user_id:
                referred_by = None  # нельзя пригласить самого себя
        except ValueError:
            pass

    reward = create_user(user_id, username, referred_by)

    # Активируем пробный период для новых пользователей
    trial_activated = activate_trial(user_id)

    # Уведомляем реферера о новом приглашённом
    if referred_by:
        ref_count = get_referral_count(referred_by)
        remaining = 5 - (ref_count % 5)
        try:
            if reward:
                await bot.send_message(
                    referred_by,
                    f"🎉 Ты пригласил уже {ref_count} друзей и получил бесплатный месяц подписки!\n"
                    f"Продолжай приглашать — каждые 5 друзей = 1 месяц бесплатно!"
                )
            else:
                await bot.send_message(
                    referred_by,
                    f"👥 По твоей ссылке зарегистрировался новый пользователь!\n"
                    f"Приглашено: {ref_count} из 5 для бесплатного месяца. Осталось: {remaining}!"
                )
        except Exception:
            pass

    sub_until = get_subscription_until(user_id)
    if trial_activated:
        sub_text = f"🎁 Пробный период активирован на {TRIAL_DAYS} дня — безлимитный доступ!"
    elif sub_until:
        sub_text = f"✅ Подписка активна до {sub_until}"
    else:
        sub_text = f"🆓 Бесплатно: {FREE_MESSAGES_PER_DAY} сообщений/день"

    await message.answer(
        f"👋 Привет, {username}! Я AI-ассистент NeuroBot.\n\n"
        f"{sub_text}\n\n"
        "🔧 Я умею:\n"
        "• 💬 Отвечать на вопросы\n"
        "• 🖼 Распознавать картинки и решать задачи с фото\n"
        "• 🎨 Генерировать изображения по описанию\n"
        "• 📊 Создавать Excel таблицы\n"
        "• 📋 Создавать презентации PowerPoint\n\n"
        "📢 Следи за обновлениями: @NeuroBot_Info\n\n"
        "👇 Используй меню ниже или просто напиши мне!",
        reply_markup=bottom_menu()
    )
    await message.answer("Открываю меню:", reply_markup=main_keyboard())


@dp.message(Command("referral"))
async def cmd_referral(message: Message):
    user_id = message.from_user.id
    ref_count = get_referral_count(user_id)
    remaining = 5 - (ref_count % 5)
    bot_info = await bot.get_me()
    ref_link = f"https://t.me/{bot_info.username}?start={user_id}"

    await message.answer(
        f"👥 Реферальная программа\n\n"
        f"Приглашай друзей и получай бесплатные месяцы!\n\n"
        f"🎁 Каждые 5 приглашённых друзей = 1 месяц бесплатно\n\n"
        f"📊 Твоя статистика:\n"
        f"• Приглашено друзей: {ref_count}\n"
        f"• До следующего бесплатного месяца: {remaining}\n\n"
        f"🔗 Твоя ссылка:\n`{ref_link}`\n\n"
        f"Поделись ссылкой с друзьями — когда они зарегистрируются, ты получишь зачёт!",
        parse_mode="Markdown"
    )


# ─── Фоновая задача: уведомления об истечении подписки ───────────────────────
async def check_expiring_subscriptions():
    """Каждые 12 часов проверяем истекающие подписки."""
    while True:
        try:
            reset_notified_expired()
            expiring = get_users_expiring_soon()
            for user_id, username, sub_until in expiring:
                until_date = datetime.fromisoformat(sub_until).strftime("%d.%m.%Y")
                try:
                    await bot.send_message(
                        user_id,
                        f"⚠️ Твоя подписка истекает {until_date}!\n\n"
                        f"Продли подписку чтобы не потерять безлимитный доступ.",
                        reply_markup=subscription_keyboard()
                    )
                    mark_notified(user_id)
                    logger.info(f"Уведомление отправлено пользователю {user_id}")
                except Exception as e:
                    logger.error(f"Ошибка отправки уведомления {user_id}: {e}")
        except Exception as e:
            logger.error(f"Ошибка проверки подписок: {e}")
        await asyncio.sleep(43200)  # 12 часов


@dp.message(Command("clear"))
async def cmd_clear(message: Message):
    clear_history(message.from_user.id)
    await message.answer("🗑️ История очищена!")


# ─── Обработчики кнопок меню ──────────────────────────────────────────────────
@dp.callback_query(F.data == "menu_subscribe")
async def menu_subscribe(callback):
    await callback.answer()
    sub_until = get_subscription_until(callback.from_user.id)
    extra = f"\nТвоя подписка будет продлена от {sub_until}." if sub_until else ""
    await callback.message.answer(
        f"💳 Подписка на NeuroBot\n\n"
        f"• Безлимитные сообщения\n"
        f"• История диалогов сохраняется\n"
        f"• Генерация картинок без ограничений\n"
        f"• Решение задач с фото\n"
        f"• Все функции без ограничений\n\n"
        f"💰 Цена: 299₽ / месяц{extra}",
        reply_markup=subscription_keyboard()
    )


@dp.callback_query(F.data == "menu_status")
async def menu_status(callback):
    await callback.answer()
    user_id = callback.from_user.id
    sub_until = get_subscription_until(user_id)
    if sub_until:
        await callback.message.answer(f"✅ Подписка активна до {sub_until}\n\nБез ограничений на сообщения!")
    else:
        can, remaining = can_send_message(user_id)
        await callback.message.answer(
            f"🆓 Бесплатный план\n"
            f"Осталось сообщений сегодня: {remaining}/{FREE_MESSAGES_PER_DAY}\n\n"
            f"Купи подписку для безлимитного доступа!",
            reply_markup=subscription_keyboard()
        )


@dp.callback_query(F.data == "menu_referral")
async def menu_referral(callback):
    await callback.answer()
    user_id = callback.from_user.id
    ref_count = get_referral_count(user_id)
    remaining = 5 - (ref_count % 5)
    bot_info = await bot.get_me()
    ref_link = f"https://t.me/{bot_info.username}?start={user_id}"
    await callback.message.answer(
        f"👥 Реферальная программа\n\n"
        f"🎁 Каждые 5 приглашённых друзей = 1 месяц бесплатно\n\n"
        f"📊 Твоя статистика:\n"
        f"• Приглашено друзей: {ref_count}\n"
        f"• До следующего бесплатного месяца: {remaining}\n\n"
        f"🔗 Твоя ссылка:\n`{ref_link}`",
        parse_mode="Markdown"
    )


@dp.callback_query(F.data == "menu_clear")
async def menu_clear(callback):
    await callback.answer()
    clear_history(callback.from_user.id)
    await callback.message.answer("🗑️ История очищена!")


@dp.callback_query(F.data == "menu_help")
async def menu_help(callback):
    await callback.answer()
    await callback.message.answer(
        "🤖 Примеры запросов:\n\n"
        "🎨 Генерация картинки:\n"
        "«Нарисуй закат над морем»\n\n"
        "📊 Excel таблица:\n"
        "«Создай таблицу с расходами за месяц»\n\n"
        "📋 Презентация:\n"
        "«Создай презентацию о космосе на 5 слайдов»\n\n"
        "🖼 Анализ картинки:\n"
        "Просто отправь фото с подписью или без!\n\n"
        "📐 Математика и задачи:\n"
        "• Простые вычисления — точно ✅\n"
        "• Уравнения и формулы — точно ✅\n"
        "• Задачи с фото — хорошо, но проверяй сложные ⚠️\n"
        "• Задачи с чертежами — старается, но может ошибиться ⚠️\n\n"
        "💡 Если ответ неверный — напиши «неправильно» и пришли задачу заново!"
    )


@dp.callback_query(F.data == "menu_speech")
async def menu_speech(callback):
    await callback.answer()
    user_id = callback.from_user.id
    current_style = get_user_style(user_id)
    current_name = STYLE_NAMES.get(current_style, "😊 Как с другом")
    await callback.message.answer(
        f"🎙️ Настройки речи\n\n"
        f"Текущий стиль: {current_name}\n\n"
        f"Выбери как бот будет общаться с тобой:",
        reply_markup=speech_style_keyboard()
    )


@dp.callback_query(F.data.startswith("style_"))
async def set_style(callback):
    await callback.answer()
    style = callback.data.replace("style_", "")
    if style not in STYLE_PROMPTS:
        return
    set_user_style(callback.from_user.id, style)
    style_name = STYLE_NAMES.get(style, style)
    await callback.message.answer(
        f"✅ Стиль общения изменён на {style_name}!\n\n"
        f"Теперь я буду общаться с тобой в этом стиле."
    )


@dp.message(Command("status"))
async def cmd_status(message: Message):
    user_id = message.from_user.id
    sub_until = get_subscription_until(user_id)

    if sub_until:
        await message.answer(f"✅ Подписка активна до {sub_until}\n\nБез ограничений на сообщения!")
    else:
        can, remaining = can_send_message(user_id)
        await message.answer(
            f"🆓 Бесплатный план\n"
            f"Осталось сообщений сегодня: {remaining}/{FREE_MESSAGES_PER_DAY}\n\n"
            f"Купи подписку для безлимитного доступа!",
            reply_markup=subscription_keyboard()
        )


@dp.message(Command("subscribe"))
async def cmd_subscribe(message: Message):
    sub_until = get_subscription_until(message.from_user.id)
    extra = f"\nТвоя подписка будет продлена до следующего месяца от {sub_until}." if sub_until else ""

    await message.answer(
        f"💳 Подписка на NeuroBot\n\n"
        f"• Безлимитные сообщения\n"
        f"• История диалогов сохраняется\n"
        f"• Генерация картинок без ограничений\n"
        f"• Решение задач с фото\n"
        f"• Все функции без ограничений\n\n"
        f"💰 Цена: 299₽ / месяц{extra}",
        reply_markup=subscription_keyboard()
    )


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
        "Просто отправь фото с подписью или без!\n\n"
        "📐 Математика и задачи:\n"
        "• Простые вычисления — точно ✅\n"
        "• Уравнения и формулы — точно ✅\n"
        "• Задачи с фото — хорошо, но проверяй сложные ⚠️\n"
        "• Задачи с чертежами — старается, но может ошибиться ⚠️\n\n"
        "💡 Если ответ неверный — напиши «неправильно» и пришли задачу заново!"
    )


@dp.message(Command("admin"))
async def cmd_admin(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("❌ У тебя нет доступа к этой команде.")
        return

    conn = sqlite3.connect("bot.db")
    c = conn.cursor()

    # Всего пользователей
    c.execute("SELECT COUNT(*) FROM users")
    total_users = c.fetchone()[0]

    # Активных сегодня
    today = datetime.now().date().isoformat()
    c.execute("SELECT COUNT(*) FROM users WHERE last_message_date = ?", (today,))
    active_today = c.fetchone()[0]

    # Платящих подписчиков
    now = datetime.now().isoformat()
    c.execute("SELECT COUNT(*) FROM users WHERE subscription_until > ?", (now,))
    subscribers = c.fetchone()[0]

    # Новых пользователей за последние 7 дней
    week_ago = (datetime.now() - timedelta(days=7)).isoformat()
    c.execute("SELECT COUNT(*) FROM users WHERE created_at > ?", (week_ago,))
    new_week = c.fetchone()[0]

    # Топ 5 активных пользователей
    c.execute("""
        SELECT username, messages_today FROM users
        ORDER BY messages_today DESC LIMIT 5
    """)
    top_users = c.fetchall()

    # Подписчики — список
    c.execute("""
        SELECT username, subscription_until FROM users
        WHERE subscription_until > ?
        ORDER BY subscription_until DESC
    """, (now,))
    sub_list = c.fetchall()

    conn.close()

    top_text = "\n".join([f"  • {u[0] or 'unknown'} — {u[1]} сообщ." for u in top_users]) or "нет данных"
    sub_text = "\n".join([
        f"  • {u[0] or 'unknown'} до {datetime.fromisoformat(u[1]).strftime('%d.%m.%Y')}"
        for u in sub_list
    ]) or "нет подписчиков"

    await message.answer(
        f"📊 Админ-панель NeuroBot\n"
        f"{'─' * 30}\n\n"
        f"👥 Всего пользователей: {total_users}\n"
        f"🟢 Активных сегодня: {active_today}\n"
        f"🆕 Новых за 7 дней: {new_week}\n"
        f"💎 Подписчиков: {subscribers}\n\n"
        f"🏆 Топ активных сегодня:\n{top_text}\n\n"
        f"💳 Подписчики:\n{sub_text}"
    )


# ─── Оплата через Telegram Stars ──────────────────────────────────────────────
@dp.callback_query(F.data == "buy_subscription")
async def process_buy(callback):
    await callback.answer()
    await bot.send_invoice(
        chat_id=callback.from_user.id,
        title="Подписка на AI-ассистента",
        description=f"Безлимитный доступ на {SUBSCRIPTION_DAYS} дней. История сообщений сохраняется.",
        payload="subscription_1month",
        currency="XTR",  # Telegram Stars
        prices=[LabeledPrice(label="Подписка 1 месяц", amount=SUBSCRIPTION_PRICE_STARS)],
    )


@dp.pre_checkout_query()
async def pre_checkout(query: PreCheckoutQuery):
    await query.answer(ok=True)


@dp.message(F.successful_payment)
async def successful_payment(message: Message):
    user_id = message.from_user.id
    activate_subscription(user_id)
    until = get_subscription_until(user_id)
    username = message.from_user.username or message.from_user.first_name
    await message.answer(
        f"🎉 Оплата прошла успешно!\n\n"
        f"✅ Подписка активна до {until}\n\n"
        f"Теперь у тебя безлимитный доступ ко всем функциям!"
    )
    # Уведомление админу
    for admin_id in ADMIN_IDS:
        try:
            await bot.send_message(
                admin_id,
                f"💰 Новая оплата!\n\n"
                f"👤 Пользователь: @{username}\n"
                f"📅 Подписка до: {until}"
            )
        except Exception:
            pass


# ─── Проверка лимитов ──────────────────────────────────────────────────────────
async def check_limits(message: Message) -> bool:
    """Возвращает True если пользователь может отправить сообщение"""
    if message.from_user.id in ADMIN_IDS:
        return True  # админ без лимитов
    user_id = message.from_user.id
    create_user(user_id, message.from_user.username or message.from_user.first_name)

    # Проверяем ежедневный бонус
    if not is_subscribed(user_id):
        bonus_given = claim_daily_bonus(user_id)
        if bonus_given:
            await message.answer(
                f"🎁 Ежедневный бонус! +{DAILY_BONUS_MESSAGES} сообщения сегодня за активность!"
            )

    can, remaining = can_send_message(user_id)

    if not can:
        await message.answer(
            f"⛔ Ты израсходовал лимит бесплатных сообщений на сегодня ({FREE_MESSAGES_PER_DAY} шт.)\n\n"
            f"Лимит обновится завтра, или купи подписку для безлимитного доступа!",
            reply_markup=subscription_keyboard()
        )
        return False

    if not is_subscribed(user_id) and remaining <= 3:
        await message.answer(
            f"⚠️ Осталось {remaining} бесплатных сообщений на сегодня.",
            reply_markup=subscription_keyboard()
        )

    if not is_subscribed(user_id):
        increment_message_count(user_id)

    return True


# ─── Обработчики сообщений ────────────────────────────────────────────────────
@dp.message(F.photo)
async def handle_photo(message: Message):
    if not await check_limits(message):
        return

    await bot.send_chat_action(message.chat.id, "typing")
    try:
        photo = message.photo[-1]
        file = await bot.get_file(photo.file_id)
        buf = io.BytesIO()
        await bot.download_file(file.file_path, buf)
        photo_bytes = buf.getvalue()

        photo_b64 = base64.standard_b64encode(photo_bytes).decode()
        caption = message.caption or "Если на фото математическая или геометрическая задача — реши её строго пошагово, записывая все формулы и промежуточные вычисления. Проверь ответ в конце. Если это не задача — опиши что на изображении."

        content = [
            {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": photo_b64}},
            {"type": "text", "text": caption}
        ]

        reply = await ask_claude(message.from_user.id, content)
        await message.answer(reply)

    except Exception as e:
        logger.error(f"Ошибка обработки фото: {e}")
        await message.answer("❌ Не удалось обработать изображение. Попробуй ещё раз.")


@dp.message(F.voice)
async def handle_voice(message: Message):
    if not await check_limits(message):
        return

    await bot.send_chat_action(message.chat.id, "typing")
    try:
        # Скачиваем голосовое сообщение
        file = await bot.get_file(message.voice.file_id)
        buf = io.BytesIO()
        await bot.download_file(file.file_path, buf)
        voice_bytes = buf.getvalue()
        voice_b64 = base64.standard_b64encode(voice_bytes).decode()

        # Транскрибируем через Claude
        transcribe_response = await anthropic.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=500,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Это голосовое сообщение в формате ogg. Транскрибируй его текст. Верни ТОЛЬКО текст без пояснений."
                    }
                ]
            }]
        )
        # Поскольку Claude не слышит аудио напрямую, просим пользователя написать текст
        await message.answer(
            "🎤 Голосовые сообщения пока в разработке!\n\n"
            "Напиши свой вопрос текстом — отвечу быстро 😊"
        )

    except Exception as e:
        logger.error(f"Ошибка обработки голосового: {e}")
        await message.answer("❌ Не удалось обработать голосовое. Напиши текстом!")


@dp.message(Command("broadcast"))
async def cmd_broadcast(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        await message.answer("❌ У тебя нет доступа к этой команде.")
        return

    text = message.text.replace("/broadcast", "").strip()
    if not text:
        await message.answer(
            "📢 Использование: /broadcast текст сообщения\n\n"
            "Пример: /broadcast 🎉 Добавили новую функцию!"
        )
        return

    await message.answer("📤 Начинаю рассылку...")

    conn = sqlite3.connect("bot.db")
    c = conn.cursor()
    c.execute("SELECT user_id FROM users")
    users = c.fetchall()
    conn.close()

    sent = 0
    failed = 0
    for (user_id,) in users:
        try:
            await bot.send_message(user_id, text)
            sent += 1
            await asyncio.sleep(0.05)  # небольшая задержка чтобы не превысить лимиты
        except Exception:
            failed += 1

    await message.answer(
        f"✅ Рассылка завершена!\n\n"
        f"📨 Отправлено: {sent}\n"
        f"❌ Не доставлено: {failed}"
    )


@dp.message(F.text)
async def handle_text(message: Message):
    if not await check_limits(message):
        return

    user_id = message.from_user.id
    text = message.text

    await bot.send_chat_action(message.chat.id, "typing")

    try:
        # Обработка кнопки меню
        if text == "📋 Меню":
            await message.answer("Открываю меню:", reply_markup=main_keyboard())
            return

        # Определяем что пользователь недоволен ответом
        retry_keywords = [
            "неправильно", "неверно", "ошибка", "ошибся", "не так",
            "перерешай", "попробуй снова", "попробуй ещё", "реши заново",
            "wrong", "incorrect", "try again", "redo", "mistake"
        ]
        if any(kw in text.lower() for kw in retry_keywords):
            clear_history(user_id)
            await message.answer(
                "🔄 Понял, давай попробуем заново с чистого листа!\n\n"
                "Пришли задачу ещё раз — решу пошагово и аккуратно."
            )
            return
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

        # Пробуем WolframAlpha для точных вычислений
        if is_math_request(text) and WOLFRAM_API_KEY:
            wolfram_result = await wolfram_calculate(text)
            if wolfram_result:
                # Claude объясняет ответ WolframAlpha
                enhanced_text = (
                    f"{text}\n\n"
                    f"[Точный ответ от WolframAlpha: {wolfram_result}]\n"
                    f"Объясни решение пошагово, используя этот точный ответ."
                )
                reply = await ask_claude(user_id, enhanced_text)
                await message.answer(f"🔢 *Точный ответ:* `{wolfram_result}`\n\n{reply}", parse_mode="Markdown")
                return

        reply = await ask_claude(user_id, text)
        await message.answer(reply)

    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await message.answer("❌ Произошла ошибка. Попробуй ещё раз.")


# ─── Запуск ───────────────────────────────────────────────────────────────────
async def main():
    init_db()
    threading.Thread(target=run_health_server, daemon=True).start()
    logger.info("Бот запускается...")
    asyncio.create_task(check_expiring_subscriptions())
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
