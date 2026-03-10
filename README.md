# Telegram Gemini Bot

A Telegram bot powered by **Google Gemini AI** that supports free-form conversation and can generate professional **PPTX presentations** on any topic.

## Features

| Feature | Description |
|---|---|
| 💬 Free chat | Ask anything — Gemini answers |
| 📊 `/presentation <topic>` | Generate and receive a PPTX file |
| 🎨 Styled slides | Dark theme, colour palette, visual RJ-45 crimping diagram |

## Quick start

### 1. Clone & install dependencies

```bash
git clone https://github.com/1mix667/telegram-gemini-bot.git
cd telegram-gemini-bot
pip install -r requirements.txt
```

### 2. Configure environment variables

Create a `.env` file in the project root:

```env
TELEGRAM_TOKEN=your_telegram_bot_token
GEMINI_API_KEY=your_google_gemini_api_key
```

* **TELEGRAM_TOKEN** — obtain from [@BotFather](https://t.me/BotFather)
* **GEMINI_API_KEY** — obtain from [Google AI Studio](https://aistudio.google.com/apikey)

### 3. Run the bot

```bash
python main.py
```

## Bot commands

| Command | Description |
|---|---|
| `/start` | Welcome message |
| `/help` | Command reference |
| `/presentation <topic>` | Generate a PPTX presentation on `<topic>` |

### Presentation example

```
/presentation среда передачи данных
```

The bot will:
1. Ask Gemini to produce a 7-slide outline on the requested topic
2. Build a styled PPTX (dark theme, bullet points, colour accents)
3. Append a dedicated **RJ-45 T568B crimping colour-order** slide
4. Send the `.pptx` file directly in the chat

## RJ-45 (T568B) crimping colour sequence

The presentation always includes a visual diagram of the T568B standard:

| Pin | Colour |
|-----|--------|
| 1 | White-Orange |
| 2 | Orange |
| 3 | White-Green |
| 4 | Blue |
| 5 | White-Blue |
| 6 | Green |
| 7 | White-Brown |
| 8 | Brown |

> T568A swaps pairs 2 and 3 (pins 1, 2 ↔ 3, 6).

## Requirements

* Python 3.11+
* `python-telegram-bot` 21.x
* `google-generativeai` 0.8.x
* `python-pptx` 1.0.x
* `python-dotenv` 1.0.x
