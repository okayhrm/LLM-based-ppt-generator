# LLM-based-ppt-generator
# 🧠 LLM-Based PPT Generator 

A personal-use AI tool that generates beautiful, structured PowerPoint presentations using Large Language Models (LLMs) like GPT-3.5, Mixtral, Mistral, and Claude — all powered via the OpenRouter API.

Built with Python and Streamlit, designed for speed, clarity, and simplicity.

---

## ✨ Features

- 💡 **AI Slide Generation** — generate entire slide decks from a single prompt
- 🎨 **Template Picker** — choose from multiple professionally designed PPT templates
- 🌐 **Live Data Mode** — optionally add real-time insights using DuckDuckGo
- 🖼️ **Thumbnail Previews** — visually select your favorite slide theme
- 📥 **Download as `.pptx`** — export ready-to-use PowerPoint files instantly
- 🔒 **Local and Private** — keep your API key secure via `.env`

---

## 🛠️ Tech Stack

| Tool               | Purpose                        |
|--------------------|--------------------------------|
| `Streamlit`        | UI framework                   |
| `python-pptx`      | PowerPoint slide generation    |
| `OpenRouter`       | LLM API layer                  |
| `duckduckgo-search`| Real-time web snippets         |
| `dotenv`           | Secure API key management      |
| `Pillow`           | Template thumbnail display     |

---

## ⚙️ Setup Instructions

### 1. Clone this repository

```bash
git clone https://github.com/okayhrm/llm-based-ppt-generator.git
cd llm-based-ppt-generator
