# ghost-report – оформление учебных отчётов по правилам и ГОСТ

Скилл для ИИ-агентов: по просьбе «сделай отчёт» агент пишет содержание, а оформление (структура, шрифты, поля,
заголовки, рисунки, таблицы, формулы, код, список источников по ГОСТ Р 7.0.100-2018) выполняет и проверяет
программа. Результат – `.docx`; при наличии Microsoft Word – ещё PDF и PNG страниц для проверки глазами.

## СУПЕР МЕГА ПРОСТОЙ И БЫСТЫЙ СТАРТ

1. заходим в папку /skills вашего агента
2. git clone https://github.com/maxximben/ghost-report
3. если нужно собрать отчёт - вызываем /ghost-report и указываем условие задания и что сделали

## Быстрый старт (без агента)

```
pip install -r requirements.txt
python scripts/build_report.py assets/example/report.md -o out/пример.docx --pdf --png out/pages
python scripts/check_report.py out/пример.docx --work rgr
```

## Подключение к агенту

Формат (`SKILL.md` + скрипты рядом) – общий стандарт Agent Skills, **одна и та же папка работает в разных агентах**.
Логика целиком в Python-скриптах, поэтому от агента нужно только умение читать `SKILL.md` и запускать команды.

```
python install.py --list                      # куда можно поставить
python install.py claude-user                 # Claude Code, для всех проектов  (~/.claude/skills)
python install.py claude-project              # Claude Code, только этот проект (./.claude/skills)
python install.py agents-user                 # общий ~/.agents/skills – его читают Codex, opencode и др.
python install.py agents-project              # общий ./.agents/skills
python install.py --dest D:\мои\скиллы        # любая папка
```

| Агент | Где ищет скиллы (по открытой документации; проверьте для своей версии) |
|---|---|
| Claude Code | `~/.claude/skills`, `./.claude/skills` |
| OpenAI Codex CLI | `~/.agents/skills`, `./.agents/skills` (ранее `~/.codex/skills`) |
| opencode | `./.opencode/skills`, `~/.config/opencode/skills`; также читает `.claude/skills` и `.agents/skills` |
| Cursor, Gemini CLI, Cline и др. | `.agents/skills` или `.claude/skills` (зависит от версии) |

Если агент **не поддерживает скиллы**, вставьте в его файл инструкций (`AGENTS.md`, `CLAUDE.md`, `.cursorrules` и т. п.):

```
Для отчётов (ЛР, ДЗ, РГР, практика) используй инструкцию и скрипты из папки ghost-report:
прочитай ghost-report/SKILL.md и следуй ему. Не оформляй docx вручную –
только через ghost-report/scripts/build_report.py и check_report.py.
```

Различия между агентами – только в каталоге установки, отдельные версии скилла не нужны.
Возможности, специфичные для Claude Code (`allowed-tools`, hooks и т. п.), не используются намеренно ради переносимости.

## Требования

- Python 3.9+, `python-docx`, `lxml`, `Pillow`; `pymupdf` – по желанию (PNG страниц).
- **Windows + Microsoft Word** – для точных номеров страниц в содержании, «Продолжение таблицы N» и PDF.
  Без Word всё остальное работает на любой ОС; номера страниц в содержании обновятся при открытии файла (F9).

## Что внутри

```
SKILL.md                    инструкция для агента
scripts/build_report.py     report.md + sources.json -> report.docx (+ проверка, PDF, PNG)
scripts/check_report.py     проверка любого .docx по правилам
scripts/word_probe.ps1      расчёт вёрстки в Word (COM)
scripts/gostlib/            модули: разбор Markdown, формулы LaTeX->Word, источники ГОСТ, сборка, проверка
references/                 правила, формат report.md, формат источников, руководство по написанию
assets/example/             короткий пример РГР
assets/example-full/        развёрнутый пример РГР (~30 страниц, все элементы)
```

Настройки (шрифты, поля, профили видов работ) – `scripts/gostlib/config.py`.
