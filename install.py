#!/usr/bin/env python3
"""Устанавливает скилл ghost-report в папку(и) выбранных ИИ-агентов (копированием).

    python install.py --list
    python install.py claude-project agents-project          # в текущую папку проекта
    python install.py claude-user                            # для всех проектов пользователя
    python install.py --dest D:\\somewhere\\skills           # в произвольную папку скиллов
    python install.py claude-user --pip                      # заодно поставить зависимости

Формат скилла (папка с SKILL.md, поля name/description) – общий открытый стандарт Agent Skills,
поэтому одна и та же папка подходит разным агентам; различаются только каталоги, где они её ищут.
Пути ниже взяты из открытой документации/сообществ – если ваш агент ищет скиллы в другом месте,
используйте --dest.
"""
import argparse
import os
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
HOME = os.path.expanduser("~")
NAME = "ghost-report"

PRESETS = {
    "claude-project": (os.path.join(".claude", "skills"), "Claude Code – только для текущего проекта"),
    "claude-user": (os.path.join(HOME, ".claude", "skills"), "Claude Code – для всех проектов пользователя"),
    "agents-project": (os.path.join(".agents", "skills"), "общий каталог .agents (Codex, opencode и др.) – текущий проект"),
    "agents-user": (os.path.join(HOME, ".agents", "skills"), "общий каталог ~/.agents (Codex, opencode и др.) – пользователь"),
    "codex-user": (os.path.join(HOME, ".codex", "skills"), "OpenAI Codex CLI – пользователь (старый путь)"),
    "opencode-project": (os.path.join(".opencode", "skills"), "opencode – текущий проект"),
    "opencode-user": (os.path.join(HOME, ".config", "opencode", "skills"), "opencode – пользователь"),
}
IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc", "*.docx", "*.pdf", "pages", ".git")


def install(dest_root: str):
    target = os.path.join(os.path.abspath(dest_root), NAME)
    if os.path.abspath(target) == HERE:
        print(f"  пропуск: {target} – это сама папка-источник")
        return
    if os.path.exists(target):
        shutil.rmtree(target)
    shutil.copytree(HERE, target, ignore=IGNORE)
    print(f"  ✓ {target}")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("targets", nargs="*", choices=list(PRESETS), metavar="ВАРИАНТ", help="куда ставить (см. --list)")
    ap.add_argument("--dest", action="append", default=[], help="произвольная папка скиллов (можно несколько раз)")
    ap.add_argument("--list", action="store_true", help="показать готовые варианты")
    ap.add_argument("--pip", action="store_true", help="поставить зависимости (pip install -r requirements.txt)")
    a = ap.parse_args()
    if a.list or (not a.targets and not a.dest):
        for k, (p, d) in PRESETS.items():
            print(f"  {k:18} {d}\n{'':20}{os.path.abspath(p)}")
        return
    for t in a.targets:
        install(PRESETS[t][0])
    for d in a.dest:
        install(d)
    if a.pip:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", os.path.join(HERE, "requirements.txt")])
    print("Готово. Перезапустите агента, чтобы он увидел скилл.")


if __name__ == "__main__":
    main()
