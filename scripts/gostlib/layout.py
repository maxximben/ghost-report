"""Расчёт вёрстки через Microsoft Word (COM): номера страниц для оглавления и границы таблиц."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROBE_PS1 = os.path.join(os.path.dirname(HERE), "word_probe.ps1")


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    if shutil.which("powershell") is None:
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe"):
            return True
    except OSError:
        return False


def probe(docx: str, pdf: str | None = None, paragraphs: bool = False, timeout: int = 240) -> dict:
    """Запускает word_probe.ps1. Бросает RuntimeError при сбое."""
    docx = os.path.abspath(docx)
    fd, out = tempfile.mkstemp(suffix=".json")
    os.close(fd)
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", PROBE_PS1, "-Docx", docx, "-Out", out]
    if pdf:
        cmd += ["-Pdf", os.path.abspath(pdf)]
    if paragraphs:
        cmd += ["-Paragraphs"]
    try:
        cp = subprocess.run(cmd, capture_output=True, timeout=timeout)
        if cp.returncode != 0:
            raw = cp.stderr or cp.stdout
            try:
                msg = raw.decode("utf-8")
            except UnicodeDecodeError:
                msg = raw.decode("cp866", "replace")
            raise RuntimeError("Word не смог обработать документ: " + msg[-600:])
        with open(out, encoding="utf-8-sig") as f:
            info = json.load(f)
        # PowerShell 5.1 «схлопывает» массивы из одного элемента
        for key in ("tables", "headings"):
            if isinstance(info.get(key), dict):
                info[key] = [info[key]]
            info.setdefault(key, [])
        for t in info["tables"]:
            if isinstance(t.get("rows"), int):
                t["rows"] = [t["rows"]]
        return info
    except subprocess.TimeoutExpired:
        raise RuntimeError("Word не ответил за отведённое время")
    finally:
        try:
            os.remove(out)
        except OSError:
            pass


def table_breaks_from_probe(info: dict) -> dict[str, set[int]]:
    """По номерам страниц строк находит строки, с которых таблица продолжается на новой странице.

    Возвращает {tid: {индекс_строки_данных, ...}} в логической нумерации исходной таблицы (без шапки).
    """
    breaks: dict[str, set[int]] = {}
    for t in info.get("tables", []):
        title = t.get("title") or ""
        if "|" not in title:
            continue
        tid, start = title.rsplit("|", 1)
        try:
            start = int(start)
        except ValueError:
            continue
        pages = t.get("rows", [])
        if isinstance(pages, int):
            pages = [pages]
        # pages[0] – шапка; pages[k] – k-я строка данных части (логический индекс start + k - 1)
        for k in range(2, len(pages)):
            if pages[k] != pages[k - 1]:
                breaks.setdefault(tid, set()).add(start + k - 1)
        # шапка осталась одна внизу страницы – Word не оторвёт её (keep_with_next), отдельный случай не нужен
    return breaks
