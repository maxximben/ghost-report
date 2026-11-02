"""Эксперименты для примера отчёта (assets/example-full): считает все числа и рисует графики.

    python experiments.py          # -> results.json и img/*.png
"""
import json
import os
import random
import time
from math import gcd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt                                   # noqa: E402
from matplotlib.patches import FancyBboxPatch                     # noqa: E402

import affine as af                                               # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
REFS = os.path.join(HERE, "..", "..", "references")
TRAIN = ["rules.md", os.path.join("..", "SKILL.md")]               # по этим текстам строится эталон частот
TEST = ["markdown-format.md", "sources-format.md", "writing-guide.md"]   # из этих берутся фрагменты для атак
SEED = 20261102
plt.rcParams.update({"font.family": "Times New Roman", "font.size": 11, "axes.grid": True, "grid.alpha": 0.3})


def corpus(files):
    text = ""
    for f in files:
        with open(os.path.join(REFS, f), encoding="utf-8") as fh:
            text += fh.read()
    return af.normalize(text)


def main():
    rnd = random.Random(SEED)
    train, test = corpus(TRAIN), corpus(TEST)
    freq = af.frequencies(train)
    ref = {ch: max(v, 1e-5) for ch, v in freq.items()}              # нули заменяем малым числом (логарифм)
    res = {"train_len": len(train), "test_len": len(test), "m": af.M, "phi": af.euler_phi(af.M)}
    res["keys"] = res["phi"] * af.M
    res["freq"] = {ch: round(100 * freq[ch], 2) for ch in af.ALPHABET}
    res["top5"] = sorted(res["freq"].items(), key=lambda kv: -kv[1])[:5]

    # --- Э1. Пространство ключей для разных алфавитов
    rows = []
    for m in (26, 32, 33, 36, 64, 128, 256):
        phi = af.euler_phi(m)
        rows.append({"m": m, "phi": phi, "keys": m * phi})
    res["keyspace"] = rows

    # --- Э2. Обратимые элементы по модулю 33
    res["units"] = [{"a": a, "inv": af.inverse(a, af.M)} for a in range(1, af.M) if gcd(a, af.M) == 1]

    # --- Э3. Пример шифрования
    sample, a, b = "криптография", 7, 3
    res["example"] = {"plain": sample, "a": a, "b": b, "cipher": af.encrypt(sample, a, b),
                      "back": af.decrypt(af.encrypt(sample, a, b), a, b)}

    # --- Э4. Частотная атака: доля верно найденных ключей в зависимости от длины текста
    lengths = (10, 20, 30, 50, 100, 200, 500)
    trials = 300
    units = [x for x in range(1, af.M) if gcd(x, af.M) == 1]
    succ = []
    for L in lengths:
        ok = 0
        for _ in range(trials):
            start = rnd.randrange(0, len(test) - L)
            plain = test[start:start + L]
            a, b = rnd.choice(units), rnd.randrange(af.M)
            cipher = af.encrypt(plain, a, b)
            ba, bb, _ = af.brute_force(cipher, ref)
            ok += (af.decrypt(cipher, ba, bb) == plain)             # ключ определён однозначно, если расшифровка совпала
        succ.append({"L": L, "ok": ok, "trials": trials, "rate": round(100 * ok / trials, 1)})
    res["success"] = succ

    # --- Э5. Атака по двум известным парам «открытый – шифр»
    stats = {"unique": 0, "several": 0, "none": 0}
    total = 0
    for x1 in range(af.M):
        for x2 in range(af.M):
            if x1 == x2:
                continue
            for _ in range(20):
                a, b = rnd.choice(units), rnd.randrange(af.M)
                y1, y2 = (a * x1 + b) % af.M, (a * x2 + b) % af.M
                keys = af.known_plaintext(x1, y1, x2, y2)
                assert (a, b) in keys
                stats["unique" if len(keys) == 1 else "several"] += 1
                total += 1
    res["kpa"] = {**stats, "total": total, "pairs": af.M * (af.M - 1),
                  "gcd_ok": sum(1 for d in range(1, af.M) if gcd(d, af.M) == 1)}
    # пример системы из текста: ключ (a, b) = (7, 3), открытые буквы «к» (11) и «р» (18)
    x1, x2 = af.ALPHABET.index("к"), af.ALPHABET.index("р")
    y1, y2 = (7 * x1 + 3) % af.M, (7 * x2 + 3) % af.M
    res["kpa_example"] = {"x1": x1, "x2": x2, "y1": y1, "y2": y2, "keys": af.known_plaintext(x1, y1, x2, y2)}

    # --- Э6. Время полного перебора ключей
    plain = test[:50]
    times = []
    for m in (26, 33, 64, 128, 256, 512):
        alphabet = "".join(chr(0x4e00 + i) for i in range(m)) if m not in (26, 33) else \
            ("abcdefghijklmnopqrstuvwxyz" if m == 26 else af.ALPHABET)
        text = "".join(alphabet[ord(ch) % m] for ch in plain)
        refm = {ch: 1.0 / m for ch in alphabet}
        phi = af.euler_phi(m)
        t0 = time.perf_counter()
        af.brute_force(text, refm, alphabet)
        dt = time.perf_counter() - t0
        times.append({"m": m, "keys": m * phi, "seconds": round(dt, 3)})
    res["times"] = times

    with open(os.path.join(HERE, "results.json"), "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)

    img = os.path.join(HERE, "img")
    # --- Рис. 1: схема
    fig, ax = plt.subplots(figsize=(7.2, 2.4))
    ax.axis("off")
    boxes = [("Открытый\nтекст", 0.02), ("Нумерация\nбукв 0…32", 0.215), ("y = a·x + b\n(mod 33)", 0.41),
             ("Номера\nв буквы", 0.605), ("Шифртекст", 0.80)]
    for label, x in boxes:
        ax.add_patch(FancyBboxPatch((x, 0.3), 0.17, 0.4, boxstyle="round,pad=0.01", fc="#eef3fb", ec="#33507a"))
        ax.text(x + 0.085, 0.5, label, ha="center", va="center", fontsize=10)
    for i in range(len(boxes) - 1):
        ax.annotate("", xy=(boxes[i + 1][1] - 0.005, 0.5), xytext=(boxes[i][1] + 0.175, 0.5),
                    arrowprops=dict(arrowstyle="->", color="#33507a"))
    ax.text(0.5, 0.88, "Ключ (a, b), НОД(a, 33) = 1", ha="center", fontsize=10, color="#7a3333")
    ax.annotate("", xy=(0.495, 0.72), xytext=(0.495, 0.84), arrowprops=dict(arrowstyle="->", color="#7a3333"))
    ax.set_xlim(0, 1); ax.set_ylim(0, 1)
    fig.savefig(os.path.join(img, "scheme.png"), dpi=150, bbox_inches="tight"); plt.close(fig)

    # --- Рис. 2: частоты букв
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    order = sorted(af.ALPHABET, key=lambda c: -freq[c])
    ax.bar(range(af.M), [100 * freq[c] for c in order], color="#33507a")
    ax.set_xticks(range(af.M)); ax.set_xticklabels(order)
    ax.set_xlabel("Буква (по убыванию частоты)"); ax.set_ylabel("Частота, %")
    fig.savefig(os.path.join(img, "freq.png"), dpi=150, bbox_inches="tight"); plt.close(fig)

    # --- Рис. 3: доля успеха
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    ax.plot([s["L"] for s in succ], [s["rate"] for s in succ], "o-", color="#33507a")
    ax.set_xscale("log"); ax.set_ylim(0, 105)
    ax.set_xlabel("Длина шифртекста, букв"); ax.set_ylabel("Доля верно найденных ключей, %")
    fig.savefig(os.path.join(img, "success.png"), dpi=150, bbox_inches="tight"); plt.close(fig)

    # --- Рис. 4: время перебора
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    ax.loglog([t["keys"] for t in times], [max(t["seconds"], 1e-3) for t in times], "s-", color="#7a3333")
    ax.set_xlabel("Число ключей m·φ(m)"); ax.set_ylabel("Время перебора, с")
    fig.savefig(os.path.join(img, "bruteforce.png"), dpi=150, bbox_inches="tight"); plt.close(fig)
    print("ok")


if __name__ == "__main__":
    main()
