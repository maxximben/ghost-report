"""Собирает report.md из report.template.md: подставляет таблицы из results.json и листинг affine.py.

    python make_report.py
"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
fmt = lambda x: f"{x:,}".replace(",", " ")          # 12 065 – с неразрывным пробелом


def pct(x):
    return f"{x:.1f}".replace(".", ",")


def main():
    r = json.load(open(os.path.join(HERE, "results.json"), encoding="utf-8"))
    A = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"
    rows = {}

    half = (len(A) + 1) // 2
    rows["LETTER_ROWS"] = "\n".join(
        "| " + " | ".join([A[i], str(i), f"{r['freq'][A[i]]:.2f}".replace(".", ",")] +
                         ([A[i + half], str(i + half), f"{r['freq'][A[i + half]]:.2f}".replace(".", ",")]
                          if i + half < len(A) else ["", "", ""])) + " |"
        for i in range(half))

    ex = r["example"]
    ex_rows = []
    for ch, c in zip(ex["plain"], ex["cipher"]):
        x = A.index(ch)
        ex_rows.append(f"| {ch} | {x} | {ex['a'] * x + ex['b']} | {(ex['a'] * x + ex['b']) % 33} | {c} |")
    rows["EXAMPLE_ROWS"] = "\n".join(ex_rows)

    rows["KEY_ROWS"] = "\n".join(f"| {k['m']} | {k['phi']} | {k['keys']} |" for k in r["keyspace"])
    rows["SUCCESS_ROWS"] = "\n".join(f"| {s['L']} | {s['trials']} | {s['ok']} | {pct(s['rate'])} |" for s in r["success"])
    rows["TIME_ROWS"] = "\n".join(f"| {t['m']} | {fmt(t['keys'])} | {f'{t['seconds']:.3f}'.replace('.', ',')} |" for t in r["times"])
    rows["UNIT_ROWS"] = "\n".join(f"| {u['a']} | {u['inv']} | {u['a'] * u['inv'] % 33} |" for u in r["units"])
    rows["RAW_ROWS"] = "\n".join(f"| {s['L']} | {s['ok']} | {s['trials'] - s['ok']} | {s['trials']} |" for s in r["success"])
    import affine
    vectors = [("привет", 5, 7), ("алфавит", 2, 1), ("шифр", 10, 3), ("ключ", 13, 20), ("криптография", 7, 3)]
    rows["VECTOR_ROWS"] = "\n".join(f"| {w} | {a} | {b} | {affine.encrypt(w, a, b)} |" for w, a, b in vectors)
    rows["AFFINE"] = open(os.path.join(HERE, "affine.py"), encoding="utf-8").read().rstrip("\n")

    text = open(os.path.join(HERE, "report.template.md"), encoding="utf-8").read()
    for key, val in rows.items():
        assert f"@@{key}@@" in text, key
        text = text.replace(f"@@{key}@@", val)
    assert "@@" not in text
    with open(os.path.join(HERE, "report.md"), "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    print("report.md готов,", len(text.split()), "слов")


if __name__ == "__main__":
    main()
