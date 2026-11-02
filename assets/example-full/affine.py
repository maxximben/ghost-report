"""Аффинный шифр над русским алфавитом и методы его криптоанализа (учебная реализация)."""
from collections import Counter
from math import gcd, log

ALPHABET = "абвгдеёжзийклмнопрстуфхцчшщъыьэюя"   # 33 буквы, m = 33
M = len(ALPHABET)


def normalize(text, alphabet=ALPHABET):
    """Оставляет только буквы алфавита в нижнем регистре."""
    return "".join(ch for ch in text.lower() if ch in alphabet)


def egcd(a, b):
    """Расширенный алгоритм Евклида: возвращает (g, x, y), где a*x + b*y = g = НОД(a, b)."""
    x0, y0, x1, y1 = 1, 0, 0, 1
    while b:
        q = a // b
        a, b = b, a - q * b
        x0, x1 = x1, x0 - q * x1
        y0, y1 = y1, y0 - q * y1
    return a, x0, y0


def inverse(a, m):
    """Обратный элемент a^(-1) по модулю m; ошибка, если НОД(a, m) != 1."""
    g, x, _ = egcd(a % m, m)
    if g != 1:
        raise ValueError(f"{a} не обратим по модулю {m}")
    return x % m


def euler_phi(n):
    """Функция Эйлера: число чисел от 1 до n, взаимно простых с n."""
    return sum(1 for k in range(1, n + 1) if gcd(k, n) == 1)


def encrypt(text, a, b, alphabet=ALPHABET):
    m = len(alphabet)
    if gcd(a, m) != 1:
        raise ValueError("ключ a должен быть взаимно прост с m")
    return "".join(alphabet[(a * alphabet.index(ch) + b) % m] for ch in text)


def decrypt(text, a, b, alphabet=ALPHABET):
    m, a_inv = len(alphabet), inverse(a, len(alphabet))
    return "".join(alphabet[(a_inv * (alphabet.index(ch) - b)) % m] for ch in text)


def frequencies(text, alphabet=ALPHABET):
    """Относительные частоты букв текста."""
    counts, n = Counter(text), max(len(text), 1)
    return {ch: counts[ch] / n for ch in alphabet}


def log_score(text, reference):
    """Логарифм правдоподобия текста при эталонном распределении букв (чем больше, тем лучше)."""
    return sum(log(reference[ch]) for ch in text)


def brute_force(cipher, reference, alphabet=ALPHABET):
    """Полный перебор ключей; лучший ключ выбирается по правдоподобию. Возвращает (a, b, оценка)."""
    m = len(alphabet)
    best = None
    for a in range(1, m):
        if gcd(a, m) != 1:
            continue
        for b in range(m):
            score = log_score(decrypt(cipher, a, b, alphabet), reference)
            if best is None or score > best[2]:
                best = (a, b, score)
    return best


def known_plaintext(x1, y1, x2, y2, m=M):
    """Все ключи (a, b), удовлетворяющие системе a*x_i + b = y_i (mod m), i = 1, 2."""
    d, e = (x1 - x2) % m, (y1 - y2) % m
    g = gcd(d, m)
    if e % g:
        return []                                  # система несовместна
    m_g = m // g
    a0 = (e // g) * inverse(d // g, m_g) % m_g if m_g > 1 else 0
    keys = []
    for k in range(g):                             # g решений по модулю m
        a = a0 + k * m_g
        if gcd(a, m) == 1:
            keys.append((a, (y1 - a * x1) % m))
    return keys
