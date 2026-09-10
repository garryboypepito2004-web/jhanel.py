import random
import re

WORD_PUZZLES = [
    {"word": "Ethnocentric — Sci-Fi Executive", "hint": "A culturally anchored leader for tomorrow's systems.", "scramble": "Ecrh—cnnEitceo tS iFci xEivue"},
    {"word": "Strategic — Urban Visionary", "hint": "A city-first planner with long-range thinking.", "scramble": "rctSg—i aUbn Vionary"},
    {"word": "Adaptive — Lunar Architect", "hint": "A flexible designer building beyond Earth.", "scramble": "aAetvdi—Lnr uaAcrhtie"},
    {"word": "Operational — Quantum Analyst", "hint": "A systems-minded evaluator solving complex problems.", "scramble": "Oeprlntioa—Qtaum Aansylt"},
    {"word": "Cultural — Future Governor", "hint": "A public leader balancing heritage and innovation.", "scramble": "lCuatru—rFueot Gvnaero"},
    {"word": "Decisive — Data Broker", "hint": "A sharp operator turning information into action.", "scramble": "Dceisiv—Dtat Bkrero"},
    {"word": "Resilient — Civic Engineer", "hint": "A durable builder for communities and infrastructure.", "scramble": "Rielsinet—Civc Egiineer"},
    {"word": "Autonomous — Solar Strategist", "hint": "An independent thinker shaping sustainable futures.", "scramble": "Aotnmsou—Srol Sgtreaist"},
]


def _normalize(value):
    if value is None:
        return ""
    text = str(value).lower()
    text = text.replace("–", "-").replace("—", "-").replace("−", "-")
    text = re.sub(r"[^a-z0-9\s-]", "", text)
    text = re.sub(r"[-\s]+", " ", text)
    return text.strip()


def puzzle_for_level(level, seed=None):
    rng = random.Random(seed if seed is not None else 1)
    index = (int(level) - 1) % len(WORD_PUZZLES)
    puzzle = dict(WORD_PUZZLES[index])
    scramble = list(puzzle["word"])
    rng.shuffle(scramble)
    puzzle["scramble"] = "".join(scramble)
    return puzzle


def check_guess(level, guess, seed=None):
    puzzle = puzzle_for_level(level, seed)
    return _normalize(guess) == _normalize(puzzle["word"])


def math_problem(level, seed=None):
    rng = random.Random(seed if seed is not None else 1)
    base = max(2, int(level))
    a = rng.randint(2, 12 + base)
    b = rng.randint(1, 9 + base)
    operation = rng.choice(["+", "-", "*"])
    if operation == "+":
        answer = a + b
        question = f"{a} + {b} = ?"
    elif operation == "-":
        answer = a - b if a >= b else b - a
        question = f"{max(a, b)} - {min(a, b)} = ?"
    else:
        answer = a * b
        question = f"{a} x {b} = ?"
    return question, answer


def check_math_answer(expected, guess):
    try:
        return float(guess) == float(expected)
    except (TypeError, ValueError):
        return False
