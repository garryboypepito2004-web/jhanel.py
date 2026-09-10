import random
import re

WORD_PUZZLES = [
    {"word": "CEMENT", "hint": "Binding material used in construction.", "scramble": "MCTEEN"},
    {"word": "STEEL", "hint": "Strong metal used for reinforcement.", "scramble": "EETLS"},
    {"word": "TILE", "hint": "Finish for walls and floors.", "scramble": "LITE"},
    {"word": "BRICK", "hint": "Masonry unit used for walls.", "scramble": "CBRKI"},
    {"word": "SAND", "hint": "Fine aggregate for mixing mortar.", "scramble": "DSAN"},
    {"word": "ROOF", "hint": "Top protective cover of the structure.", "scramble": "OORF"},
    {"word": "BEAM", "hint": "Horizontal structural support.", "scramble": "ABEM"},
    {"word": "PLUMB", "hint": "Level and align pipework.", "scramble": "BPULM"},
]


def _normalize(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


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
