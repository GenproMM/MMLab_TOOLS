# -*- coding: utf-8 -*-
"""Каталожный тест согласованности команд MM LAB (план 03-07).

Исполняемая спецификация каталога mm-команд: 8 канонических процедур
в agents/commands/ и 24 тонких адаптера (8 слагов × Claude Code / Gemini CLI /
Kilo Code). Защищает каталог от дрейфа:

    * канонические процедуры существуют и содержат раздел «## Процедура»;
    * каждый адаптер ссылается на СВОЙ канонический файл agents/commands/;
    * Gemini-адаптеры — TOML с ключами description/prompt и плейсхолдером
      {{args}}; имена файлов плоские (mm-<слаг>.toml, без подпапок);
    * ни один адаптер не содержит shell-вставок "!{...}" (T-03-21);
    * в каталогах адаптеров нет mm-файлов вне SLUGS — защита от
      команд-двойников с опечатками в слагах (T-03-22);
    * AGENTS.md перечисляет все 8 команд каталога.

Запуск:
    py -3 -m unittest discover -s tools/tests -p "test_mm_commands*.py" -q
"""

import unittest
from pathlib import Path

# Слаги команд MM LAB — фиксированы решениями D-19/D-20.
SLUGS = [
    "mm-adopt-script",
    "mm-new-button",
    "mm-check",
    "mm-save-session",
    "mm-update-repo",
    "mm-doctor",
    "mm-new-compat",
    "mm-releasemap-download",
]

# tools/tests/ -> tools/ -> корень репозитория
REPO = Path(__file__).resolve().parents[2]

CANONICAL_DIR = REPO / "agents" / "commands"
CLAUDE_DIR = REPO / ".claude" / "commands"
GEMINI_DIR = REPO / ".gemini" / "commands"
KILO_DIR = REPO / ".kilo" / "commands"


def _read(path):
    """Читает файл строго в UTF-8 (конвенция: файлы каталога — UTF-8 без BOM)."""
    return path.read_text(encoding="utf-8")


def _adapter_paths():
    """Все 24 пути адаптеров: пары (слаг, путь) для трёх агентов."""
    for slug in SLUGS:
        yield slug, CLAUDE_DIR / (slug + ".md")
        yield slug, GEMINI_DIR / (slug + ".toml")
        yield slug, KILO_DIR / (slug + ".md")


class TestMmCommandsCatalog(unittest.TestCase):
    """Согласованность каталога: 8 слагов × 3 адаптера + 8 процедур."""

    def test_canonical_procedures_exist(self):
        """У каждого слага есть канонический файл agents/commands/<слаг>.md."""
        for slug in SLUGS:
            path = CANONICAL_DIR / (slug + ".md")
            with self.subTest(slug=slug):
                self.assertTrue(
                    path.is_file(),
                    "нет канонической процедуры: %s" % path,
                )
                text = _read(path)
                self.assertGreaterEqual(
                    len(text.splitlines()),
                    20,
                    "%s: канонический файл подозрительно короткий" % path.name,
                )
                self.assertIn(
                    "## Процедура",
                    text,
                    "%s: нет раздела «## Процедура»" % path.name,
                )

    def test_claude_adapters(self):
        """Адаптеры Claude Code ссылаются на свои канонические файлы."""
        for slug in SLUGS:
            path = CLAUDE_DIR / (slug + ".md")
            with self.subTest(slug=slug):
                self.assertTrue(path.is_file(), "нет адаптера Claude: %s" % path)
                text = _read(path)
                self.assertIn(
                    "agents/commands/%s.md" % slug,
                    text,
                    "%s: нет ссылки на канонический файл" % path.name,
                )

    def test_gemini_adapters(self):
        """Адаптеры Gemini CLI: TOML c description/prompt, ссылкой и {{args}}."""
        for slug in SLUGS:
            path = GEMINI_DIR / (slug + ".toml")
            with self.subTest(slug=slug):
                self.assertTrue(path.is_file(), "нет адаптера Gemini: %s" % path)
                text = _read(path)
                self.assertIn("description =", text, "%s: нет description" % path.name)
                self.assertIn("prompt =", text, "%s: нет prompt" % path.name)
                self.assertIn(
                    "agents/commands/%s.md" % slug,
                    text,
                    "%s: нет ссылки на канонический файл" % path.name,
                )
                self.assertIn("{{args}}", text, "%s: нет плейсхолдера {{args}}" % path.name)

    def test_kilo_adapters(self):
        """Адаптеры Kilo Code ссылаются на свои канонические файлы."""
        for slug in SLUGS:
            path = KILO_DIR / (slug + ".md")
            with self.subTest(slug=slug):
                self.assertTrue(path.is_file(), "нет адаптера Kilo: %s" % path)
                text = _read(path)
                self.assertIn(
                    "agents/commands/%s.md" % slug,
                    text,
                    "%s: нет ссылки на канонический файл" % path.name,
                )

    def test_no_shell_injection_in_adapters(self):
        """Ни один из 24 адаптеров не содержит shell-вставок "!{" (T-03-21)."""
        for slug, path in _adapter_paths():
            with self.subTest(path=str(path.relative_to(REPO))):
                self.assertNotIn(
                    "!{",
                    _read(path),
                    "%s: обнаружена shell-вставка — адаптеры декларативные" % path,
                )

    def test_no_extra_mm_files(self):
        """В каталогах адаптеров нет mm-файлов вне SLUGS (T-03-22).

        Ловит команды-двойники с опечатками в слагах и файлы с неверным
        расширением. Файлы БЕЗ префикса mm- (чужие инструменты) не трогаем.
        """
        expected = {
            CLAUDE_DIR: {slug + ".md" for slug in SLUGS},
            GEMINI_DIR: {slug + ".toml" for slug in SLUGS},
            KILO_DIR: {slug + ".md" for slug in SLUGS},
        }
        for directory, allowed in expected.items():
            for path in sorted(directory.glob("mm-*")):
                with self.subTest(path=str(path.relative_to(REPO))):
                    self.assertIn(
                        path.name,
                        allowed,
                        "лишний mm-файл (опечатка в слаге?): %s" % path,
                    )

    def test_agents_md_lists_all_slugs(self):
        """AGENTS.md (§Команды MM LAB) перечисляет каждый из 8 слагов."""
        text = _read(REPO / "AGENTS.md")
        for slug in SLUGS:
            with self.subTest(slug=slug):
                self.assertIn(slug, text, "AGENTS.md не упоминает команду %s" % slug)


if __name__ == "__main__":
    unittest.main()
