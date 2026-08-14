---
phase: 3
slug: convention
status: planned
nyquist_compliant: true
wave_0_complete: false
created: 2026-07-24
updated: 2026-07-24
---

# Phase 3 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Python stdlib `unittest` (фреймворков в репо нет; pytest сознательно не вводим) |
| **Config file** | none — Wave 1 / план 03-01 Task 1 создаёт `tools/tests/` (это и есть Wave 0-инфраструктура фазы) |
| **Quick run command** | `py -3 -m unittest discover -s tools/tests -q` |
| **Full suite command** | `py -3 -m unittest discover -s tools/tests -q && py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json && py -3 tools/check_convention.py "templates/НоваяКнопка.pushbutton" --strict && py -3 -m py_compile "MM_LAB.extension/lib/revit_compat.py" "templates/НоваяКнопка.pushbutton/script.py"` |
| **Estimated runtime** | ~10 seconds (quick), ~30 seconds (full) |

Примечание: до завершения плана 03-04 полная команда невыполнима целиком (baseline/шаблон ещё не созданы) — на волнах 1–2 прогонять только unittest-часть с шаблонами имён (`-p "test_check_convention*.py"`, `-p "test_revit_compat*.py"`), чтобы параллельные планы не ловили чужой RED.

---

## Sampling Rate

- **After every task commit:** `py -3 -m unittest discover -s tools/tests -p "<шаблон своего плана>" -q`
- **After every plan wave:** `py -3 -m unittest discover -s tools/tests -q` (+ с волны 3 — full suite)
- **Before `/gsd-verify-work`:** full suite green + manual UAT чек-лист (compat в Revit, /mm-adopt-script на IFC_Двери)
- **Max feedback latency:** 30 seconds

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Threat Ref | Secure Behavior | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|------------|-----------------|-----------|-------------------|-------------|--------|
| 03-01 T1 (RED: фикстуры+тесты) | 03-01 | 1 | CONV-CHECK, CONV-REG | T-03-01 | тесты фиксируют ast-only анализ | unit (RED) | `py -3 -m unittest discover -s tools/tests -p "test_check_convention*.py" -q` (ожидаемо ≠0) | создаёт | ⬜ pending |
| 03-01 T2 (GREEN: чекер MM000–MM007, MM013) | 03-01 | 1 | CONV-CHECK, CONV-REG | T-03-01, T-03-02, T-03-03 | нет import/exec/eval; пути нормализованы | unit | `py -3 -m unittest discover -s tools/tests -p "test_check_convention*.py" -q` | ✅ после T1 | ⬜ pending |
| 03-02 T1 (revit_compat.py) | 03-02 | 1 | CONV-STD | T-03-05, T-03-06 | compat не открывает транзакций; fail-fast D-03 | smoke | `py -3 -m py_compile "MM_LAB.extension/lib/revit_compat.py"` | ✅ | ⬜ pending |
| 03-02 T2 (контрактный тест) | 03-02 | 1 | CONV-STD | — | — | unit (ast-контракт) | `py -3 -m unittest discover -s tools/tests -p "test_revit_compat*.py" -q` | создаёт | ⬜ pending |
| 03-03 T1 (RED: AST-правила MM008–MM012, MM014) | 03-03 | 2 | CONV-CHECK | T-03-01 | — | unit (RED) | `py -3 -m unittest discover -s tools/tests -p "test_check_convention*.py" -q` (ожидаемо ≠0) | ✅ | ⬜ pending |
| 03-03 T2 (GREEN: MM008–MM012, MM014) | 03-03 | 2 | CONV-CHECK | T-03-01, T-03-08 | first-party список из стемов, без исполнения | unit | `py -3 -m unittest discover -s tools/tests -p "test_check_convention*.py" -q` | ✅ | ⬜ pending |
| 03-04 T1 (шаблон кнопки) | 03-04 | 3 | CONV-STD | T-03-09, T-03-10 | Transaction+RollBack; require_supported_version | integration | `py -3 -m py_compile "templates/НоваяКнопка.pushbutton/script.py" && py -3 tools/check_convention.py "templates/НоваяКнопка.pushbutton" --strict` | ✅ | ⬜ pending |
| 03-04 T2 (baseline + tab fix + full suite) | 03-04 | 3 | CONV-CHECK, CONV-REG | T-03-11, T-03-12 | орфан удаляется только при отсутствии папки | integration | full suite command | ✅ | ⬜ pending |
| 03-05 T1 (AGENTS.md) | 03-05 | 3 | CONV-STD | T-03-13, T-03-15 | git-запреты в тексте; без userEmail/currentDate | doc-assert | py-однострочник из плана (маркеры разделов) | ✅ | ⬜ pending |
| 03-05 T2 (указатели) | 03-05 | 3 | CONV-STD | T-03-14 | GSD-блок сохранён; graphify/Obsidian перенесены | doc-assert | py-однострочник из плана | ✅ | ⬜ pending |
| 03-06 T1 (mm-adopt-script) | 03-06 | 4 | CONV-ADAPT, CONV-REG, CONV-GSD | T-03-16, T-03-17, T-03-18 | ревью-гейт до регистрации; не исполнять чужой код | doc-assert | py-однострочник из плана (маркеры шагов) | ✅ | ⬜ pending |
| 03-06 T2 (mm-check/new-button/new-compat) | 03-06 | 4 | CONV-CHECK, CONV-STD | T-03-18 | — | doc-assert | py-однострочник из плана | ✅ | ⬜ pending |
| 03-06 T3 (mm-save-session/update-repo/doctor) | 03-06 | 4 | CONV-STD | T-03-18, T-03-19, T-03-20 | ff-only; пофайловый стейджинг; doctor read-only | doc-assert | py-однострочник из плана | ✅ | ⬜ pending |
| 03-07 T1 (21 адаптер) | 03-07 | 5 | CONV-ADAPT, CONV-CHECK | T-03-21, T-03-22, T-03-23 | без shell-вставок; только .claude/commands | file-assert | py-однострочник из плана (21 файл) | ✅ | ⬜ pending |
| 03-07 T2 (каталожный тест + финал) | 03-07 | 5 | CONV-ADAPT, CONV-CHECK | T-03-21, T-03-22 | test_no_shell_injection | unit + full suite | full suite command | создаёт | ⬜ pending |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

Wave 0 фазы = план 03-01 Task 1 (первая волна, до любой реализации чекера):

- [ ] `tools/tests/test_check_convention.py` — покрывает CONV-CHECK/CONV-REG (тест на каждое правило)
- [ ] `tools/tests/fixtures/repo_ok/…` и `fixtures/repo_bad/…` — эталонные деревья pushbutton (bad: BOM, нет шапки/README/bundle.yaml, орфаны layout; AST-дефекты, включая legacy-бутстрап EXTENSION_ROOT для MM014, добавляет план 03-03)
- [ ] `tools/convention_baseline.json` — генерируется чекером из фактического аудита (план 03-04 Task 2, `--write-baseline`)
- [ ] Framework install — не требуется (stdlib `unittest`)

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| Ревью-гейт адаптации стороннего скрипта | CONV-ADAPT | D-08: обязательное одобрение человеком (diff → approve) | После плана 03-07 прогнать `/mm-adopt-script` на IFC_Двери в Claude Code; проверить: чекер отработал, панель спрошена (D-10), diff показан, без «да» регистрации нет; после «да» — bundle.yaml + quick task |
| compat-хелперы в реальном Revit | CONV-STD (D-01..D-04) | Revit API доступен только в среде Revit | UAT-чек-лист из 03-02-PLAN §verification: кнопка-пример на Revit 2020/2022/2024 + fail-fast на прочих |
| Quick task при приёмке | CONV-GSD | Агентская команда, исполняется интерактивно | После приёмки проверить `.planning/quick/<id>/` (PLAN+SUMMARY) и строку в таблице STATE.md |
| Кнопка из шаблона видна на панели | CONV-STD, CONV-REG | pyRevit парсит расширение на старте | Скопировать шаблон в панель, зарегистрировать в layout, pyRevit Reload, убедиться, что кнопка появилась и запускается |

---

## Validation Sign-Off

- [x] All tasks have `<automated>` verify or Wave 0 dependencies
- [x] Sampling continuity: no 3 consecutive tasks without automated verify
- [x] Wave 0 covers all MISSING references (план 03-01 Task 1 — первая задача первой волны)
- [x] No watch-mode flags
- [x] Feedback latency < 30s
- [x] `nyquist_compliant: true` set in frontmatter

**Approval:** planned 2026-07-24 (планировщик фазы 3)
