---
phase: 3
slug: convention
status: draft
nyquist_compliant: false
wave_0_complete: false
created: 2026-07-24
---

# Phase 3 — Validation Strategy

> Per-phase validation contract for feedback sampling during execution.

---

## Test Infrastructure

| Property | Value |
|----------|-------|
| **Framework** | Python stdlib `unittest` (фреймворков в репо нет; pytest сознательно не вводим) |
| **Config file** | none — Wave 0 создаёт `tools/tests/` |
| **Quick run command** | `py -3 -m unittest discover -s tools/tests -q` |
| **Full suite command** | `py -3 -m unittest discover -s tools/tests -v && py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json` |
| **Estimated runtime** | ~10 seconds (quick), ~30 seconds (full) |

---

## Sampling Rate

- **After every task commit:** Run `py -3 -m unittest discover -s tools/tests -q`
- **After every plan wave:** Run full suite command (unittest + чекер по репо + py_compile)
- **Before `/gsd-verify-work`:** Full suite must be green + manual UAT чек-лист для compat/команд
- **Max feedback latency:** 30 seconds

---

## Per-Task Verification Map

| Task ID | Plan | Wave | Requirement | Threat Ref | Secure Behavior | Test Type | Automated Command | File Exists | Status |
|---------|------|------|-------------|------------|-----------------|-----------|-------------------|-------------|--------|
| (заполняется планировщиком из PLAN.md) | — | — | CONV-CHECK / CONV-STD / CONV-REG / CONV-ADAPT / CONV-GSD | — | — | unit/integration/smoke/manual | см. карту ниже | ❌ W0 | ⬜ pending |

Карта требований → тесты (из RESEARCH.md · Validation Architecture):

| Req | Behavior | Test Type | Automated Command | File Exists? |
|-----|----------|-----------|-------------------|--------------|
| CONV-CHECK | Чекер ловит каждое правило MM001–MM013 на bad-fixture и молчит на good-fixture | unit | `py -3 -m unittest tools.tests.test_check_convention -q` | ❌ Wave 0 |
| CONV-CHECK | Чекер по всему репо с baseline завершает exit 0 | integration | `py -3 tools/check_convention.py --all --baseline tools/convention_baseline.json` | ❌ Wave 0 |
| CONV-STD | templates/ проходит `--strict` на 100% | integration | `py -3 tools/check_convention.py "templates/НоваяКнопка.pushbutton" --strict` | ❌ Wave 0 |
| CONV-STD | revit_compat.py и шаблон синтаксически валидны | smoke | `py -3 -m py_compile "MM LAB.extension/lib/revit_compat.py" "templates/НоваяКнопка.pushbutton/script.py"` | ✅ |
| CONV-REG | MM007: layout-регистрация и орфаны детектируются | unit | входит в test_check_convention (fixtures с panel bundle.yaml) | ❌ Wave 0 |
| CONV-ADAPT | /mm-adopt-script: чекер→diff→approve→регистрация | manual-only | сценарий в Claude Code на IFC_Двери (ревью-гейт D-08 требует человека) | — |
| CONV-GSD | Quick task артефакты создаются (`.planning/quick/<id>/`, STATE-таблица) | manual | проверка файлов после прогона команды | — |
| D-01..D-04 | compat-хелперы на Revit 2020/2022/2024 | manual-only | чек-лист UAT в PLAN (Revit smoke — прецедент фазы 1) | — |

*Status: ⬜ pending · ✅ green · ❌ red · ⚠️ flaky*

---

## Wave 0 Requirements

- [ ] `tools/tests/test_check_convention.py` — покрывает CONV-CHECK/CONV-REG (по одному тесту на правило)
- [ ] `tools/tests/fixtures/good_button/…` и `fixtures/bad_button/…` — эталонные деревья pushbutton (bad: BOM, wildcard-import, lookupParameter, нет README и т.д.)
- [ ] `tools/convention_baseline.json` — сгенерировать из фактического аудита 17 кнопок
- [ ] Framework install — не требуется (stdlib `unittest`)

---

## Manual-Only Verifications

| Behavior | Requirement | Why Manual | Test Instructions |
|----------|-------------|------------|-------------------|
| Ревью-гейт адаптации стороннего скрипта | CONV-ADAPT | D-08: обязательное одобрение человеком (diff → approve) | Прогнать `/mm-adopt-script` на IFC_Двери в Claude Code; проверить diff-гейт, регистрацию в bundle.yaml, quick task |
| compat-хелперы в реальном Revit | D-01..D-04 | Revit API доступен только в среде Revit (2020/2022/2024) | UAT чек-лист из PLAN: запуск кнопки-примера на каждой поддерживаемой версии |
| Quick task при приёмке | CONV-GSD | Агентская команда, исполняется интерактивно | После приёмки проверить `.planning/quick/<id>/` и запись в STATE.md |

---

## Validation Sign-Off

- [ ] All tasks have `<automated>` verify or Wave 0 dependencies
- [ ] Sampling continuity: no 3 consecutive tasks without automated verify
- [ ] Wave 0 covers all MISSING references
- [ ] No watch-mode flags
- [ ] Feedback latency < 30s
- [ ] `nyquist_compliant: true` set in frontmatter

**Approval:** pending
