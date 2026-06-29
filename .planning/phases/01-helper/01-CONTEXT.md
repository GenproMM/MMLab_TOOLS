# Phase 1: Дедупликация общих helper-функций ИОС - Context

**Gathered:** 2026-06-09
**Status:** Ready for planning
**Source:** Obsidian architecture decision + code scan

<domain>
## Phase Boundary

Фаза покрывает только 5 ИОС-кнопок:
- Доп расход 0
- Доп расход 1
- Конфузор-Диффузор
- Приточный по классификации
- Сброс потерь

Цель: убрать локальные копии повторяющихся helpers и использовать единый shared-модуль верхнего уровня.

</domain>

<decisions>
## Implementation Decisions

### Locked decisions
- Shared helper-модуль размещается в MM LAB.extension/lib
- script.py кнопок остаются тонкими: orchestration + UI конкретной кнопки
- Логика Transaction не меняется относительно текущей реализации
- Silent except не добавляются; существующее поведение ошибок не ухудшается

### the agent's Discretion
- Имена вспомогательных shared-модулей
- Порядок миграции и минимально инвазивная техника правок

</decisions>

<canonical_refs>
## Canonical References

### Planning
- .planning/ROADMAP.md — цель и критерии фазы
- .planning/REQUIREMENTS.md — IOS-01..IOS-04, SAFE-01

### Architecture Knowledge
- MMLabs_OBSIDIAN/knowledge/patterns/Кнопки pyRevit должны быть тонкими и выносить общую логику в модули.md
- MMLabs_OBSIDIAN/knowledge/decisions/Общие функции ИОС-кнопок выносятся в верхнеуровневый модуль.md

</canonical_refs>

<specifics>
## Specific Ideas

Приоритетно выносятся функции, найденные как массовые дубликаты:
- to_text, normalize_text, element_id_value
- get_parameter*, collect_elements, is_writable, nearly_equal
- get_connector_manager/get_hvac_connectors/get_system_classification
- set_additional_flow_value

</specifics>

<deferred>
## Deferred Ideas

- Расширение shared-подхода на кнопки Архитектура и Координация

</deferred>

---

*Phase: 01-helper*
*Context gathered: 2026-06-09*
