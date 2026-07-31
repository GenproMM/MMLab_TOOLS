# -*- coding: utf-8 -*-
"""
Базовый класс для создания проверок ВОР.
Наследуйтесь от этого класса чтобы создавать новые проверки.
"""

from abc import ABC, abstractmethod
from pyrevit import revit, DB


class BaseValidation(ABC):
    """
    Базовый класс проверки ВОР.
    
    При создании новой проверки:
    1. Унаследуйтесь от BaseValidation
    2. Реализуйте методы: name(), description(), run()
    3. Поместите файл в папку validations/
    """
    
    @abstractmethod
    def name(self):
        """
        Краткое имя проверки (отображается в UI).
        
        Returns:
            str — имя проверки
        """
        pass
    
    @abstractmethod
    def description(self):
        """
        Описание проверки (что именно проверяется).
        
        Returns:
            str — описание
        """
        pass
    
    @abstractmethod
    def run(self, doc, section, project):
        """
        Логика проверки.
        
        Args:
            doc: Revit Document
            section: str — выбранный раздел (АР, КР, ...)
            project: str — выбранный проект
        
        Returns:
            ValidationResult — результат проверки
        """
        pass


class ValidationResult:
    """Результат проверки."""
    
    def __init__(self, check_name, passed, message, element_ids=None):
        """
        Args:
            check_name: str — имя проверки
            passed: bool — пройдена ли проверка
            message: str — сообщение с описанием результата
            element_ids: list[ElementId] — проблемные элементы (опционально)
        """
        self.check_name = check_name
        self.passed = passed
        self.message = message
        self.element_ids = element_ids or []
    
    def __str__(self):
        status = "OK" if self.passed else "FAIL"
        return "[{}] {}: {}".format(status, self.check_name, self.message)


# ============================================================
# ПРИМЕР: Заглушка проверки
# ============================================================

class ExampleValidation(BaseValidation):
    """Пример проверки — заглушка для демонстрации."""
    
    def name(self):
        return "Пример: Проверка заполненности параметра ВОР"
    
    def description(self):
        return "Проверяет что у элементов выбранного раздела заполнен параметр ВОР"
    
    def run(self, doc, section, project):
        # Здесь будет реальная логика
        return ValidationResult(
            check_name=self.name(),
            passed=True,
            message="Все элементы имеют заполненный параметр ВОР (заглушка)"
        )
