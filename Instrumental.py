import argparse
import os
import shutil
import subprocess
from datetime import date
from pathlib import Path
from typing import Dict, Any, Optional

#Константы
#список полей ТЗ, будет использоваться
DEFAULT_FIELDS = [
    ("project_name", "Название проекта"),
    ("goal", "Цель проекта (в 1–3 предложениях)"),
    ("scope", "Область применения"),
    ("feature_list", "Ключевые функции (перечислить кратко)"),
    ("ui_requirements", "Требования к интерфейсу"),
    ("deliverables", "Выходные документы"),
    ("deadlines", "Сроки"),
    ("customer", "Заказчик"),
    ("performers", "Исполнители"),
    ("tz_version", "Версия ТЗ (для отслеживания редактирования)"),
]
