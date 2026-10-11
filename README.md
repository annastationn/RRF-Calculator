# RRF Calculator — prototype

Пробная desktop-версия для macOS и Windows.

## Что умеет

- выбрать или перетащить `.xlsx`;
- проверить наличие `Group Name` и `Volume`;
- показать ошибки по неизвестным классам;
- показать предупреждения для строк без `C<number>` и для `C19+` / `C14 ?`;
- посчитать `C`, `H`, `Formula`, `RRF`, `RRF*V`, `w, %`;
- создать лист `Class summary`;
- сгруппировать данные по `summary_group` из статического словаря;
- добавить heatmap в Excel;
- редактировать словарь в отдельной вкладке;
- сохранить пользовательский словарь и восстановить значения по умолчанию.

## Запуск на macOS

```bash
cd /путь/к/rrf_calculator_prototype
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## Запуск на Windows PowerShell

```powershell
cd C:\path\to\rrf_calculator_prototype
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

## Сборка отдельного приложения

Собирать нужно отдельно на каждой ОС.

### macOS

```bash
pyinstaller --noconfirm --windowed --name "RRF Calculator" \
  --add-data "default_classes.json:." app.py
```

### Windows

```powershell
pyinstaller --noconfirm --windowed --name "RRF Calculator" `
  --add-data "default_classes.json;." app.py
```

Готовая сборка появится в `dist/`.

## Файлы проекта

- `app.py` — интерфейс;
- `core.py` — расчётный модуль;
- `default_classes.json` — встроенный исходный словарь;
- `requirements.txt` — зависимости.
