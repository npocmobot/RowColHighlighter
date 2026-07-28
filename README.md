# RowColHighlighter — надстройка Excel

Подсветка строки и столбца активной ячейки для Excel 2021 (Windows).

## Возможности

- Автоматическая подсветка строки и столбца при выделении ячейки
- Восстановление исходной заливки при переходе на другую ячейку
- Снятие подсветки перед сохранением файла (жёлтый цвет не сохраняется)
- Кнопка включения/выключения на панели быстрого доступа
- Запоминание состояния (вкл/выкл) между запусками Excel
- Восстановление стандартной сетки при выключении

## Установка

1. Скачайте `RowColHighlighter.xlam`
2. **Файл** → **Параметры** → **Центр управления безопасностью** → **Параметры макросов**
   - Включить все макросы
   - Доверять доступ к объектной модели проектов VBA
3. **Файл** → **Параметры** → **Надстройки** → **Перейти** → **Обзор**
4. Выберите скачанный файл → ОК
5. Перезапустите Excel

## Использование

- Просто выделите ячейку — строка и столбец подсветятся
- Кнопка **«Подсветка ВКЛ/ВЫКЛ»** на панели быстрого доступа
- При выключении сетка автоматически восстанавливается

## Изменение цвета подсветки

1. `Alt+F11`
2. Найти `HIGHLIGHT_COLOR` в модуле `AppEventClass`
3. Заменить `&HB4FFFF` на нужный цвет

## Требования

- Excel 2021 для Windows
- Включённые макросы
- Доверие к объектной модели VBA

## Лицензия

Бесплатно для личного и коммерческого использования.

# RowColHighlighter — Excel Add-in

Highlights the row and column of the active cell in Excel 2021 (Windows).

## Features

- Automatic row and column highlighting on cell selection
- Restores original cell colors when moving to another cell
- Removes highlighting before saving (yellow color is not saved)
- Toggle button on the Quick Access Toolbar
- Remembers state (on/off) between Excel launches
- Restores gridlines when turned off

## Installation

1. Download `RowColHighlighter.xlam`
2. **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings**
   - Enable all macros
   - Trust access to the VBA project object model
3. **File** → **Options** → **Add-ins** → **Go** → **Browse**
4. Select the downloaded file → OK
5. Restart Excel

## Usage

- Select any cell — the row and column will be highlighted
- Use the **«Highlight ON/OFF»** button on the Quick Access Toolbar
- Gridlines are automatically restored when turning off

## Changing Highlight Color

1. `Alt+F11`
2. Find `HIGHLIGHT_COLOR` in the `AppEventClass` module
3. Replace `&HB4FFFF` with your color

## Requirements

- Excel 2021 for Windows
- Macros enabled
- Trust access to VBA project object model

## License

Free for personal and commercial use.
