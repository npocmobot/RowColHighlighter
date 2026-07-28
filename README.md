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
2. Откройте Excel
3. **Файл** → **Параметры** → **Центр управления безопасностью** → **Параметры центра управления безопасностью** → **Параметры макросов**
   - Выберите **«Включить все макросы»**
   - Поставьте галочку **«Доверять доступ к объектной модели проектов VBA»**
   - Нажмите ОК → ОК
4. **Файл** → **Параметры** → **Надстройки**
5. Внизу «Управление:» выберите **«Надстройки Excel»** → нажмите **«Перейти»**
6. Нажмите **«Обзор»**
7. Найдите и выберите скачанный файл `RowColHighlighter.xlam` → ОК
8. Убедитесь, что стоит галочка напротив **RowColHighlighter** → ОК
9. Закройте Excel и откройте заново

## Использование

- Выделите любую ячейку — строка и столбец подсветятся жёлтым цветом
- Для включения/выключения используйте кнопку **«Подсветка ВКЛ/ВЫКЛ»** на панели быстрого доступа
- При выключении сетка автоматически восстанавливается
- Состояние сохраняется между запусками Excel

## Как изменить цвет подсветки

1. Нажмите `Alt + F11` — откроется редактор VBA
2. В левом окне найдите **VBAProject (RowColHighlighter.xlam)**
3. Разверните **Class Modules** → дважды щёлкните **AppEventClass**
4. Найдите строку: `Private Const HIGHLIGHT_COLOR As Long = &HB4FFFF`
5. Замените `&HB4FFFF` на нужный цвет:
   - Светло-голубой: `&HFFFFCC`
   - Светло-зелёный: `&HC6EFCE`
   - Светло-розовый: `&HD8E4F0`
   - Светло-серый: `&HE0E0E0`
6. Нажмите `Ctrl + S`, закройте редактор, перезапустите Excel

## Как изменить цвет восстанавливаемой сетки

1. `Alt + F11` → **Modules** → **Module1**
2. Найдите `RGB(217, 217, 217)` (в двух местах)
3. Замените на нужный цвет, например `RGB(200, 200, 200)`

## Как удалить надстройку

1. **Файл** → **Параметры** → **Надстройки**
2. «Управление:» → **Надстройки Excel** → **Перейти**
3. Снимите галочку с **RowColHighlighter** → ОК

## Как сбросить сохранённое состояние

1. Нажмите `Win + R`
2. Введите `regedit` → Enter
3. Перейдите по пути: `HKEY_CURRENT_USER\Software\RowColHighlighter`
4. Удалите папку `RowColHighlighter`

## Требования

- Excel 2021 для Windows
- Включённые макросы
- Доверие к объектной модели VBA

## Лицензия

Бесплатно для личного и коммерческого использования. Распространяется «как есть» без каких-либо гарантий.

---

## English

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
2. Open Excel
3. **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Macro Settings**
   - Select **«Enable all macros»**
   - Check **«Trust access to the VBA project object model»**
   - Click OK → OK
4. **File** → **Options** → **Add-ins**
5. At the bottom «Manage:» select **«Excel Add-ins»** → click **«Go»**
6. Click **«Browse»**
7. Find and select the downloaded `RowColHighlighter.xlam` file → OK
8. Make sure **RowColHighlighter** is checked → OK
9. Close Excel and reopen

## Usage

- Select any cell — the row and column will be highlighted in yellow
- Use the **«Highlight ON/OFF»** button on the Quick Access Toolbar
- Gridlines are automatically restored when turning off
- State is saved between Excel launches

## License

Free for personal and commercial use. Distributed «as is» without any warranties.
