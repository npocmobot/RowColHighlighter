Надстройка для подсветки активной ячейки
Полная инструкция: создание, установка и перенос надстройки
________________________________________
Возможности надстройки
•	Автоматическая подсветка строки и столбца активной ячейки жёлтым цветом
•	Восстановление исходных цветов при переходе на другую ячейку
•	Снятие подсветки при сохранении файла (жёлтый цвет не сохраняется)
•	Кнопка включения/выключения на панели быстрого доступа
•	Запоминание состояния (вкл/выкл) между запусками Excel
•	Восстановление стандартной сетки при выключении и сохранении
________________________________________
Часть 1. Создание надстройки
Шаг 1. Настройка безопасности Excel
1.	Откройте Excel
2.	Файл → Параметры → Центр управления безопасностью
3.	Параметры центра управления безопасностью → Параметры макросов
4.	Выберите «Включить все макросы»
5.	Поставьте галочку: «Доверять доступ к объектной модели проектов VBA»
6.	ОК → ОК
7.	Закройте Excel
________________________________________
Шаг 2. Создайте файл надстройки
1.	Откройте Excel (пустая «Книга1»)
2.	Нажмите Alt + F11 — откроется редактор VBA
3.	Слева найдите VBAProject (Книга1)
________________________________________
Шаг 3. Создайте классовый модуль AppEventClass
1.	Правая кнопка на VBAProject (Книга1) → Insert → Class Module
2.	Появится папка Class Modules с файлом Class1
3.	Нажмите F4 (окно Properties)
4.	В строке (Name) сотрите Class1, введите: AppEventClass
5.	Нажмите Enter
6.	Дважды щёлкните по AppEventClass — справа пустое поле
7.	Скопируйте и вставьте этот код:
vba
Option Explicit

Public WithEvents App As Application
Private DictRows As Object
Private DictCols As Object
Private PrevWb As Workbook
Private PrevWs As Worksheet
Private PrevRow As Long
Private PrevCol As Long
Private IsHighlighted As Boolean
Private IsRestoringForSave As Boolean

Private Const HIGHLIGHT_COLOR As Long = &HB4FFFF

Private Sub Class_Initialize()
    Set DictRows = CreateObject("Scripting.Dictionary")
    Set DictCols = CreateObject("Scripting.Dictionary")
    PrevRow = -1
    PrevCol = -1
    IsHighlighted = False
    IsRestoringForSave = False
End Sub

Private Sub Class_Terminate()
    If IsHighlighted Then RestoreHighlighting
    Set DictRows = Nothing
    Set DictCols = Nothing
End Sub

Private Function IsValidSheet(Sh As Object) As Boolean
    On Error Resume Next
    IsValidSheet = (TypeName(Sh) = "Worksheet")
    On Error GoTo 0
End Function

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo EH
    If IsRestoringForSave Then Exit Sub
    If Not IsValidSheet(Sh) Then Exit Sub
    If Target Is Nothing Then Exit Sub
    
    Application.EnableEvents = False
    
    If IsHighlighted Then RestoreHighlighting
    
    Set DictRows = Nothing
    Set DictCols = Nothing
    Set DictRows = CreateObject("Scripting.Dictionary")
    Set DictCols = CreateObject("Scripting.Dictionary")
    
    Set PrevWb = Sh.Parent
    Set PrevWs = Sh
    PrevRow = Target.Row
    PrevCol = Target.Column
    
    ApplyHighlight PrevWs, PrevRow, PrevCol
    IsHighlighted = True
    
    Application.EnableEvents = True
    Exit Sub
    
EH:
    Application.EnableEvents = True
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error GoTo SaveEH
    If Not IsHighlighted Then Exit Sub
    If PrevWb Is Nothing Then Exit Sub
    If Not PrevWb Is Wb Then Exit Sub
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    RestoreHighlighting
    IsHighlighted = False
    IsRestoringForSave = True
    
    RestoreGridlinesForSheet PrevWs
    
    Application.OnTime Now, "RestoreHighlightAfterSave"
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
    
SaveEH:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    On Error Resume Next
    If IsRestoringForSave Then Exit Sub
    If Not IsValidSheet(Wb.ActiveSheet) Then Exit Sub
    
    Application.EnableEvents = False
    If IsHighlighted Then RestoreHighlighting
    
    Set DictRows = Nothing
    Set DictCols = Nothing
    Set DictRows = CreateObject("Scripting.Dictionary")
    Set DictCols = CreateObject("Scripting.Dictionary")
    
    Set PrevWb = Wb
    Set PrevWs = Wb.ActiveSheet
    PrevRow = ActiveCell.Row
    PrevCol = ActiveCell.Column
    
    ApplyHighlight PrevWs, PrevRow, PrevCol
    IsHighlighted = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    If IsRestoringForSave Then Exit Sub
    If Not IsValidSheet(Sh) Then Exit Sub
    
    Application.EnableEvents = False
    If IsHighlighted Then RestoreHighlighting
    
    Set DictRows = Nothing
    Set DictCols = Nothing
    Set DictRows = CreateObject("Scripting.Dictionary")
    Set DictCols = CreateObject("Scripting.Dictionary")
    
    Set PrevWb = Sh.Parent
    Set PrevWs = Sh
    PrevRow = ActiveCell.Row
    PrevCol = ActiveCell.Column
    
    ApplyHighlight PrevWs, PrevRow, PrevCol
    IsHighlighted = True
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

Private Sub RestoreHighlighting()
    Dim v As Variant
    On Error Resume Next
    If Not PrevWs Is Nothing Then
        For Each v In DictRows.Keys
            PrevWs.Range(CStr(v)).Interior.Color = CLng(DictRows(CStr(v)))
        Next
        For Each v In DictCols.Keys
            PrevWs.Range(CStr(v)).Interior.Color = CLng(DictCols(CStr(v)))
        Next
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyHighlight(ByVal ws As Worksheet, ByVal rw As Long, ByVal cl As Long)
    Dim cell As Range
    Dim usedRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    
    On Error GoTo EH
    Set usedRange = ws.UsedRange
    lastRow = usedRange.Row + usedRange.Rows.Count - 1
    lastCol = usedRange.Column + usedRange.Columns.Count - 1
    
    For i = 1 To lastCol
        Set cell = ws.Cells(rw, i)
        If cell.Interior.Color <> HIGHLIGHT_COLOR Then
            DictRows.Add cell.Address, cell.Interior.Color
            cell.Interior.Color = HIGHLIGHT_COLOR
        End If
    Next i
    
    For j = 1 To lastRow
        Set cell = ws.Cells(j, cl)
        If cell.Interior.Color <> HIGHLIGHT_COLOR Then
            DictCols.Add cell.Address, cell.Interior.Color
            cell.Interior.Color = HIGHLIGHT_COLOR
        End If
    Next j
    
    Exit Sub
EH:
End Sub

Public Sub RestoreAfterSave()
    On Error Resume Next
    IsRestoringForSave = False
    If Not PrevWs Is Nothing Then
        If Not IsHighlighted Then
            Application.EnableEvents = False
            ApplyHighlight PrevWs, PrevRow, PrevCol
            IsHighlighted = True
            Application.EnableEvents = True
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub RestoreGridlinesForSheet(ws As Worksheet)
    Dim usedRange As Range
    Dim cell As Range
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim rw As Long, cl As Long
    
    On Error Resume Next
    
    If ws Is Nothing Then Exit Sub
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then Exit Sub
    
    firstRow = Application.Max(1, usedRange.Row - 20)
    lastRow = Application.Min(ws.Rows.Count, usedRange.Row + usedRange.Rows.Count + 19)
    firstCol = Application.Max(1, usedRange.Column - 5)
    lastCol = Application.Min(ws.Columns.Count, usedRange.Column + usedRange.Columns.Count + 4)
    
    For rw = firstRow To lastRow
        For cl = firstCol To lastCol
            Set cell = ws.Cells(rw, cl)
            
            If cell.Borders(xlEdgeLeft).LineStyle = xlNone And _
               cell.Borders(xlEdgeTop).LineStyle = xlNone And _
               cell.Borders(xlEdgeBottom).LineStyle = xlNone And _
               cell.Borders(xlEdgeRight).LineStyle = xlNone And _
               cell.Borders(xlInsideVertical).LineStyle = xlNone And _
               cell.Borders(xlInsideHorizontal).LineStyle = xlNone Then
                
                cell.Borders.LineStyle = xlContinuous
                cell.Borders.Color = RGB(217, 217, 217)
                cell.Borders.Weight = xlThin
            End If
        Next cl
    Next rw
End Sub
________________________________________
Шаг 4. Создайте обычный модуль Module1
1.	Правая кнопка на VBAProject (Книга1) → Insert → Module
2.	Появится папка Modules с файлом Module1
3.	Дважды щёлкните по Module1 — справа пустое поле
4.	Скопируйте и вставьте этот код:
vba
Option Explicit

Private AppHandler As AppEventClass
Private IsEnabled As Boolean

Private Const REG_PATH As String = "HKEY_CURRENT_USER\Software\RowColHighlighter"
Private Const REG_KEY As String = "Enabled"

Public Sub Auto_Open()
    IsEnabled = ReadStateFromRegistry
    
    If IsEnabled Then
        StartCatch
    End If
    
    AddButtonToRibbon
End Sub

Public Sub Auto_Close()
    SaveStateToRegistry IsEnabled
    
    StopCatch
    RemoveButtonFromRibbon
End Sub

Public Sub ToggleHighlight()
    If IsEnabled Then
        StopCatch
        IsEnabled = False
        RestoreGridlines
        SaveStateToRegistry False
        MsgBox "Подсветка ВЫКЛЮЧЕНА. Сетка восстановлена.", vbInformation, "RowColHighlighter"
    Else
        StartCatch
        IsEnabled = True
        SaveStateToRegistry True
        MsgBox "Подсветка ВКЛЮЧЕНА", vbInformation, "RowColHighlighter"
    End If
End Sub

Public Sub StartCatch()
    If AppHandler Is Nothing Then
        Set AppHandler = New AppEventClass
        Set AppHandler.App = Application
    End If
End Sub

Public Sub StopCatch()
    If Not AppHandler Is Nothing Then
        Set AppHandler.App = Nothing
        Set AppHandler = Nothing
    End If
End Sub

Public Sub RestoreHighlightAfterSave()
    If Not AppHandler Is Nothing Then
        AppHandler.RestoreAfterSave
    End If
End Sub

Public Sub RestoreGridlines()
    Dim ws As Worksheet
    
    On Error Resume Next
    For Each ws In ActiveWorkbook.Worksheets
        RestoreGridlinesOnSheet ws
    Next ws
End Sub

Private Sub RestoreGridlinesOnSheet(ws As Worksheet)
    Dim usedRange As Range
    Dim cell As Range
    Dim firstRow As Long, lastRow As Long
    Dim firstCol As Long, lastCol As Long
    Dim rw As Long, cl As Long
    
    On Error Resume Next
    
    Set usedRange = ws.UsedRange
    If usedRange Is Nothing Then Exit Sub
    
    firstRow = Application.Max(1, usedRange.Row - 20)
    lastRow = Application.Min(ws.Rows.Count, usedRange.Row + usedRange.Rows.Count + 19)
    firstCol = Application.Max(1, usedRange.Column - 5)
    lastCol = Application.Min(ws.Columns.Count, usedRange.Column + usedRange.Columns.Count + 4)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    For rw = firstRow To lastRow
        For cl = firstCol To lastCol
            Set cell = ws.Cells(rw, cl)
            
            If cell.Borders(xlEdgeLeft).LineStyle = xlNone And _
               cell.Borders(xlEdgeTop).LineStyle = xlNone And _
               cell.Borders(xlEdgeBottom).LineStyle = xlNone And _
               cell.Borders(xlEdgeRight).LineStyle = xlNone And _
               cell.Borders(xlInsideVertical).LineStyle = xlNone And _
               cell.Borders(xlInsideHorizontal).LineStyle = xlNone Then
                
                cell.Borders.LineStyle = xlContinuous
                cell.Borders.Color = RGB(217, 217, 217)
                cell.Borders.Weight = xlThin
            End If
        Next cl
    Next rw
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub SaveStateToRegistry(enabled As Boolean)
    On Error Resume Next
    CreateObject("WScript.Shell").RegWrite REG_PATH & "\" & REG_KEY, CInt(enabled), "REG_DWORD"
End Sub

Private Function ReadStateFromRegistry() As Boolean
    On Error Resume Next
    Dim value As Integer
    value = CreateObject("WScript.Shell").RegRead(REG_PATH & "\" & REG_KEY)
    If Err.Number <> 0 Then
        ReadStateFromRegistry = True
    Else
        ReadStateFromRegistry = (value = 1)
    End If
End Function

Private Sub AddButtonToRibbon()
    On Error Resume Next
    Application.CommandBars("Quick Access Toolbar").Controls("Подсветка ВКЛ/ВЫКЛ").Delete
    With Application.CommandBars("Quick Access Toolbar").Controls.Add(Type:=1)
        .Caption = "Подсветка ВКЛ/ВЫКЛ"
        .OnAction = "ToggleHighlight"
        .Style = 1
        .FaceId = 159
    End With
End Sub

Private Sub RemoveButtonFromRibbon()
    On Error Resume Next
    Application.CommandBars("Quick Access Toolbar").Controls("Подсветка ВКЛ/ВЫКЛ").Delete
End Sub
________________________________________
Шаг 5. Проверьте структуру
В левом окне редактора VBA должно быть:
text
VBAProject (Книга1)
    Microsoft Excel Objects
        Лист1 (Лист1)
        ЭтаКнига
    Class Modules
        AppEventClass
    Modules
        Module1
________________________________________
Шаг 6. Сохраните как надстройку
1.	Нажмите Ctrl + S
2.	Закройте редактор VBA
3.	Нажмите «Файл» → «Сохранить как»
4.	В поле «Тип файла» выберите: «Надстройка Excel (*.xlam)»
5.	В поле «Имя файла» введите: RowColHighlighter
6.	В адресную строку окна сохранения вставьте: %APPDATA%\Microsoft\AddIns → Enter
7.	Нажмите «Сохранить»
8.	Если спросит про замену — «Да»
9.	Закройте Excel. На вопрос «Сохранить Книгу1?» — «Не сохранять»
________________________________________
Шаг 7. Подключите надстройку
1.	Откройте Excel
2.	Файл → Параметры → Надстройки
3.	Внизу «Управление:» → Надстройки Excel → Перейти
4.	Если в списке есть RowColHighlighter — поставьте галочку
5.	Если нет — Обзор → адресная строка: %APPDATA%\Microsoft\AddIns → Enter → выбрать RowColHighlighter.xlam → ОК
6.	Убедитесь, что галочка стоит → ОК
7.	Закройте Excel и откройте заново
________________________________________
Шаг 8. Проверка
1.	Выделите любую ячейку — строка и столбец подсветятся жёлтым
2.	На верхней панели быстрого доступа появится кнопка «Подсветка ВКЛ/ВЫКЛ»
3.	Нажмите кнопку — подсветка выключится, сетка восстановится
4.	Закройте Excel и откройте снова — подсветка останется выключенной
5.	Нажмите кнопку ещё раз — включится и запомнит состояние
________________________________________
Часть 2. Перенос на другой компьютер
Шаг 1. Найдите файл на старом ПК
1.	Нажмите Win + R
2.	Введите: %APPDATA%\Microsoft\AddIns → Enter
3.	Найдите файл RowColHighlighter.xlam
4.	Скопируйте его на флешку, в облако или отправьте по почте
________________________________________
Шаг 2. Настройте безопасность на новом ПК
1.	Откройте Excel
2.	Файл → Параметры → Центр управления безопасностью
3.	Параметры центра управления безопасностью → Параметры макросов
4.	Выберите «Включить все макросы»
5.	Поставьте галочку: «Доверять доступ к объектной модели проектов VBA»
6.	ОК → ОК
7.	Закройте Excel
________________________________________
Шаг 3. Скопируйте файл в папку AddIns
1.	Нажмите Win + R
2.	Введите: %APPDATA%\Microsoft\AddIns → Enter
3.	Скопируйте файл RowColHighlighter.xlam в эту папку
________________________________________
Шаг 4. Подключите надстройку
1.	Откройте Excel
2.	Файл → Параметры → Надстройки
3.	«Управление:» → Надстройки Excel → Перейти
4.	Обзор
5.	Адресная строка: %APPDATA%\Microsoft\AddIns → Enter
6.	Выберите RowColHighlighter.xlam → ОК
7.	Поставьте галочку → ОК
8.	Закройте Excel и откройте заново
________________________________________
Шаг 5. Если кнопка не появилась
Добавьте вручную:
1.	Стрелочка вниз в конце панели быстрого доступа
2.	«Другие команды...»
3.	«Выбрать команды из:» → Макросы
4.	Выберите ToggleHighlight → Добавить >> → ОК
________________________________________
Памятка
Действие	Как сделать
Включить/выключить подсветку	Кнопка на панели быстрого доступа
Изменить цвет подсветки	Alt+F11 → AppEventClass → строка HIGHLIGHT_COLOR
Изменить цвет сетки	Module1 → RGB(217, 217, 217) (в двух местах)
Восстановить сетку	Нажать кнопку выключения подсветки
Сбросить состояние	Win+R → regedit → удалить HKEY_CURRENT_USER\Software\RowColHighlighter
Удалить надстройку	Параметры → Надстройки → снять галочку
Перенести на другой ПК	Скопировать файл RowColHighlighter.xlam

