Attribute VB_Name = "MSR"
Sub Macro_Specification_v011s()
'переменные для макроса ---------------------------------
    For_Sheet_Specification = "Спецификация"
    NameSheetCopy = ActiveSheet.Name
    
'Раздел переменных для языковых вариантов ----------------------------
    'messageGood = "The search ended with the beginning of the table "
    messageGood = "Поиск завершился начало работы таблицы "
    'messageFailed = "The search failed!"
    messageFailed = "Поиск завершился неудачно!"
    'messageEnd = "A report has been created for the client!"
    messageEnd = "Отчет для клиента создан"
    'messageErFix = "fix the field or create a new one"
    messageErFix = "Исправьте или создайте поле"
'---------------------------------------------------------------------
    Position_1 = "Позиция"
    Position_2 = "Наименование и техническая характеристика"
    Position_3 = "Тип, марка, обозначение документа, опросного листа"
    Position_4 = "Код оборудования, изделия, материала"
    Position_5 = "Завод-изготовитель"
    Position_6 = "Единица измерения"
    Position_7 = "Количество"
    Position_8 = "Масса единицы, кг"
    Position_9 = "Примечание"
'переменные для цикла --
    Allrecs = Application.WorksheetFunction.CountA(Sheets(NameSheetCopy).Range("A:A"))
    
'создаем лист для спецификации -------------------------
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(For_Sheet_Specification).Delete
    Application.DisplayAlerts = True
    Worksheets.Add.Name = For_Sheet_Specification
   
    Sheets(For_Sheet_Specification).Activate


'Создаем наименование таблицы ---------------------------
    Range("A2:A2").Value = Position_1
    Range("B2:B2").Value = Position_2
    Range("C2:C2").Value = Position_3
    Range("D2:D2").Value = Position_4
    Range("E2:E2").Value = Position_5
    Range("F2:F2").Value = Position_6
    Range("G2:G2").Value = Position_7
    Range("H2:H2").Value = Position_8
    Range("I2:I2").Value = Position_9

'подгон текста -------------------------------------------
    Range("A2:I2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 39
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Underline = xlUnderlineStyleNone
        .Font.ThemeFont = xlThemeFontMinor
        .ReadingOrder = xlContext
        .HorizontalAlignment = xlCenter
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlThin
    End With
    
    Columns("A:A").ColumnWidth = 8
    Columns("B:B").ColumnWidth = 68.14
    Columns("C:C").ColumnWidth = 24
    Columns("D:D").ColumnWidth = 16.14
    Columns("E:E").ColumnWidth = 14.14
    Columns("F:F").ColumnWidth = 9.14
    Columns("G:G").ColumnWidth = 6.14
    Columns("H:H").ColumnWidth = 8.14
    Columns("I:I").ColumnWidth = 10.68
    
'нумерация столбцов -------------------------------------

    Range("A3:A3").Value = 1
    Range("B3:B3").Value = 2
    Range("C3:C3").Value = 3
    Range("D3:D3").Value = 4
    Range("E3:E3").Value = 5
    Range("F3:F3").Value = 6
    Range("G3:G3").Value = 7
    Range("H3:H3").Value = 8
    Range("I3:I3").Value = 9
    
    Range("A3:I3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
'начало цикла для копирования данных --------------------

    Number_t = "№"
    Сhapter_t_N = "№ разд."
    Chapter_t = "Раздел"
    Manufacture_t = "Произв."
    Model_t = "Модель"
    Name_t = "Наименование"
    Index_t = "Ед. изм."
    Count_t = "Кол-во"
    Price_t = "Цена"
    Note_t = "Примечание"
    Subsection_Name_Index = "Подраздел"
    'определили искомые объекты в таблице
    Dim massiv_table(11) As String
    massiv_table(0) = Number_t
    massiv_table(1) = Сhapter_t_N
    massiv_table(2) = Manufacture_t
    massiv_table(3) = Model_t
    massiv_table(4) = Name_t
    massiv_table(5) = Index_t
    massiv_table(6) = Count_t
    massiv_table(7) = Price_t
    massiv_table(8) = Note_t
    massiv_table(9) = Chapter_t
    massiv_table(10) = Subsection_Name_Index
    
    Line_From_Copy_Begin = 1
    
'всё тот же проклятый цикл для поиска начала таблицы ----
For posLine = 1 To 10
    LineFind = False

    For posCol = 1 To 40
        If Sheets(NameSheetCopy).Cells(posLine, posCol) = Number_t Then
            massiv_table(0) = "done"
            Number_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Сhapter_t_N Then
            massiv_table(1) = "done"
            Сhapter_t_N = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Manufacture_t Then
            massiv_table(2) = "done"
            Manufacture_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Model_t Then
            massiv_table(3) = "done"
            Model_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Name_t Then
            massiv_table(4) = "done"
            Name_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Index_t Then
            massiv_table(5) = "done"
            Index_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Count_t Then
            massiv_table(6) = "done"
            Count_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Price_t Then
            massiv_table(7) = "done"
            Price_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Note_t Then
            massiv_table(8) = "done"
            Note_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Chapter_t Then
            massiv_table(9) = "done"
            Chapter_t = posCol
            LineFind = True
        ElseIf Sheets(NameSheetCopy).Cells(posLine, posCol) = Subsection_Name_Index Then
            massiv_table(10) = "done"
            Subsection_Name_Index = posCol
            LineFind = True
        End If
    
    Next posCol
    
        If LineFind = True Then
            Line_From_Copy_Begin = posLine + 1  'устанавливаем строку начала покирования <-
            Exit For
        End If
    
Next posLine
'--------------------------------------------------------
For CurrecTest = 0 To 10
    If massiv_table(CurrecTest) <> "done" Then
        error_type = massiv_table(CurrecTest)
        MsgBox messageErFix & " - ( " & error_type & " )"   'сообщение об ошибке
        Exit Sub
    End If
Next CurrecTest

'раздел копирования данных ------------------------------

'переменные --
begin_chapter = 4
end_chanter = 4
FindMaxPos = 0
Old_Chapter = 0
New_Chapter = 0
New_Sub_Name_Index = Subsection_Name_Index

'-------------
Allrecs = Allrecs + Line_From_Copy_Begin

For it = Line_From_Copy_Begin To Allrecs
    If FindMaxPos < Sheets(NameSheetCopy).Cells(it, Number_t) And Sheets(NameSheetCopy).Cells(it, Сhapter_t_N) <> Empty Then
        FindMaxPos = Sheets(NameSheetCopy).Cells(it, Number_t)
    End If
Next it
'--------------------------------------------------------

NewPre_Chapter = True

For Copy_Line = 1 To FindMaxPos

'Если раздел новый --------------------------------------
    If New_Chapter <> Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Сhapter_t_N) Then
       NewPre_Chapter = True
        For_Coler_Item = "A" & begin_chapter & ":I" & begin_chapter
        Range(For_Coler_Item).Select
        With Selection
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Font.Bold = True
            .Interior.ThemeColor = xlThemeColorDark2
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 1).Value = "Раздел " & Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Сhapter_t_N)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 1).Font.Size = 10
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 2).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Chapter_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 2).Font.Size = 10

        begin_chapter = begin_chapter + 1
        New_Chapter = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Сhapter_t_N)
    End If
'--------------------------------------------------------

'Если новый подраздел -----------------------------------
    If NewPre_Chapter Or (New_Sub_Name_Index <> Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Subsection_Name_Index) And Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Subsection_Name_Index) <> "") Then
        NewPre_Chapter = False
        For_Coler_Item = "A" & begin_chapter & ":I" & begin_chapter
        Range(For_Coler_Item).Select
        With Selection
            .Font.Italic = True
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
        End With

        Sheets(For_Sheet_Specification).Cells(begin_chapter, 2).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Subsection_Name_Index)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 2).Font.Size = 10
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 2).HorizontalAlignment = xlCenter

        begin_chapter = begin_chapter + 1
        New_Sub_Name_Index = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Subsection_Name_Index)
    End If
'--------------------------------------------------------
'Копируем данные ----------------------------------------

        Sheets(For_Sheet_Specification).Cells(begin_chapter, 1).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Number_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 1).HorizontalAlignment = xlCenter
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 2).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Name_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 4).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Model_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 4).HorizontalAlignment = xlCenter
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 5).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Manufacture_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 5).HorizontalAlignment = xlCenter
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 6).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Index_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 6).HorizontalAlignment = xlCenter
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 7).Value = Sheets(NameSheetCopy).Cells(Line_From_Copy_Begin, Count_t)
        Sheets(For_Sheet_Specification).Cells(begin_chapter, 7).HorizontalAlignment = xlCenter
        
        For_Coler_Item = "A" & begin_chapter & ":I" & begin_chapter
        Range(For_Coler_Item).Select
        With Selection
        .Rows.AutoFit
        .WrapText = True
        .VerticalAlignment = xlCenter
            .Font.Size = 10
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
        End With
        
        begin_chapter = begin_chapter + 1

'--------------------------------------------------------

Line_From_Copy_Begin = Line_From_Copy_Begin + 1
Next Copy_Line
'Конец --------------------------------------------------
End Sub
