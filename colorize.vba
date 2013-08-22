Sub FreezeTopRow()
'
' FreezeTopRow Macro
' Freeze top row in current worksheet and make it bold

    Rows("1:1").Select
    Selection.Font.Bold = True
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

Sub ClearFormatConditions()
    Do While Selection.FormatConditions.Count > 0
        Selection.FormatConditions(1).Delete
    Loop
End Sub

Sub ClearAllFormatConditions()
    Cells.Select
    ClearFormatConditions
    Cells(1, 1).Select
End Sub

Function selectColumn(strColName As String) As Boolean
    Dim aCell As Range

    Set aCell = Rows(1).Find(What:=strColName, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

    If Not aCell Is Nothing Then
        'MsgBox "Value Found in Cell " & aCell.Address & _
        '" and the Cell Column Number is " & aCell.Column
        Columns(aCell.Column).Select
        selectColumn = True
    Else
        selectColumn = False
    End If
End Function

Sub ColorizeAPlog()
'
' ColorizeAPlog Macro
' Colorize AP log
'
' Keyboard Shortcut: Ctrl+j
'
    Rows("1:1").Select
    Selection.Font.Bold = True
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ClearAllFormatConditions
    
    'Columns("B:B").Select
    If selectColumn("Accel. Pedal Pos*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
    End If
    
    
    'Columns("P:P").Select
    If selectColumn("Throttle Position*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 13012579
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
        End With
        ActiveWindow.SmallScroll Down:=27
    End If

    'Columns("C:C").Select
    If selectColumn("Actual AFR (*") Then
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=14.7"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    End If
    
    'Columns("D:D").Select
    If selectColumn("Boost (*") Then
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueLowestValue
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .ThemeColor = xlThemeColorAccent5
            .TintAndShade = -0.249977111117893
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValueHighestValue
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .Color = 192
            .TintAndShade = 0
        End With
    End If
    
    'Columns("G:G").Select
    If selectColumn("HPFP Act. Press. (*") Then
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:="= G1 < H1"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 49407
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    End If
    
    'Columns("J:J").Select
    If selectColumn("Knock Retard*") Then
        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Font
            .Color = -16751204
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 10284031
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    End If
    
    'Columns("K:K").Select
    If selectColumn("Long Term FT (%)") Then
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -12
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .Color = 192
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 5287936
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 12
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .Color = 255
            .TintAndShade = 0
        End With
    End If
    
    'Columns("M:M").Select
    If selectColumn("Mass Airflow (g/s)*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 15698432
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 15698432
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.BorderColor
            .Color = 255
            .TintAndShade = 0
        End With
    End If
    
    'Columns("N:N").Select
    If selectColumn("RPM (*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 2668287
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
    End If
    
    'Columns("Q:Q").Select
    If selectColumn("Vehicle Speed*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 2668287
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
        Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
        With Selection.FormatConditions(1).AxisColor
            .Color = 0
            .TintAndShade = 0
        End With
        With Selection.FormatConditions(1).NegativeBarFormat.Color
            .Color = 255
            .TintAndShade = 0
        End With
    End If
    
    Cells(1, 1).Select
End Sub
