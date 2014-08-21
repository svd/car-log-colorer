''' Macros to colorize AccessPort logs
''' * v 1.4
''' + Changed color style for AFR: using 3-color scale formatting
''' + Changed colors for LTFT
''' + Added colors for STFT
'''
''' * v 1.3
''' + Added color scale for BAT
'''
''' * v 1.2
''' + Added color bar for WGDC
'''
''' * v1.1
''' + Added column selection by name
''' + Highlight Actual AFR cells which are bigger than Equiv. Ratio
'''   by more than 2%
'''

Sub FreezeTopRow()
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

Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Function getColumnName(strColName As String) As String
    Dim aCell As Range

    Set aCell = Rows(1).Find(What:=strColName, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)

    If Not aCell Is Nothing Then
        getColumnName = ConvertToLetter(aCell.Column)
    'Else
    '    getColumnName = Nothing
    End If
End Function

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

    Dim colA As String, colB As String
    Dim f As String
    
    Rows("1:1").Select
    Selection.Font.Bold = True
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
    
    ClearAllFormatConditions
    
    ' Accel. Pedal Position
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
    
    ' Throttle Position
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

    ' Actual AFR
    If selectColumn("Actual AFR (*") Then
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 10.5
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.599993896298105
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 14
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 5287936
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 16
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
        End With
    End If
    
    ' Boost
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
    
    ' Boost Air Temp
    'Columns("E:E").Select
    If selectColumn("Boost Air Temp*") Then
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 30
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .Color = 8109667
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 8711167
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 70
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .Color = 7039480
            .TintAndShade = 0
        End With
    End If
    
    ' HPFP Act. Press.
    'Columns("G:G").Select
    colA = getColumnName("HPFP Act. Press. (*")
    colB = getColumnName("HPFP Des. Press. (*")
    ' Hightlight cell where "HPFP Act. Press." < "HPFP Des. Press."
    ' "= H1 < I1"
    f = "= " & colA & "1 < " & colB & "1"
    If selectColumn("HPFP Act. Press. (*") Then
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:=f
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
    End If
    
    ' Knock Retard
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
    
    ' Long Term FT (%)
    'Columns("K:K").Select
    If selectColumn("Long Term FT (%)") Then
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -12
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = -0.249977111117893
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .Color = 5296274
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 12
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = -0.249977111117893
        End With
    End If
    
    ' Short Term FT (%)
    If selectColumn("Short Term FT (%)") Then
        Selection.FormatConditions.AddColorScale ColorScaleType:=3
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(1).Value = -20
        With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 0
        With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
        End With
        Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
            xlConditionValueNumber
        Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 20
        With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.399975585192419
        End With
    End If

    ' Mass Airflow (g/s)
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
    
    ' RPM
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
    
    ' Vehicle Speed
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
        'Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
        'Selection.FormatConditions(1).Direction = xlContext
        'Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        'Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
        
        Selection.FormatConditions(1).BarFillType = xlDataBarFillGradient
        Selection.FormatConditions(1).Direction = xlContext
        Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
        Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderSolid
        Selection.FormatConditions(1).NegativeBarFormat.BorderColorType = _
            xlDataBarColor
        With Selection.FormatConditions(1).BarBorder.Color
            .Color = 2668287
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
    End If
    
    ' WGDC
    'Columns("T:T").Select
    If selectColumn("Wastegate Duty*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 8061142
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
    
    ' Calculated Load
    If selectColumn("Calculated Load*") Then
        Selection.FormatConditions.AddDatabar
        Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1)
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
        End With
        With Selection.FormatConditions(1).BarColor
            .Color = 8061142
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
