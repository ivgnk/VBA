Attribute VB_Name = "NewMacros"
 Option Compare Text
 Public Type RGB_Color
         Red_color As Byte
       Green_color As Byte
        Blue_color As Byte
End Type

Dim skvNum As String
 Const Rzlt = vbLf + vbLf
 Const ScaleLogging = 640
 Const LeftLL = 120, NumTickY = 5
 Const TickSize = 10, LabelWidth = 90, LabelYWidth = 65
 Const IsLogDepthScale = False
 Const MaxHeightDraw = 350
 Const XPosSecCurve = MaxHeightDraw + 15

 Dim NewRGBColor(1 To 3) As RGB_Color
 Dim PatternArray(1 To 12) As Long
 Dim MaxResistivity, MinResistivity As Double
 Dim XPos, YPos, WidthShX, HeightY, MaxXX, MinXX, MaxYY, MinYY, _
 TopY_col, BotY_col, TheMax_Depth, MashFact, CarotageShift As Single

 Public LogData As Collection ' Создаем объект Collection. Для
 'хранения насчитанной "ломаной" каротажа
 Public MyStrColl As Collection ' Создаем объект Collection. Для
 'хранения образцовых имен и цветов
 Public NameMoshnStrColl As Collection ' Создаем объект Collection. Для
 'данных геологиии

Sub DrawAxes()
Dim СurrX
Dim i, NumTickX As Integer
Dim StepofResist, StepOfDepth, NewWidthShX, TempSi, TempSi2, _
TempSi3 As Single
Dim Man As Double
Dim Por As Integer
Dim NuumTick As Integer
Dim TickArr(1 To 4) As Byte
 '================ начало процедуры
 '================ горизонтальная ось - сопротивление
    Selection.WholeStory
    Selection.Cut
    Call DrawColumn
    Call DrawLegend
    XPos = MinXX
    YPos = MinYY - 10
    NewWidthShX = WidthShX + 30
    СurrX = XPos + NewWidthShX
    ActiveDocument.Shapes.AddLine(XPos, YPos, СurrX, YPos).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.Line.EndArrowheadLength = msoArrowheadLengthMedium
    Selection.ShapeRange.Line.EndArrowheadWidth = msoArrowheadWidthMedium
    
'    Selection.ShapeRange.Line.Weight = 0.75
'    Selection.ShapeRange.Line.DashStyle = msoLineSolid
'    Selection.ShapeRange.Line.Style = msoLineSingle
'    Selection.ShapeRange.Line.Transparency = 0#
'    Selection.ShapeRange.Line.Visible = msoTrue
    ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        СurrX - 35, YPos + 4, 72.8, 30#).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.TextFrame.TextRange.Select
    '----------- Надпись на оси сопротивлений
    Selection.Font.Name = "Symbol"
    Selection.Font.Size = 12
    Selection.TypeText text:="r"
    Selection.Font.Name = "Times New Roman"
    Selection.TypeText text:=", Ом·м"
    '------------ Метки на оси сопротивлений
    
    NumTickX = 5
    
    For i = 1 To NumTickX
    '------ Метки на оси
     ActiveDocument.Shapes.AddLine(XPos, YPos - (TickSize \ 2), XPos, YPos + (TickSize \ 2)).Select
    '------ Линии разграфки сопротивлений
    ActiveDocument.Shapes.AddLine(XPos, YPos + (TickSize \ 2), XPos, BotY_col).Select
    '+ (YPos + (TickSize \ 2))
    Selection.ShapeRange.Line.DashStyle = msoLineRoundDot

      '------- "Текстовый ящик"
     ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        XPos, YPos - 20, LabelWidth, 36#).Select
     Selection.ShapeRange.Fill.Visible = msoFalse
     Selection.ShapeRange.Line.Visible = msoFalse
     Selection.ShapeRange.TextFrame.TextRange.Select
     '----------- Надпись на ящике
     Selection.Font.Name = "Times New Roman"
     Selection.Font.Size = 12
     Selection.TypeText text:=Str(Int(Exp(Log(MinResistivity) + (i - 1) * _
     (Log(MaxResistivity) - Log(MinResistivity)) / NumTickX) * 10) / 10)
     XPos = XPos + (WidthShX \ NumTickX)
    Next i
    
    '================ вертикальная ось - глубины
    
    YPos = BotY_col
    StepOfDepth = (TopY_col - BotY_col) / NumTickY
    ActiveDocument.Shapes.AddLine(LeftLL - 10, TopY_col, LeftLL - 10, BotY_col + 25).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.Line.EndArrowheadLength = msoArrowheadLengthMedium
    Selection.ShapeRange.Line.EndArrowheadWidth = msoArrowheadWidthMedium
    '----------- Надпись на оси глубин
    ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
     LeftLL - LabelYWidth + 20, BotY_col + 15, LabelYWidth \ 2, 20#).Select
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.TextFrame.TextRange.Select
    Selection.Font.Name = "Times New Roman"
    Selection.TypeText text:="H, м"
    
If Not IsLogDepthScale Then
    TempSi = TheMax_Depth
    For i = 1 To NumTickY + 1
    '------ Метки на оси
    ActiveDocument.Shapes.AddLine(LeftLL - 20, YPos, LeftLL - 10, YPos).Select
'    '------ Линии разграфки
'    'ActiveDocument.Shapes.AddLine(XPos, YPos + (TickSize \ 2), XPos, 461.5).Select
'    'Selection.ShapeRange.Line.DashStyle = msoLineRoundDot
    '------- "Текстовый ящик"
     ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        LeftLL - LabelYWidth - 9, YPos - 25, LabelYWidth, 36#).Select
     Selection.ShapeRange.Fill.Visible = msoFalse
     Selection.ShapeRange.Line.Visible = msoFalse
     Selection.ShapeRange.TextFrame.TextRange.Select
    '----------- Надпись на ящике
     Selection.Font.Name = "Times New Roman"
     Selection.Font.Size = 12
     If i = 1 Then
      Selection.TypeText text:=Str(TheMax_Depth)
     ElseIf (i = NumTickY + 1) Then
      Selection.TypeText text:="0"
     Else
      Selection.TypeText text:=Str(Int(TempSi * 10) / 10)
     End If
     Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      TempSi = TempSi - TheMax_Depth / NumTickY
      YPos = YPos + StepOfDepth
    Next i
Else   '-------If Not IsLogDepthScale Then
 Call Div_Man_por(TheMax_Depth, Man, Por)
 For i = 1 To 4
  TickArr(i) = i * 2 - 1
 Next i
 '----- все равно глубины > 1
' NuumTick = Por * 4
' If Man >= 7 Then
'  NuumTick = NuumTick + 4
' ElseIf Man >= 5 Then
'  NuumTick = NuumTick + 3
' ElseIf Man >= 3 Then
'  NuumTick = NuumTick + 2
' ElseIf Man >= 1 Then
'  NuumTick = NuumTick + 1
' End If
 For i = 0 To Por
  For j = 1 To 4
   TempSi3 = Exp(i * Log(10)) * TickArr(j)
   If TempSi3 > TheMax_Depth Then Exit For
   TempSi2 = 110 + Log(TempSi3) * MashFact
   ActiveDocument.Shapes.AddLine(LeftLL - 20, TempSi2, LeftLL - 10, TempSi2).Select
   ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        LeftLL - LabelYWidth - 19, TempSi2 - 12, LabelYWidth, 30#).Select
   Selection.ShapeRange.Fill.Visible = msoFalse
   Selection.ShapeRange.Line.Visible = msoFalse
   Selection.ShapeRange.TextFrame.TextRange.Select
   '----------- Надпись на ящике
   Selection.Font.Name = "Times New Roman"
   If j = 1 Then
    Selection.Font.Size = 12
   Else
    Selection.Font.Size = 10
   End If
   Selection.TypeText text:=Str(TempSi3)
   Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
  Next j
 Next i
 
End If '-------If Not IsLogDepthScale Then
    '=========== Подпись рисунка
    Selection.ShapeRange.Select
    ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, 113.6, _
        541.2, 653.2, 35.5).Select
    Selection.ShapeRange.TextFrame.TextRange.Select
    Selection.ShapeRange.Fill.Visible = msoFalse
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.Font.Size = 14
    Selection.TypeText text:= _
    "Геологическая колонка и данные электрического каротажа по скважине " + skvNum
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    
    
    '=========== Группировка
    ActiveDocument.Shapes.SelectAll
    Selection.ShapeRange.Group.Select
    MsgBox "Has Ended", , "Has Ended"
End Sub

Sub Geology()
 Dim MyString
 Dim TempS As String
 Dim gData, w As Класс1, Ind As Byte
 Dim MinTolsh, TempSi, Tolsh, Resist, CurrH, Sum As Double
 '===========
 MinTolsh = 120000
 Application.ScreenUpdating = False
 Call MyPageSetup
 
 skvNum = 1091
 TempS = Trim(Str(skvNum))
 TempS = "d:\genik\co\txt\d-kvit\" + TempS + "\" + TempS + "-g.txw"
 'MsgBox TempS
 'End
 Open TempS For Input As #1
 ' Открывает файл для чтения.
 Set NameMoshnStrColl = Nothing
 If NameMoshnStrColl Is Nothing Then Set NameMoshnStrColl = New Collection
 Ind = 0: MaxResistivity = -10
 MinResistivity = 1E+20
 Do While Not EOF(1) ' Цикл до конца файла.
    Set gData = New Класс1
    Ind = Ind + 1
    Input #1, Tolsh, Resist, MyString    ' Читает данные в две переменные.
    'Debug.Print MyString, MyNumber  ' Выводит данные в окно отладки.
    MyString = UCase(Trim(MyString))
    gData.gPoroda = MyString
    gData.gMoshn = Tolsh
    gData.gResistivity = Resist
    If MaxResistivity < Resist Then MaxResistivity = Resist
    If MinResistivity > Resist Then MinResistivity = Resist
    'gData.gMoshn_LN = 1
    NameMoshnStrColl.Add Item:=gData, Key:=Str(Ind)
    'MsgBox Str(Ind) + " " + gData.gPoroda + " " + Str(gData.gMoshn), , "String"
    Set gData = Nothing
 Loop
 Close #1    ' Закрывает файл.
 '================ поиск минимального и максимального значения
 TheMax_Depth = 0
 For Each w In NameMoshnStrColl
  If w.gMoshn > 0 Then TheMax_Depth = TheMax_Depth + w.gMoshn
  If (w.gMoshn < MinTolsh) And (w.gMoshn > 0) Then MinTolsh = w.gMoshn
 Next w
 If MinTolsh < 1 Then
  TempSi = Int(1 / MinTolsh) + 1
 Else
  TempSi = 1
 End If
'================ назначение соотвествующего логарифма
If IsLogDepthScale Then
 MashFact = MaxHeightDraw / Log(TheMax_Depth)
Else
 MashFact = MaxHeightDraw / TheMax_Depth
End If
'MsgBox MashFact & " " & TempSi, , "MashFact TempSi"
CurrH = 0
 For Each w In NameMoshnStrColl
 CurrH = w.gMoshn + CurrH
  If IsLogDepthScale Then
   If w.gMoshn > 0 Then
    w.gMoshn_LN = Log(CurrH * TempSi) * MashFact
   Else
    w.gMoshn_LN = 0
   End If
   Else
    If w.gMoshn > 0 Then
     w.gMoshn_LN = w.gMoshn * MashFact
    Else
     w.gMoshn_LN = 0
    End If
   End If
  'MsgBox w.gMoshn_LN, , ""
 Next w
 '=========== проверка
' i = 0
' For Each w In NameMoshnStrColl
'  i = i + 1
'  MsgBox Str(i) + " " + Str(Int(w.gMoshn_LN * 10) / 10), , "I  gMoshn_LN - before"
' Next w
 
 '=========== НАЗНАЧЕНИЕ НОРМАЛЬНЫХ ВЕЛИСЧИН
 If IsLogDepthScale Then
  For i = NameMoshnStrColl.Count To 2 Step -1
   If i = NameMoshnStrColl.Count Then
    Set gData = NameMoshnStrColl(i)
    gData.gMoshn_LN = 0
    Set gData = Nothing
  Else
   Set gData = NameMoshnStrColl(i)
   Set w = NameMoshnStrColl(i - 1)
    gData.gMoshn_LN = gData.gMoshn_LN - w.gMoshn_LN
   Set w = Nothing
   Set gData = Nothing
   End If
  Next i
' i = 0: Sum = 0
' For Each w In NameMoshnStrColl
'  i = i + 1
'  Sum = Sum + w.gMoshn_LN
'  MsgBox Str(i) + " " + Str(Int(w.gMoshn_LN * 10) / 10), , "I  gMoshn_LN"
' Next w
 'MsgBox Sum, , "Sum"
 End If '======If IsLogDepthScale Then
 
 Application.ScreenUpdating = True
 'MsgBox TheMax_Depth, , "TheMax_Depth"
 End Sub
 
 
Sub CalcLogging2()
'======= Процедура использует вычисления координат для построения
'======= каротажа как графика. Это не нужно, так как я строю в виде отрезков
Dim NumNew As Byte, i, TempI As Integer ' число узлов ломаной
Dim Sum, Sc, StepOfRazrez As Double, TempS, TempS2, TempS3 As String
Dim LogCurveData As Класс2
Dim NameList As String
Dim TempSngl As Single
 '================ начало процедуры
 If LogData Is Nothing Then
  Set LogData = New Collection
 Else
  Set LogData = Nothing
  Set LogData = New Collection
 End If

 If NameMoshnStrColl Is Nothing Then Call Geology
 NumNew = NameMoshnStrColl.Count * 2
 'MsgBox NumNew, , "NumNew"
 Sum = 0
 
 Sc = (ScaleLogging / Log(MaxResistivity))
 For i = 1 To NumNew
  Set LogCurveData = New Класс2
  If i = 1 Then '=========== 0
   LogCurveData.xLogData = 0
   LogCurveData.yLogData = Log(NameMoshnStrColl(1).gResistivity) * Sc
   LogData.Add Item:=LogCurveData
   Sum = Sum + NameMoshnStrColl(1).gMoshn - StepOfRazrez
  Else
   If i = NumNew Then '=========== 1
    Sum = Sum * 1.1
    LogCurveData.xLogData = Sum
    LogCurveData.yLogData = Log(NameMoshnStrColl(NameMoshnStrColl.Count).gResistivity) * Sc
    LogData.Add Item:=LogCurveData
   Else
    If (i Mod 2) = 0 Then '=========== 2
      ' четное
     TempI = (i \ 2) + (i Mod 2)
     LogCurveData.xLogData = Sum
     LogCurveData.yLogData = Log(NameMoshnStrColl(TempI).gResistivity) * Sc
     LogData.Add Item:=LogCurveData
     Sum = Sum + 2 * StepOfRazrez
    Else
      ' нечетное
     TempI = (i \ 2) + (i Mod 2)
     LogCurveData.xLogData = Sum
     LogCurveData.yLogData = Log(NameMoshnStrColl(TempI).gResistivity) * Sc
     LogData.Add Item:=LogCurveData
     Sum = Sum + NameMoshnStrColl(TempI).gMoshn - 2 * StepOfRazrez
    End If '=========== 2
  End If '=========== 1
 End If '=========== 0
 Set LogCurveData = Nothing
Next i
NameList = ""
For i = 1 To LogData.Count
     Set LogCurveData = LogData(i)
     TempS = Str(i)
'     TempS2 = Str(LogCurveData.xLogData)
'     TempS3 = Str(LogCurveData.yLogData)
     TempS2 = Format(LogCurveData.xLogData, "###E+")
     TempS3 = Format(LogCurveData.yLogData, "###E-")
     NameList = NameList + TempS + _
     "  " + TempS2 + "  " + TempS3
     If (i Mod 2) = 0 Then
      NameList = NameList + vbLf
     Else
      NameList = NameList + "  |  "
     End If
     'NameList = NameList + Chr(13)
     Set LogCurveData = Nothing
Next i
 'MsgBox NameList, , "LogData.Count " + Str(LogData.Count)
 'Set LogData = Nothing
End Sub
'Sub CalcLogging()
''======= Процедура использует вычисления координат для построения
''======= каротажа как графика. Это не нужно, так как я строю в виде отрезков
'Dim NumNew As Byte, I, TempI As Integer ' число узлов ломаной
'Dim Sum, Sc, StepOfRazrez As Double, TempS, TempS2, TempS3 As String
'Dim LogCurveData As Класс2
'Dim NameList As String
'Dim TempSngl As Single
' '================ начало процедуры
' If LogData Is Nothing Then
'  Set LogData = New Collection
' Else
'  Set LogData = Nothing
'  Set LogData = New Collection
' End If
'
' If NameMoshnStrColl Is Nothing Then Call Geology
' NumNew = NameMoshnStrColl.Count * 2
' 'MsgBox NumNew, , "NumNew"
' Sum = 0
' StepOfRazrez = 0.0005
' 'g.Moshn = g.Moshn * ScaleLogging / MaxResistivity
' Sc = (ScaleLogging / Log(MaxResistivity))
' For I = 1 To NumNew
'  Set LogCurveData = New Класс2
'  If I = 1 Then '=========== 0
'   LogCurveData.xLogData = 0
'   LogCurveData.yLogData = Log(NameMoshnStrColl(1).gResistivity) * Sc
'   LogData.Add Item:=LogCurveData
'   Sum = Sum + NameMoshnStrColl(1).gMoshn - StepOfRazrez
'  Else
'   If I = NumNew Then '=========== 1
'    Sum = Sum * 1.1
'    LogCurveData.xLogData = Sum
'    LogCurveData.yLogData = Log(NameMoshnStrColl(NameMoshnStrColl.Count).gResistivity) * Sc
'    LogData.Add Item:=LogCurveData
'   Else
'    If (I Mod 2) = 0 Then '=========== 2
'      ' четное
'     TempI = (I \ 2) + (I Mod 2)
'     LogCurveData.xLogData = Sum
'     LogCurveData.yLogData = Log(NameMoshnStrColl(TempI).gResistivity) * Sc
'     LogData.Add Item:=LogCurveData
'     Sum = Sum + 2 * StepOfRazrez
'    Else
'      ' нечетное
'     TempI = (I \ 2) + (I Mod 2)
'     LogCurveData.xLogData = Sum
'     LogCurveData.yLogData = Log(NameMoshnStrColl(TempI).gResistivity) * Sc
'     LogData.Add Item:=LogCurveData
'     Sum = Sum + NameMoshnStrColl(TempI).gMoshn - 2 * StepOfRazrez
'    End If '=========== 2
'  End If '=========== 1
' End If '=========== 0
' Set LogCurveData = Nothing
'Next I
'NameList = ""
'For I = 1 To LogData.Count
'     Set LogCurveData = LogData(I)
'     TempS = Str(I)
''     TempS2 = Str(LogCurveData.xLogData)
''     TempS3 = Str(LogCurveData.yLogData)
'     TempS2 = Format(LogCurveData.xLogData, "###E+")
'     TempS3 = Format(LogCurveData.yLogData, "###E-")
'     NameList = NameList + TempS + _
'     "  " + TempS2 + "  " + TempS3
'     If (I Mod 2) = 0 Then
'      NameList = NameList + vbLf
'     Else
'      NameList = NameList + "  |  "
'     End If
'     'NameList = NameList + Chr(13)
'     Set LogCurveData = Nothing
'Next I
' 'MsgBox NameList, , "LogData.Count " + Str(LogData.Count)
' 'Set LogData = Nothing
'End Sub

'Sub DrawLogging()
'Dim LogCurveData1, LogCurveData2 As Класс2
'Dim WidthWW, LeftLL As Integer, xx, yy, xx1, yy1, ScaleDepth As Single
'Dim ListView As String
' '================ начало процедуры
'Application.ScreenUpdating = False
'Call DrawColumn
'Set LogData = Nothing
'If LogData Is Nothing Then Call CalcLogging
'TopYY = -80
'WidthWW = 20
'LeftLL = 0
'RealDraw = False
'
'ScaleDepth = 1.4
'MaxXX = -100
'MinXX = 10000000
'MaxYY = -100
'MinYY = 10000000
'
'Set LogCurveData = Nothing
' For I = 1 To LogData.Count - 1
'  Set LogCurveData1 = LogData(I)
'  Set LogCurveData2 = LogData(I + 1)
'  xx = CSng(LogCurveData1.yLogData)
'  yy = (CSng(LogCurveData1.xLogData) - TopYY) * ScaleDepth
'  xx1 = CSng(LogCurveData2.yLogData)
'  yy1 = (CSng(LogCurveData2.xLogData) - TopYY) * ScaleDepth
'    ActiveDocument.Shapes.AddLine(xx, yy, xx1, yy1).Select
'   If xx > MaxXX Then MaxXX = xx
'   If xx1 > MaxXX Then MaxXX = xx1
'   If yy > MaxYY Then MaxYY = yy
'   If yy1 > MaxYY Then MaxYY = yy1
'   If xx < MinXX Then MinXX = xx
'   If xx1 < MinXX Then MinXX = xx1
'   If yy < MinYY Then MinYY = yy
'   If yy1 < MinYY Then MinYY = yy1
'
'   xx1 = xx
'   yy1 = yy
'  Set LogCurveData2 = Nothing
'  Set LogCurveData1 = Nothing
' Next I
' 'Selection.WholeStory
' WidthShX = MaxXX - MinXX
' HeightY = MaxYY - MinYY
'  Set myRange = ActiveDocument.Sections(1).Range
'  myRange.ShapeRange.Select
'  Selection.ShapeRange.Group.Select
'  Application.ScreenUpdating = True
'End Sub

Sub DrawColumn()
Dim WidthWW, i, j, SavJ As Integer
Dim TopYY As Double, TmpSi As Single
Dim xx, yy, xx1, yy1, PrevX, CurrSum, TickSum As Single
Dim w As Класс1
Dim ColorEl As Class1

'================ начало процедуры
Selection.WholeStory
Selection.Cut
If NameMoshnStrColl Is Nothing Then Call Geology
Call AssMainRockWithCollection
Call LegendRebuilding

 '=============== прорисовка колонки
 Sc = (ScaleLogging / Log(MaxResistivity))
 MaxYY = -100
 MinYY = 10000000
 MaxXX = -100
 MinXX = 10000000

 TopYY = 110
 TopY_col = TopYY
 WidthWW = 20
 CurrSum = 0
 '==============определяем необходимый сдвиг каротажа
 CarotageShift = 6000
 For i = 1 To NameMoshnStrColl.Count
  Set w = NameMoshnStrColl(i)
  TmpSi = Log(w.gResistivity) * Sc
  If CarotageShift > TmpSi Then CarotageShift = TmpSi
  Set w = Nothing
 Next i
 If CarotageShift < LeftLL + WidthWW + 10 Then
  CarotageShift = LeftLL + WidthWW + 10 - CarotageShift
 Else
  CarotageShift = 4
 End If
  
 '===================================================
 TickSum = TheMax_Depth / NumTickY
 For i = 1 To NameMoshnStrColl.Count
  Set w = NameMoshnStrColl(i)
  '======== прорисовка линий каротажа
   TmpSi = CarotageShift + Log(w.gResistivity) * Sc
   yy = TopYY
   yy1 = TopYY + w.gMoshn_LN
  If yy > MaxYY Then MaxYY = yy
  If yy1 > MaxYY Then MaxYY = yy1
  If yy < MinYY Then MinYY = yy
  If yy1 < MinYY Then MinYY = yy1
  If TmpSi > MaxXX Then MaxXX = TmpSi
  If TmpSi < MinXX Then MinXX = TmpSi
  
  ActiveDocument.Shapes.AddLine(TmpSi, yy, TmpSi, yy1).Select
  'Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
  If (i <> 1) And (i <> NameMoshnStrColl.Count) Then
   ActiveDocument.Shapes.AddLine(TmpSi, yy, PrevX, yy).Select
  End If
  PrevX = TmpSi
  '======== прорисовка границ слоев
  xx = LeftLL + WidthWW
  xx1 = TmpSi
  If (i <> 1) And (i <> NameMoshnStrColl.Count) Then
   ActiveDocument.Shapes.AddLine(xx, TopYY, xx1, TopYY).Select
   Selection.ShapeRange.Line.DashStyle = msoLineRoundDot
  End If
  If Not IsLogDepthScale Then
  ActiveDocument.Shapes.AddShape(msoShapeRectangle, LeftLL, _
        TopYY, WidthWW, CSng(w.gMoshn_LN)).Select
  Else
  ActiveDocument.Shapes.AddShape(msoShapeRectangle, LeftLL, _
        TopYY, WidthWW, CSng(w.gMoshn_LN)).Select
  End If
  For j = 1 To MyStrColl.Count
   Set ColorEl = MyStrColl(j)
   'MsgBox Str(j) + " " + ColorEl.Poroda + " " + w.gPoroda, , ""
   If (Trim(ColorEl.Poroda) Like Trim(w.gPoroda)) Then
      SavJ = j
      Set ColorEl = Nothing
      Exit For    ' Завершает цикл.
   End If
   Set ColorEl = Nothing
  Next j
  Set ColorEl = MyStrColl(SavJ)
  Selection.ShapeRange.Fill.ForeColor.RGB = ColorEl.MyRGBColor
  Selection.ShapeRange.Fill.BackColor.RGB = ColorEl.MyRGBBackColor
  Selection.ShapeRange.Fill.Patterned ColorEl.MyPattern
  Set ColorEl = Nothing
  If Not IsLogDepthScale Then
   TopYY = TopYY + w.gMoshn_LN
  Else
   TopYY = TopYY + w.gMoshn_LN
  End If

  CurrSum = CurrSum + w.gMoshn
  Set w = Nothing
 Next i
 BotY_col = TopYY
 WidthShX = MaxXX - MinXX
 HeightY = MaxYY - MinYY

End Sub
Sub DrawCurve()
 XPosSecCurve
End Sub

Sub DrawLegend()
Dim i, j, NumEl, TheSize, TheSize2, Xtopp, Ytopp As Integer
Dim scR, scT, TempL, TempLC, TempLBC, TempPatt As Long
Dim MaxWidth, CurrWidth, Width, CurrHeight, Height, _
RWidth, RHeight As Single
Dim TempS As String
Dim MyObject As Class1
Dim TempRange As Range
Dim WithCombine As Boolean

'============== Начало процедуры
Set MyObject = Nothing
Application.ScreenUpdating = False
'============== Создание если надо коллекции
If MyStrColl Is Nothing Then
 'MsgBox "LegendRebuilding", , " Info "
 Call LegendRebuilding
End If
'===================================
NumEl = MyStrColl.Count
'MsgBox NumEl, , " "
scR = msoShapeRectangle
scT = msoTextOrientationHorizontal
MaxWidth = 640
RWidth = 10: RHeight = 10
Width = 20:   Height = 22
' = = = =  Заголовок легенды
WithCombine = True
If WithCombine Then
 Xtopp = 600
Else
 Xtopp = 600
End If
Xtopp = 100
Ytopp = 500
ActiveDocument.Shapes.AddTextbox(scT, _
                                 Xtopp, Ytopp, MaxWidth, 20).Select
Selection.ShapeRange.Line.Visible = msoFalse
Selection.ShapeRange.TextFrame.TextRange.Select
Selection.TypeText text:="УCЛОВНЫЕ ЗНАКИ"
Selection.Font.Name = "Times New Roman Cyr"
Selection.Font.Size = 14

If WithCombine Then
 CurrHeight = Ytopp + 20 + 1
Else
 CurrHeight = Ytopp + 5
End If
CurrWidth = 0
TheSize = 8
'= = = = = Элементы легенды
For i = 1 To NumEl
 Set MyObject = MyStrColl(i)
 If MyObject.ThisUsed Then
  If WithCombine Then
   If CurrWidth > MaxWidth Then '==== с переносом
    CurrHeight = CurrHeight + 20
    CurrWidth = 0
   End If
  Else
   CurrHeight = CurrHeight + 20
   CurrWidth = 0
  End If

  TempS = Trim(MyObject.Poroda)
  TempLC = MyObject.MyRGBColor
  TempLBC = MyObject.MyRGBBackColor
  TempPatt = MyObject.MyPattern
  ActiveDocument.Shapes.AddShape(scR, Xtopp + CurrWidth, CurrHeight, RWidth, RHeight).Select
  CurrWidth = CurrWidth + RWidth + 3
  With Selection.ShapeRange(1).Fill
    .Patterned (TempPatt)
    .ForeColor.RGB = TempLC
    .BackColor.RGB = TempLBC
  End With
  Set MyObject = Nothing
 
 TheSize2 = Len(Trim(TempS))
 'MsgBox "1" + TempS + " " + Str(TheSize2)
 Width = TheSize2 * (TheSize + 2)
 ActiveDocument.Shapes.AddTextbox(scT, Xtopp + CurrWidth, CurrHeight - 5, Width, Height).Select
 Selection.ShapeRange.Line.Visible = msoFalse
 Selection.ShapeRange.TextFrame.TextRange.Select
 Selection.TypeText text:=Trim(TempS)
    With Selection.Font
        .Name = "Times New Roman Cyr"
        .Size = TheSize
     End With
 TempS = ""
 CurrWidth = CurrWidth + Width
 'MsgBox CurrWidth & "     " & Width, , "CurrWidth         Width"
 Else
  Set MyObject = Nothing
 End If
Next i
ActiveDocument.Shapes.SelectAll
Selection.ShapeRange.Group
End Sub

Sub AssMainRockWithCollection()
    Dim Num ' Счетчик индексов.
    Dim Msg As String  ' Переменная для записи приглашения.
    Dim TheName As String * 20
    Dim Inst As Class1
    Dim MyStr As String
    Dim DefaultPattern As Long
'============== Начало процедуры
    
    DefaultPattern = 1
    Set MyStrColl = New Collection
    Num = 0 'хотя и это и делает по умолчанию DIM,
    'но делаю это еще раз для наглядности
    'Dim Inst As New Класс1  ' Создаем новый экземпляр Класс1.
    If Not Inst Is Nothing Then Set Inst = Nothing
    Set Inst = New Class1
    Num = Num + 1   ' Увеличиваем значение Num, а затем
    ' устанавливаем имя.
'---------------------- светло-желтый  (1)
    TheName = "ПЕСЧАНИК"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(255, 255, 180)
    Inst.MyRGBBackColor = RGB(255, 255, 180)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- белый   (2)
    Set Inst = New Class1
    Num = Num + 1: TheName = "НЕИЗВЕСТНОЕ"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(255, 255, 255)
    Inst.MyRGBBackColor = RGB(255, 255, 255)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- зеленый   (3)
    Set Inst = New Class1
    Num = Num + 1: TheName = "ИЗВЕСТНЯК"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(0, 255, 0)
    Inst.MyRGBBackColor = RGB(0, 255, 0)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- темно-голубой  (4)
    Set Inst = New Class1
    Num = Num + 1: TheName = "ГЛИНА"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(0, 64, 128)
    Inst.MyRGBBackColor = RGB(0, 64, 128)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- малиновый   (5)
    Set Inst = New Class1
    Num = Num + 1: TheName = "МЕРГЕЛЬ"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(255, 0, 255)
    Inst.MyRGBBackColor = RGB(255, 0, 255)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- светло-голубой  (6)
    Set Inst = New Class1
    Num = Num + 1: TheName = "ГИПС"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(128, 128, 255)
    Inst.MyRGBBackColor = RGB(128, 128, 255)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'Call MemoryInf("after Del")
'---------------------- красный   (7)
    Set Inst = New Class1
    Num = Num + 1: TheName = "АЛЕВРОЛИТ"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(255, 0, 0)
    Inst.MyRGBBackColor = RGB(255, 0, 0)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'----------------------  голубой   (8)
    Set Inst = New Class1
    Num = Num + 1: TheName = "АРГИЛЛИТ"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(0, 0, 255)
    Inst.MyRGBBackColor = RGB(0, 0, 255)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- Cyan  (9)
    Set Inst = New Class1
    Num = Num + 1: TheName = "КАМ.СОЛЬ"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(0, 255, 255)
    Inst.MyRGBBackColor = RGB(0, 255, 255)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
'---------------------- желтый  (10)
    Set Inst = New Class1
    Num = Num + 1: TheName = "ПЕСОК"
    Inst.Poroda = TheName: Inst.MyRGBColor = RGB(255, 255, 0)
    Inst.MyRGBBackColor = RGB(255, 255, 0)
    Inst.MyPattern = DefaultPattern
    'Inst.ThisUsed = False
    MyStrColl.Add Item:=Inst, Key:=TheName
    Set Inst = Nothing
 '---------------------------------------
    'Call PrintAllFromColl(MyStrColl)
    'For Num = 1 To MyStrColl.Count  ' Удаляем имя из семейства.
    '    MyStrColl.Remove 1  ' Поскольку семейства индексируются
    '            ' автоматически, каждый раз удаляем
    'Next        ' первый компонент семейства.
End Sub
Sub AssNewRGBColor_and_Patterns()
Dim i As Byte
'=======светло-серый
NewRGBColor(1).Red_color = 192
NewRGBColor(1).Green_color = 192
NewRGBColor(1).Blue_color = 192
'=======светло-зеленоватый
NewRGBColor(2).Red_color = 0
NewRGBColor(2).Green_color = 255
NewRGBColor(2).Blue_color = 128
'=======темно-зеленый
NewRGBColor(3).Red_color = 0
NewRGBColor(3).Green_color = 128
NewRGBColor(3).Blue_color = 0
For i = 1 To 12 'для назначения через interior.pattern
 PatternArray(1) = i
Next i
PatternArray(1) = msoPatternLightDownwardDiagonal 'msoPattern20Percent
PatternArray(2) = msoPattern30Percent
PatternArray(3) = msoPattern40Percent
PatternArray(4) = msoPattern60Percent
PatternArray(5) = msoPatternDarkHorizontal
PatternArray(6) = msoPatternDarkVertical
PatternArray(7) = msoPatternDarkDownwardDiagonal
PatternArray(8) = msoPatternDarkUpwardDiagonal
PatternArray(9) = msoPatternSmallCheckerBoard
PatternArray(10) = msoPatternTrellis
PatternArray(11) = msoPatternLightHorizontal
PatternArray(12) = msoPatternLightVertical
'PatternArray(13) = msoPatternLightDownwardDiagonal
'PatternArray(14) = msoPatternLightUpwardDiagonal
   'MsgBox msoPatternLightHorizontal, , "msoPatternLightHorizontal"
   'MsgBox msoPatternLightDownwardDiagonal, , "msoPatternLightDownwardDiagonal"
End Sub
Sub PrintAllFromColl(MyStrColl_ As Collection)
Dim MyObject, MyString, NameList ' Переменные типа Variant.
Dim TempS, TempS2, TempS3 As String
Dim TempI
'============== Начало процедуры
NameList = ""
For Each MyObject In MyStrColl_
     TempS = Trim(MyObject.Poroda)
     TempS2 = Format(MyObject.MyRGBColor, "###,###,###,###")
     TempS3 = Str(MyObject.MyPattern)
     MyString = "@@@@@@@"
     LSet MyString = TempS3
     TempS3 = MyString
     TempI = Len(TempS)
     NameList = NameList + TempS + _
     "         " + TempS2 + "           " + TempS3 + Chr(13)
Next MyObject
' Выводим список имен
    MsgBox NameList, , "Имена компонентов в семействе MyClasses = " + Str(MyStrColl_.Count)
End Sub
Sub LegendRebuilding()
Dim i, j, TempJ, NumNewColor, NumPattern As Integer
Dim RockName, Name, Name1, TempS, TempS2 As String
Dim TempB, TempBN, TempB3 As Boolean
Dim lr, lg, lb, loclen As Byte

Dim DataEl As Класс1
Dim ColorEl, MyObject As Class1
Dim CurrColor
'============ Добавление новых элементов в список слоев
Application.ScreenUpdating = False
Call AssNewRGBColor_and_Patterns
NumNewColor = 0: NumPattern = 0
Set Inst = Nothing
'=========== только для отладки - очистка коллекции
'Set MyStrColl = Nothing
If NameMoshnStrColl Is Nothing Then Call Geology
If MyStrColl Is Nothing Then Call AssMainRockWithCollection
  TempJ = NameMoshnStrColl.Count
 'MsgBox TempJ, , "GetNumLayer"
  For j = 1 To TempJ
   If NameMoshnStrColl(j).gMoshn > 0 Then
   RockName = NameMoshnStrColl(j).gPoroda
   RockName = Trim(UCase(RockName))
   TempB = False
   For Each MyObject In MyStrColl
    If Not TempB Then
     '======== определение повторения
     TempS = UCase(Trim(MyObject.Poroda))
     TempS2 = RockName
     TempBN = (TempS Like TempS2)
     'If TempBN Then CurrColor = MyObject.MyRGBColor
     TempB = TempB Or TempBN
     If TempBN Then '===== есть в колллекции - назначаем true
      MyObject.ThisUsed = True
     End If
     '======== определение цвета
     If Len(RockName) > Len(TempS) Then
      loclen = Len(TempS)
      TempS2 = Left(RockName, loclen)
      TempB3 = (TempS Like TempS2)
      If TempB3 Then CurrColor = MyObject.MyRGBColor
     End If
     '========
    End If
   Next MyObject
   If Not TempB Then 'новая порода
    MsgBox RockName, , "New Rock"
    NumNewColor = NumNewColor + 1
    NumPattern = NumPattern + 1
    Set ColorEl = New Class1
    ColorEl.Poroda = RockName
    ColorEl.MyRGBColor = CurrColor
    ColorEl.MyRGBBackColor = RGB(255, 255, 255)
    ColorEl.MyPattern = PatternArray(NumPattern)
    MyStrColl.Add Item:=Inst, Key:=(RockName)
    Set ColorEl = Nothing
    End If
   End If
  Next j
'Call PrintAllFromColl(MyStrColl)
'Call PrintAllFromColl(MyStrColl)
'Call RedrawDiagramm
'Call DrawLegend
'Call MemoryClearing
End Sub
Sub Ма22()
Attribute Ма22.VB_Description = "Макрос записан 23.11.97 Windows95"
Attribute Ма22.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.Ма22"
'
' Ма22 Макрос
' Макрос записан 23.11.97 Windows95
'
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
    ActiveDocument.Shapes.AddShape(msoShapeRectangle, 71#, 92.3, 695.8, _
        482.8).Select
End Sub
