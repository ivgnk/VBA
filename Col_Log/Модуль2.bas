Attribute VB_Name = "Модуль2"
Sub MyPageSetup()
    With ActiveDocument.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientLandscape
        .TopMargin = CentimetersToPoints(3.3)
        .BottomMargin = CentimetersToPoints(1)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.25)
        .FooterDistance = CentimetersToPoints(1.25)
        .PageWidth = CentimetersToPoints(29.7)
        .PageHeight = CentimetersToPoints(21)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
    End With
 End Sub
 
 Sub Div_Man_por(ByVal s1 As Double, Man As Double, Por As Integer)
 Dim i, savsign As Integer, s As Double
  If s1 < 0 Then
   savsign = -1
  Else
   savsign = 1
  End If
  s = Abs(s1)
  Man = s
  Por = 0
  
  If ((s >= 1) And (s < 10)) Or (s = 0) Then
   GoTo 2
  ElseIf 10 <= s Then
1:   s = s / 10
     Man = s
     Por = Por + 1
     If s >= 10 Then
      GoTo 1
     Else
      GoTo 2
     End If
       Else
3:   s = s * 10
     Man = s
     Por = Por - 1
     If s < 1 Then
      GoTo 3
     Else
      GoTo 2
     End If
    End If
2: Man = Man * savsign
End Sub

Sub TEST()
 Dim Dat, Man As Double, Por As Integer
 Dat = -0.12
 Call Div_Man_por(Dat, Man, Por)
 Application.ScreenUpdating = False
 MsgBox Str(Dat) + "    " + Str(Man) + "    " + Str(Por), , "Dat Man Por"
End Sub
