VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Calculator_BEL()

Application.ScreenUpdating = False

N = 2

ThisWorkbook.Worksheets("CFS_BEL").Select

While Not ThisWorkbook.Worksheets("CFS_BEL").Cells(N, 1) = ""
  N = N + 1
Wend
  N = N - 1

For i = 2 To N
  ThisWorkbook.Worksheets("CFS_BEL").Range(Cells(i, 1), Cells(i, 5)).Copy
  ThisWorkbook.Worksheets("BEL").Select
  Cells(9, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False

  Application.Calculation = xlAutomatic
  
  Application.Calculation = xlManual
  
  Range(Cells(9, 6), Cells(9, 1280)).Copy
  ThisWorkbook.Worksheets("CFS_BEL").Select
  Cells(i, 6).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
  

Next i

ThisWorkbook.Save

End Sub




Sub Calculator_PAD()

Application.ScreenUpdating = False

N = 2

ThisWorkbook.Worksheets("CFS_PAD").Select

While Not ThisWorkbook.Worksheets("CFS_PAD").Cells(N, 1) = ""
  N = N + 1
Wend
  N = N - 1

For i = 2 To N
  ThisWorkbook.Worksheets("CFS_PAD").Range(Cells(i, 1), Cells(i, 5)).Copy
  ThisWorkbook.Worksheets("PAD").Select
  Cells(9, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False

  Application.Calculation = xlAutomatic
  
  Application.Calculation = xlManual
  
  Range(Cells(9, 6), Cells(9, 1280)).Copy
  ThisWorkbook.Worksheets("CFS_PAD").Select
  Cells(i, 6).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
  

Next i

ThisWorkbook.Save

End Sub





Sub Calculator()

Application.ScreenUpdating = False

N = 2

ThisWorkbook.Worksheets("CFS_BEL").Select

While Not ThisWorkbook.Worksheets("CFS_BEL").Cells(N, 1) = ""
  N = N + 1
Wend
  N = N - 1

For i = 2 To N
  ThisWorkbook.Worksheets("CFS_BEL").Range(Cells(i, 1), Cells(i, 5)).Copy
  ThisWorkbook.Worksheets("BEL").Select
  Cells(9, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False

  Application.Calculation = xlAutomatic
  
  Application.Calculation = xlManual
  
  Range(Cells(9, 6), Cells(9, 1280)).Copy
  ThisWorkbook.Worksheets("CFS_BEL").Select
  Cells(i, 6).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
  

Next i

Application.ScreenUpdating = False

N = 2

ThisWorkbook.Worksheets("CFS_PAD").Select

While Not ThisWorkbook.Worksheets("CFS_PAD").Cells(N, 1) = ""
  N = N + 1
Wend
  N = N - 1

For i = 2 To N
  ThisWorkbook.Worksheets("CFS_PAD").Range(Cells(i, 1), Cells(i, 5)).Copy
  ThisWorkbook.Worksheets("PAD").Select
  Cells(9, 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False

  Application.Calculation = xlAutomatic
  
  Application.Calculation = xlManual
  
  Range(Cells(9, 6), Cells(9, 1280)).Copy
  ThisWorkbook.Worksheets("CFS_PAD").Select
  Cells(i, 6).PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
  

Next i

ThisWorkbook.Save

End Sub

