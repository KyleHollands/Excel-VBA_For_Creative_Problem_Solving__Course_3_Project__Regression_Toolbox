VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Regression Toolbox"
   ClientHeight    =   5175
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   7245
   OleObjectBlob   =   "UserForm1 - Button Scripts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub GoButton_Click()

Dim tWS As Worksheet
Dim UserXRange As Range, UserYRange As Range, YPRange As Range
Dim funcArray As Variant, fxn() As Variant, xData() As Variant, yData() As Variant, X() As Variant, smX() As Variant
Dim y As Variant, yp() As Variant, res() As Variant, xT() As Variant, xTx() As Variant, xTxInv() As Variant, xTy() As Variant, beta() As Variant
Dim j As Integer, i As Integer, n As Integer, xVar As Integer, numFunc As Integer, a As Integer, aa As Integer, Ans As Integer, m As Integer
Dim SSe As Double, ySum As Double, yAve As Double, SSt As Double, radJ As Double

Set tWS = Sheet1

'ACQUIRE NUMBER OF FUNCTIONS ENTERED BY USER

funcArray = Array(fxn1, fxn2, fxn3, fxn4)
j = 1
For i = 1 To 4
    If funcArray(i) <> "" Then
        ReDim Preserve fxn(j)
        fxn(j) = funcArray(i)
        j = j + 1: numFunc = j - 1
    End If
Next i

If j = 1 Then
    MsgBox ("Please enter at least one function."): Exit Sub: End If
    
xVar = WorksheetFunction.CountA(Columns("a:a"))

If xVar < j + 2 Then
    MsgBox ("Data points must be 2 greater than input functions."): Exit Sub: End If
    
ReDim X(xVar, numFunc + 1) As Variant, smX(xVar), y(xVar) As Variant
ReDim yp(xVar, 1) As Variant, res(xVar) As Variant
ReDim xData(xVar, 1) As Variant, yData(xVar, 1) As Variant
    
Set UserXRange = Application.InputBox("X Input Range", "X Input", "Sheet1!$A$1:$A" & xVar, Type:=8)
Set UserYRange = Application.InputBox("Y Input Range", "Y Input", "Sheet1!$B$1:$B" & xVar, Type:=8)
Set YPRange = Worksheets("Sheet1").Range("A1:A" & xVar)

For i = 1 To xVar
    smX(i) = UserXRange(i)
    X(i, 1) = 1
    For j = 2 To numFunc + 1
        If UserForm1.Controls("fxn" & j - 1).Value = "x" Then
            UserForm1.Controls("fxn" & j - 1).Value = "1x"
            X(i, j) = Evaluate(Replace(UserForm1.Controls("fxn" & j - 1).Value, "x", smX(i)))
        Else
            X(i, j) = Evaluate(Replace(UserForm1.Controls("fxn" & j - 1).Value, "x", smX(i)))
        End If
    Next j
Next i

'TRANSPOSE ARRAYS------------------------------------------------------------------------------------------
y = UserYRange
xT = WorksheetFunction.Transpose(X)
xTx = WorksheetFunction.MMult(xT, X)
xTxInv = WorksheetFunction.MInverse(xTx)
xTy = WorksheetFunction.MMult(xT, y)
beta = WorksheetFunction.MMult(xTxInv, xTy)

'DISPLAY MODEL TO USER--------------------------------------------------------------------------------------
Select Case numFunc
    Case Is = 4
        MsgBox ("Model is y = " & FormatNumber(beta(1, 1), 4) & " + " & FormatNumber(beta(2, 1), 4) & " * " & _
        UserForm1.fxn1.Value & " + " & FormatNumber(beta(3, 1), 4) & " * " & UserForm1.fxn2.Value & " + " & _
        FormatNumber(beta(4, 1), 4) & " * " & UserForm1.fxn3.Value & " + " & FormatNumber(beta(5, 1), 4) & _
        " * " & UserForm1.fxn4.Value)
    Case Is = 3
        MsgBox ("Model is y = " & FormatNumber(beta(1, 1), 4) & " + " & FormatNumber(beta(2, 1), 4) & " * " & _
        UserForm1.fxn1.Value & " + " & FormatNumber(beta(3, 1), 4) & " * " & UserForm1.fxn2.Value & " + " & _
        FormatNumber(beta(4, 1), 4) & " * " & UserForm1.fxn3.Value)
    Case Is = 2
        MsgBox ("Model is y = " & FormatNumber(beta(1, 1), 4) & " + " & FormatNumber(beta(2, 1), 4) & " * " & _
        UserForm1.fxn1.Value & " + " & FormatNumber(beta(3, 1), 4) & " * " & UserForm1.fxn2.Value)
    Case Else
        MsgBox ("Model is y = " & FormatNumber(beta(1, 1), 4) & " + " & FormatNumber(beta(2, 1), 4) & " * " & _
        UserForm1.fxn1.Value)
End Select

'RESIDUALS CALCULATIONS--------------------------------------------------------------------------------

Select Case numFunc
    Case Is = 4
        For a = 1 To xVar
            yp(a, 1) = beta(1, 1) + beta(2, 1) * Evaluate(Replace(UserForm1.fxn1.Value, "x", smX(a))) _
            + beta(3, 1) * Evaluate(Replace(UserForm1.fxn2.Value, "x", smX(a))) _
            + beta(4, 1) * Evaluate(Replace(UserForm1.fxn3.Value, "x", smX(a))) _
            + beta(5, 1) * Evaluate(Replace(UserForm1.fxn4.Value, "x", smX(a)))
            res(a) = yp(a, 1) - y(a, 1): SSe = SSe + res(a) ^ 2: ySum = ySum + y(a, 1)
        Next a
        yAve = ySum / xVar
        For aa = 1 To xVar
            SSt = SSt + (y(aa, 1) - yAve) ^ 2
        Next aa
    Case Is = 3
        For a = 1 To xVar
            yp(a, 1) = beta(1, 1) + beta(2, 1) * Evaluate(Replace(UserForm1.fxn1.Value, "x", smX(a))) _
            + beta(3, 1) * Evaluate(Replace(UserForm1.fxn2.Value, "x", smX(a))) _
            + beta(4, 1) * Evaluate(Replace(UserForm1.fxn3.Value, "x", smX(a)))
            res(a) = yp(a, 1) - y(a, 1): SSe = SSe + res(a) ^ 2: ySum = ySum + y(a, 1)
        Next a
        yAve = ySum / xVar
        For aa = 1 To xVar
            SSt = SSt + (y(aa, 1) - yAve) ^ 2
        Next aa
    Case Is = 2
        For a = 1 To xVar
            yp(a, 1) = beta(1, 1) + beta(2, 1) * Evaluate(Replace(UserForm1.fxn1.Value, "x", smX(a))) _
            + beta(3, 1) * Evaluate(Replace(UserForm1.fxn2.Value, "x", smX(a)))
            res(a) = yp(a, 1) - y(a, 1): SSe = SSe + res(a) ^ 2: ySum = ySum + y(a, 1)
        Next a
        yAve = ySum / xVar
        For aa = 1 To xVar
            SSt = SSt + (y(aa, 1) - yAve) ^ 2
        Next aa
    Case Else
        For a = 1 To xVar
            yp(a, 1) = beta(1, 1) + beta(2, 1) * Evaluate(Replace(UserForm1.fxn1.Value, "x", smX(a)))
            res(a) = yp(a, 1) - y(a, 1): SSe = SSe + res(a) ^ 2: ySum = ySum + y(a, 1)
        Next a
        yAve = ySum / xVar
        For aa = 1 To xVar
            SSt = SSt + (y(aa, 1) - yAve) ^ 2
        Next aa
End Select

'DISPLAY ADJUSTED R-SQUARED TO USER-------------------------------------------------------
radJ = 1 - ((SSe / (xVar - numFunc - 1)) / (SSt / (xVar - 1)))
MsgBox ("Adjusted R-Squared is " & FormatNumber(radJ, 3))

Ans = MsgBox("Would you like to plot the data?", vbYesNo)
If Ans = 6 Then
    Call Plotting(xVar, yp)
End If

End Sub

Private Sub QuitButton_Click()

Unload UserForm1

End Sub

Private Sub UserForm_Click()

End Sub

Sub Plotting(xVar As Integer, yp As Variant)

    Sheets("Sheet1").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatter
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$1:$B" & xVar)
    ActiveChart.Legend.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    Selection.Format.TextFrame2.TextRange.Characters.Text = "X"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleRotated)
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Y"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Y"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 1).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).XValues = Range("Sheet1!$A$1:$A" & xVar)
    ActiveChart.SeriesCollection(2).Values = yp
    ActiveChart.SeriesCollection(2).Select
    ActiveChart.SeriesCollection(2).AxisGroup = 2
    ActiveChart.SeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.25
        .Transparency = 0
    End With
    Selection.MarkerStyle = -4142
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.PlotArea.Select
    ActiveChart.SeriesCollection(1).Name = "=""Experimental Data"""
    ActiveChart.SeriesCollection(2).Name = "=""Model Predictions"""
    
    Unload UserForm1

End Sub


