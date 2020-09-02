Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub RunForm()

UserForm1.Show

End Sub

Sub DeleteCharts()

Dim cht As ChartObject

For Each cht In ActiveSheet.ChartObjects
    cht.Delete
Next

End Sub
