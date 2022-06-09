Option Explicit

Sub Calendar()

Dim intRow, intRowMax, intCol, intColMax As Integer
Dim intMonth As Integer
Dim bleCancel As Boolean
Dim dteDate As Date
Dim strMonth As String

Dim dictMonth As Object
Dim objKeysMonth, objValuesMonth

Set dictMonth = CreateObject("Scripting.Dictionary")
    
dictMonth.Add "1", "January"
dictMonth.Add "2", "February"
dictMonth.Add "3", "March"
dictMonth.Add "4", "April"
dictMonth.Add "5", "May"
dictMonth.Add "6", "June"
dictMonth.Add "7", "July"
dictMonth.Add "8", "August"
dictMonth.Add "9", "September"
dictMonth.Add "10", "October"
dictMonth.Add "11", "November"
dictMonth.Add "12", "December"

objKeysMonth = dictMonth.keys
objValuesMonth = dictMonth.items

dteDate = "01.01.2022"

intColMax = 12
For intCol = 1 To intColMax
  Sheet1.Cells(1, intCol).Value = objValuesMonth(intCol - 1)
  intRow = 2
  Do While bleCancel = False
    intMonth = Month(dteDate)
    If intMonth <> intCol Then
      bleCancel = True
      Exit Do
    Else
      Sheet1.Cells(intRow, intCol).Value = dteDate
      Call Weekend(intRow, intCol)
      dteDate = dteDate + 1
      intRow = intRow + 1
    End If
  Loop
  bleCancel = False
Next intCol

End Sub

' ======================================================================

Sub Weekend(intRow, intCol)

With Sheet1
  If IsDate(.Cells(intRow, intCol).Value) Then
    If Weekday(.Cells(intRow, intCol).Value, vbMonday) > 5 Then
      Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 4
    End If
  End If
End With

End Sub
