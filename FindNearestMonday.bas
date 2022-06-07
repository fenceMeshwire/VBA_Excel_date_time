Option Explicit

' First add a few date values in tbl_Calendar.Range("A1:A" & your_last_row)  

' Find nearest monday
Sub AddDaysCalendar() 

Dim lngRow, lngRowMax As Long
Dim dteDate, dteDateEntry As Date

With tbl_Calendar
  lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
  For lngRow = 1 To lngRowMax
    If IsDate(.Cells(lngRow, 1).Value) Then
      dteDate = .Cells(lngRow, 1).Value
      dteDateEntry = FindMonday(dteDate)
      .Cells(lngRow, 2).Value = dteDateEntry
    End If
  Next lngRow
End With

End Sub

Function FindMonday(ByVal dteDate As Date) As Date

Dim intWeekday, intDateAdd As Integer

intDateAdd = -14
intWeekday = Weekday(dteDate)

Select Case intWeekday
  Case 1: FindMonday = DateAdd("d", intDateAdd, dteDate)
  Case 2: FindMonday = DateAdd("d", intDateAdd - 1, dteDate)
  Case 3: FindMonday = DateAdd("d", intDateAdd - 2, dteDate)
  Case 4: FindMonday = DateAdd("d", intDateAdd - 3, dteDate)
  Case 5: FindMonday = DateAdd("d", intDateAdd - 4, dteDate)
  Case 6: FindMonday = DateAdd("d", intDateAdd - 5, dteDate)
  Case 7: FindMonday = DateAdd("d", intDateAdd - 6, dteDate)
End Select

End Function

Sub DateTest()

Dim lngRow, lngRowMax As Long
Dim dteDate, dteDateTest As Date
Dim intDate As Integer

With tbl_Calendar
  lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
  For lngRow = 1 To lngRowMax
    dteDate = .Cells(lngRow, 1).Value
    dteDateTest = .Cells(lngRow, 2).Value
    intDate = DateDiff("d", dteDateTest, dteDate)
    .Cells(lngRow, 3).Value = intDate
  Next lngRow
End With

End Sub

Sub WeekDayTest()

Dim lngRow, lngRowMax As Long
Dim dteDate, dteDateTest As Date
Dim intWeekday As Integer
Dim strWeekday As String

With tbl_Calendar
  lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
  For lngRow = 1 To lngRowMax
    dteDateTest = .Cells(lngRow, 2).Value
    intWeekday = Weekday(.Cells(lngRow, 2).Value)
    Select Case intWeekday
      Case 1: strWeekday = "Monday"
      Case 2: strWeekday = "Tuesday"
      Case 3: strWeekday = "Wednesday"
      Case 4: strWeekday = "Thursday"
      Case 5: strWeekday = "Friday"
      Case 6: strWeekday = "Saturday"
      Case 7: strWeekday = "Sunday"
    End Select
    .Cells(lngRow, 4).Value = strWeekday
  Next lngRow
End With

End Sub

' Additional macro to clear the testing area at tbl_Calendar.Range("B1:D" & lngRowMax)
Sub ClearTestingArea()

Dim lngRow, lngRowMax As Long

With tbl_Calendar

lngRowMax = .Cells(.Rows.Count, 1).End(xlUp).Row
.Range("B1:D" & lngRowMax).Clear

End With

End Sub
