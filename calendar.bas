Option Explicit

Sub Calendar()

Dim intRow, intRowMax, intCol, intColMax As Integer
Dim intMonth As Integer
Dim bleCancel As Boolean
Dim dteDate, dteToday As Date
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
dteToday = Format(Now(), "dd.mm.yyyy")
Sheet1.UsedRange.Clear

intColMax = 12
For intCol = 1 To intColMax
  Sheet1.Cells(1, intCol).Value = objValuesMonth(intCol - 1)
  intRow = 2
  Do While bleCancel = False
    intMonth = Month(dteDate)
    Debug.Print (intMonth)
    If intMonth <> intCol Then
      bleCancel = True
      Exit Do
    Else
      Sheet1.Cells(intRow, intCol).Value = dteDate
      Call Weekend(intRow, intCol)
      Call Vacations(intRow, intCol)
      Call BankHoliday(intRow, intCol)
      If dteDate <= dteToday Then
        Sheet1.Cells(intRow, intCol).Value = "X " & Sheet1.Cells(intRow, intCol).Value
      End If
      dteDate = DateAdd("d", 1, dteDate)
      intRow = intRow + 1
    End If
  Loop
  bleCancel = False
Next intCol

Call InformationLegend

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

' ======================================================================
Sub Vacations(intRow, intCol)

Dim dteDate As Date
Dim dteWinterStart, dteWinterEnd As Date
Dim dteEasterStart, dteEasterEnd As Date
Dim dtePfingstenStart, dtePfingstenEnd As Date
Dim dteSummerStart, dteSummerEnd As Date
Dim dteAutumnStart, dteAutumnEnd As Date
Dim dteChristmasStart, dteChristmasEnd As Date

' Winter Vacations
dteWinterStart = "28.02.2022"
dteWinterEnd = "04.03.2022"

' Easter Vacations
dteEasterStart = "11.04.2022"
dteEasterEnd = "23.04.2022"

' Pfingsten
dtePfingstenStart = "07.06.2022"
dtePfingstenEnd = "18.06.2022"

' Summer Vacations
dteSummerStart = "01.08.2022"
dteSummerEnd = "12.09.2022"

' Autumn Vacations
dteAutumnStart = "31.10.2022"
dteAutumnEnd = "04.11.2022"

' Christmas Vacations
dteChristmasStart = "24.12.2022"
dteChristmasEnd = "31.12.2022"

With Sheet1
  
  If IsDate(.Cells(intRow, intCol).Value) Then
  dteDate = .Cells(intRow, intCol).Value
    ' Winter holiday
    If dteWinterStart <= dteDate And dteDate <= dteWinterEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
    ' Easter holidays
     If dteEasterStart <= dteDate And dteDate <= dteEasterEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
    ' Pfingsten holidays
     If dtePfingstenStart <= dteDate And dteDate <= dtePfingstenEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
    ' Summer holidays
     If dteSummerStart <= dteDate And dteDate <= dteSummerEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
    ' Autumn holidays
     If dteAutumnStart <= dteDate And dteDate <= dteAutumnEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
    ' Christmas Vacations
     If dteChristmasStart <= dteDate And dteDate <= dteChristmasEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
  End If
End With

End Sub

' ======================================================================
Sub BankHoliday(intRow, intCol)

Dim dteDate As Date
Dim dteNeujahr, dteHlgDreiKoenige, dteKarfreitag, dteOsterSonntag As Date
Dim dteOstermontag, dteTagDerArbeit, dteChristiHimmelfahrt, dtePfingstsonntag As Date
Dim dtePfingstMontag, dteFronleichnam, dteMHimmelfahrt, dteTagDEinheit As Date
Dim dteAllerheiligen, dteWeihnachtsfeiertag1, dteWeihnachtsfeiertag2 As Date

dteNeujahr = "01.01.2022"
dteHlgDreiKoenige = "06.01.2022"
dteKarfreitag = "15.04.2022"
dteOsterSonntag = "17.04.2022"
dteOstermontag = "18.04.2022"
dteTagDerArbeit = "01.05.2022"
dteChristiHimmelfahrt = "26.05.2022"
dtePfingstsonntag = "05.06.2022"
dtePfingstMontag = "06.06.2022"
dteFronleichnam = "16.06.2022"
dteMHimmelfahrt = "15.08.2022"
dteTagDEinheit = "03.10.2022"
dteAllerheiligen = "01.11.2022"
dteWeihnachtsfeiertag1 = "25.12.2022"
dteWeihnachtsfeiertag2 = "26.12.2022"

If IsDate(Sheet1.Cells(intRow, intCol).Value) Then
  dteDate = Sheet1.Cells(intRow, intCol).Value
  Select Case dteDate
    Case dteNeujahr: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteHlgDreiKoenige: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteKarfreitag: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteOsterSonntag: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteOstermontag: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteTagDerArbeit: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteChristiHimmelfahrt: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dtePfingstsonntag: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dtePfingstMontag: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteFronleichnam: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteMHimmelfahrt: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteTagDEinheit: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteAllerheiligen: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteWeihnachtsfeiertag1: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
    Case dteWeihnachtsfeiertag2: Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 50
  End Select
End If

End Sub

' ======================================================================
Sub InformationLegend()

Dim intRow, intRowMax As Integer

intRowMax = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row

Sheet1.Cells(intRowMax + 2, 1).Interior.ColorIndex = 4
Sheet1.Cells(intRowMax + 3, 1).Interior.ColorIndex = 15
Sheet1.Cells(intRowMax + 4, 1).Interior.ColorIndex = 50
Sheet1.Cells(intRowMax + 5, 1).Value = "X"
Sheet1.Cells(intRowMax + 2, 2).Value = "Weekend"
Sheet1.Cells(intRowMax + 3, 2).Value = "School Vacation for Bavaria"
Sheet1.Cells(intRowMax + 4, 2).Value = "Bank holiday"
Sheet1.Cells(intRowMax + 5, 2).Value = "Days passed in current year"

End Sub
