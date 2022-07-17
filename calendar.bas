Option Explicit

' PLEASE NOTE: Rename the Worksheet for output to "Sheet1"

' First day of the year
Const dteCalBeginDate = "01.01.2022"
' Bank holidays
Const dteNeujahr = "01.01.2022"
Const dteHlgDreiKoenige = "06.01.2022"
Const dteKarfreitag = "15.04.2022"
Const dteOsterSonntag = "17.04.2022"
Const dteOstermontag = "18.04.2022"
Const dteTagDerArbeit = "01.05.2022"
Const dteChristiHimmelfahrt = "26.05.2022"
Const dtePfingstsonntag = "05.06.2022"
Const dtePfingstMontag = "06.06.2022"
Const dteFronleichnam = "16.06.2022"
Const dteMHimmelfahrt = "15.08.2022"
Const dteTagDEinheit = "03.10.2022"
Const dteAllerheiligen = "01.11.2022"
Const dteWeihnachtsfeiertag1 = "25.12.2022"
Const dteWeihnachtsfeiertag2 = "26.12.2022"
'Previous Christmas vacations
Const dtePrevChristmasStart = "01.01.2022"
Const dtePrevChristmasEnd = "06.01.2022"
' Winter vacations
Const dteWinterStart = "28.02.2022"
Const dteWinterEnd = "04.03.2022"
' Easter vacations
Const dteEasterStart = "11.04.2022"
Const dteEasterEnd = "23.04.2022"
' Pfingsten vacations
Const dtePfingstenStart = "07.06.2022"
Const dtePfingstenEnd = "18.06.2022"
' Summer vacations
Const dteSummerStart = "01.08.2022"
Const dteSummerEnd = "12.09.2022"
' Autumn vacations
Const dteAutumnStart = "31.10.2022"
Const dteAutumnEnd = "04.11.2022"
' Christmas vacations
Const dteChristmasStart = "24.12.2022"
Const dteChristmasEnd = "31.12.2022"

' ======================================================================
Sub Calendar()

Dim intRow, intRowMax, intCol, intColMax As Integer
Dim intMonth As Integer
Dim bleCancel As Boolean
Dim dteDate, dteToday As Date
Dim strMonth As String

Dim varDat(1 To 12) As Variant

varDat(1) = "January"
varDat(2) = "February"
varDat(3) = "March"
varDat(4) = "April"
varDat(5) = "May"
varDat(6) = "June"
varDat(7) = "July"
varDat(8) = "August"
varDat(9) = "September"
varDat(10) = "October"
varDat(11) = "November"
varDat(12) = "December"

dteDate = dteCalBeginDate
dteToday = Format(Now(), "dd.mm.yyyy")
Sheet1.UsedRange.Clear

intColMax = 12
For intCol = 1 To intColMax
  Sheet1.Cells(1, intCol).Value = varDat(intCol)
  Sheet1.Cells(1, intCol).BorderAround ColorIndex:=1, Weight:=xlThin
  intRow = 2
  Do While bleCancel = False
    intMonth = Month(dteDate)
    If intMonth <> intCol Then
      bleCancel = True
      Exit Do
    Else
      Sheet1.Cells(intRow, intCol).Value = dteDate
      Sheet1.Cells(intRow, intCol).BorderAround _
        ColorIndex:=1, Weight:=xlThin
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

With Sheet1
  
  If IsDate(.Cells(intRow, intCol).Value) Then
  dteDate = .Cells(intRow, intCol).Value
    ' Previous christmas holidays
    If dtePrevChristmasStart <= dteDate And dteDate <= dtePrevChristmasEnd Then
      If Weekday(.Cells(intRow, intCol).Value, vbMonday) <= 5 Then
        Sheet1.Cells(intRow, intCol).Interior.ColorIndex = 15
      End If
    End If
    ' Winter holidays
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
Sheet1.Cells(intRowMax + 3, 2).Value = "School Vacation for Bavaria, Germany"
Sheet1.Cells(intRowMax + 4, 2).Value = "Bank holiday"
Sheet1.Cells(intRowMax + 5, 2).Value = "Days passed in current year"

End Sub
