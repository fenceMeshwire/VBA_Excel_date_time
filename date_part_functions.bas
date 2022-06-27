Option Explicit

Function get_year(ByVal dteDate As Date)
  get_year = DatePart("yyyy", dteDate)
End Function

Function get_quarter(ByVal dteDate As Date)
  get_quarter = DatePart("q", dteDate)
End Function

Function get_month(ByVal dteDate As Date)
  get_month = DatePart("m", dteDate)
End Function

Function get_day_of_year(ByVal dteDate As Date)
  ' Day of the year is to be understood as the day x of 365
  get_day_of_year = DatePart("y", dteDate)
End Function

Function get_day(ByVal dteDate As Date)
  get_day = DatePart("d", dteDate)
End Function

Function get_weekday(dteDate As Date)
  get_weekday = WeekdayName(Weekday(dteDate, vbUseSystemDayOfWeek), False)
End Function

Function get_calendar_weekday(dteDate As Date)
  get_calendar_weekday = DatePart("ww", dteDate)
End Function

Function get_hour(dteDate As Date)
  get_hour = DatePart("h", dteDate)
End Function

Function get_minute(dteDate As Date)
  get_minute = DatePart("n", dteDate)
End Function

Function get_second(dteDate As Date)
  get_second = DatePart("s", dteDate)
End Function

Sub get_date_parts()

Dim dteDate As Date

dteDate = "20.03.2022 12:30:15"

Debug.Print get_year(dteDate)
Debug.Print get_quarter(dteDate)
Debug.Print get_day_of_year(dteDate)
Debug.Print get_month(dteDate)
Debug.Print get_day(dteDate)
Debug.Print get_weekday(dteDate)
Debug.Print get_calendar_weekday(dteDate)
Debug.Print get_hour(dteDate)
Debug.Print get_minute(dteDate)
Debug.Print get_second(dteDate)

End Sub
