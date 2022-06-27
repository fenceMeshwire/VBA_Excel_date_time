Option Explicit

Sub get_date_parts()

Dim dteDate As Date

Dim int_year As Integer
Dim int_quarter As Integer
Dim int_month As Integer
Dim int_day_of_the_year As Integer
Dim int_day As Integer
Dim str_weekday As String
Dim int_calender_week As Integer
Dim int_hour As Integer
Dim int_minute As Integer
Dim int_second As Integer

dteDate = "20.03.2022 12:30:15"

int_year = DatePart("yyyy", dteDate)
int_quarter = DatePart("q", dteDate)
int_month = DatePart("m", dteDate)
int_day_of_the_year = DatePart("y", dteDate)
int_day = DatePart("d", dteDate)
str_weekday = WeekdayName(Weekday(dteDate, vbUseSystemDayOfWeek), False)
int_calender_week = DatePart("ww", dteDate)
int_hour = DatePart("h", dteDate)
int_minute = DatePart("n", dteDate)
int_second = DatePart("s", dteDate)

Debug.Print int_year
Debug.Print int_quarter
Debug.Print int_month
Debug.Print int_day_of_the_year
Debug.Print int_day
Debug.Print str_weekday
Debug.Print int_calender_week
Debug.Print int_hour
Debug.Print int_minute
Debug.Print int_second

End Sub
