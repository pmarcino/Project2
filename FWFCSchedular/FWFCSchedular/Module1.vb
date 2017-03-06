'Patrick Marcino
'tested for the dates 2010 - 2030

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Data.SqlClient
Module ModuleHolidays
    Sub GenerateHolidays()
        'these variables will add a year to the current year to calculate all of the holidays and weekends for the next year
        'curretly the applciation will compile completely for 2016. However, for any year other than 2016 it will not compile for the weekends, but it 
        'will still find all of the holidays for that year
        Dim nextYear As DateTime = DateTime.Today
        nextYear = nextYear.AddYears(1)

        'for testing years that aren't next year, assign the year to int y. For instance if you want to calculate all the holidays/weekends in the year 2033
        'change make int y = 2033;
        Dim y As Integer = nextYear.Year
        findWeekends(y)
        findHolidays(y)
    End Sub
    Public Sub findHolidays(y As Integer)
        'Find all of the holidays in the year
        Dim startValue As New DateTime(y, 1, 1)
        Dim endValue As DateTime = startValue.AddMonths(11)
        Dim arrayHolidays As DateTime() = New DateTime(7) {}
        'Static Holidays are the Holidays that are on the same date each year
        findStaticHolidays(y, arrayHolidays)
        'Find Easter   -  Subtract two days from this to find Easter Friday
        findEaster(y, arrayHolidays, startValue, endValue)
        'Dynamic Holidays are Holidays that land on a different date each year
        findDynamicHolidays(arrayHolidays, startValue, endValue)
        Dim count As Integer = 0
        While count <> arrayHolidays.Length - 1
            Using connection As New SqlConnection("Data Source=PAT;Initial Catalog=FWFCScheduler;Integrated Security=True")
                connection.Open()
                Dim myCommand As New SqlCommand("INSERT INTO SpecialDates (Date, DateType) VALUES(@Date, @DateType)")
                myCommand.Connection = connection
                myCommand.Parameters.AddWithValue("@Date", arrayHolidays(count))
                myCommand.Parameters.AddWithValue("@DateType", "H")
                myCommand.ExecuteNonQuery()
                connection.Close()
                count += 1
            End Using
        End While
    End Sub
    Public Function findDynamicHolidays(arrayHolidays As DateTime(), startValue As DateTime, endValue As DateTime) As DateTime()
        Dim count As Integer = 0
        While startValue <= endValue
            Select Case startValue.ToString("MMMM")
                Case "February"
                    findFamilyDay(arrayHolidays, startValue, count)
                    startValue = startValue.AddDays(28)
                    Exit Select
                Case "May"
                    findVictoriaDay(arrayHolidays, startValue, count)
                    startValue = startValue.AddDays(30)
                    Exit Select
                Case "August"
                    findCivicHolidays(arrayHolidays, startValue, count)
                    startValue = startValue.AddDays(30)
                    Exit Select
                Case "September"
                    findLabourDay(arrayHolidays, startValue, count)
                    startValue = startValue.AddDays(29)
                    Exit Select
                Case "October"
                    findThanksGiving(arrayHolidays, startValue, count)
                    startValue = startValue.AddDays(31)
                    Exit Select
                Case Else
                    Exit Select
            End Select
            startValue = startValue.AddDays(1)
        End While
        Return arrayHolidays
    End Function
    Public Function findFamilyDay(arrayHolidays As DateTime(), startValue As DateTime, count As Integer) As DateTime()
        While count <> 3
            If startValue.ToString("dddd") = "Monday" Then
                count += 1
            End If
            If count = 3 Then
                arrayHolidays(3) = startValue
            End If

            startValue = startValue.AddDays(1)
        End While
        Return arrayHolidays
    End Function
    Public Function findVictoriaDay(arrayHolidays As DateTime(), startValue As DateTime, count As Integer) As DateTime()
        startValue = startValue.AddDays(17)
        While startValue.ToString("dddd") <> "Monday"
            startValue = startValue.AddDays(1)
        End While
        arrayHolidays(4) = startValue
        Return arrayHolidays
    End Function
    Public Function findCivicHolidays(arrayHolidays As DateTime(), startValue As DateTime, count As Integer) As DateTime()
        While count <> 1
            If startValue.ToString("dddd") = "Monday" Then
                count += 1
            End If
            If count = 1 Then
                arrayHolidays(5) = startValue
            End If

            startValue = startValue.AddDays(1)
        End While
        Return arrayHolidays
    End Function
    Public Function findLabourDay(arrayHolidays As DateTime(), startValue As DateTime, count As Integer) As DateTime()
        While count <> 1
            If startValue.ToString("dddd") = "Monday" Then
                count += 1
            End If
            If count = 1 Then
                arrayHolidays(6) = startValue
            End If

            startValue = startValue.AddDays(1)
        End While
        Return arrayHolidays
    End Function
    Public Function findThanksGiving(arrayHolidays As DateTime(), startValue As DateTime, count As Integer) As DateTime()
        While count <> 2
            If startValue.ToString("dddd") = "Monday" Then
                count += 1
            End If
            If count = 2 Then
                arrayHolidays(7) = startValue
                Exit While
            End If

            startValue = startValue.AddDays(1)
        End While
        Return arrayHolidays
    End Function

    Public Function findEaster(y As Integer, arrayHolidays As DateTime(), startValue As DateTime, endValue As DateTime) As DateTime()
        'Algorithm found at http://aa.usno.navy.mil/faq/docs/easter.php
        Dim c As Integer, n As Integer, k As Integer, i As Integer, j As Integer, l As Integer, _
            m As Integer, d As Integer
        c = y / 100
        n = y - 19 * (y / 19)
        k = (c - 17) / 25
        i = c - c / 4 - (c - k) / 3 + 19 * n + 15
        i = i - 30 * (i / 30)
        i = i - (i / 28) * (1 - (i / 28) * (29 / (i + 1)) * ((21 - n) / 11))
        j = y + y / 4 + i + 2 - c + c / 4
        j = j - 7 * (j / 7)
        l = i - j
        m = 3 + (l + 40) / 44
        d = l + 28 - 31 * (m / 4)
        Console.WriteLine("Easter" + m + " " + d)
        Return arrayHolidays
    End Function
    Public Function findStaticHolidays(y As Integer, arrayHolidays As DateTime()) As DateTime()
        'add Canada day to the list
        Dim staticHolidays As New DateTime(y, 7, 1)
        arrayHolidays(0) = staticHolidays

        'add Christmas eve
        staticHolidays = staticHolidays.AddMonths(5)
        staticHolidays = staticHolidays.AddDays(23)
        arrayHolidays(1) = staticHolidays

        'add christmas day
        staticHolidays = staticHolidays.AddDays(1)
        arrayHolidays(2) = staticHolidays
        Return arrayHolidays
    End Function
    Public Sub findWeekends(y As Integer)
        Dim startValue As New DateTime(y, 1, 1)
        Dim endValue As DateTime = startValue.AddMonths(12)
        Dim arrayWeekends As DateTime() = New DateTime(159) {}
        Dim count As Integer = 0
        While startValue < endValue
            Select Case startValue.ToString("dddd")
                Case "Friday"
                    arrayWeekends(count) = startValue
                    count += 1
                    Exit Select
                Case "Saturday"
                    arrayWeekends(count) = startValue
                    count += 1
                    Exit Select
                Case "Sunday"
                    arrayWeekends(count) = startValue
                    count += 1
                    Exit Select
                Case Else
                    Exit Select
            End Select
            startValue = startValue.AddDays(1)
        End While
        count = 0
        While count <> arrayWeekends.Length - 1 OrElse arrayWeekends(count) <> Convert.ToDateTime("1/1/0001")
            Using connection As New SqlConnection("Data Source=PAT;Initial Catalog=FWFCScheduler;Integrated Security=True")
                connection.Open()
                Dim myCommand As New SqlCommand("INSERT INTO SpecialDates (Date, DateType) VALUES(@Date, @DateType)")
                myCommand.Connection = connection
                myCommand.Parameters.AddWithValue("@Date", arrayWeekends(count))
                myCommand.Parameters.AddWithValue("@DateType", "W")
                myCommand.ExecuteNonQuery()
                connection.Close()
                count += 1
            End Using
        End While
    End Sub
End Module


