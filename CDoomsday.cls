VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDoomsday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iDay As Integer           'day value
Private m_iCentury As Integer       'century value
Private m_iTwelves As Integer       'twelves value
Private m_iRemainder As Integer     'remainder value
Private m_iFours As Integer         'fours value
Private m_sDayOfWeek As String      'day of week the date falls on

Private m_mMonth As Months
Private m_sMonth As String          'value of test month
Private m_iDy As Integer            'value of test day
Private m_sYr As String             'value of test year
Private m_iWk As Integer            'value of test week

Private m_bLeapYear As Boolean      'leap year flag

Private m_arZeroDays(11) As Integer    'holds the zero day for each month
Private m_arDay(6) As String           'holds the days of the week(Sun=0, Sat=6)
Private m_arCent(3) As Integer         'holds centuries 17, 18, 19, 20
Private m_arDate() As String           'holds the date being evaluated


'Months enumeration
Public Enum Months
    January = 0
    Jan = 0
    February = 1
    Feb = 1
    March = 2
    Mar = 2
    April = 3
    Apr = 3
    May = 4
    June = 5
    Jun = 5
    July = 6
    Jul = 6
    August = 7
    Aug = 7
    September = 8
    Sept = 8
    Sep = 8
    October = 9
    Oct = 9
    November = 10
    Nov = 10
    December = 11
    Dec = 11
End Enum


'Days enumeration
Public Enum Days
    Sunday = 0
    Sun = 0
    Monday = 1
    Mon = 1
    Tuesday = 2
    Tues = 2
    Tue = 2
    Wednesday = 3
    Wed = 3
    Thursday = 4
    Thurs = 4
    Thur = 4
    Friday = 5
    Fri = 5
    Saturday = 6
    Sat = 6
End Enum

Public Property Get Day() As Integer
Attribute Day.VB_Description = "Return the Day value"
    Day = m_iDay
End Property

Private Sub CalcDay()
'Calculate the day value
'The day value is the difference between the target date and the
'nearest zero date.

    Dim iTemp As Integer
    If IsLeapYear() And (m_sMonth = "Jan" Or m_sMonth = "Feb") Then
        iTemp = m_arZeroDays(m_mMonth) + 1
    Else
        iTemp = m_arZeroDays(m_mMonth)
    End If
    
    m_iDay = m_iDy - iTemp
End Sub

Public Property Get Century() As Integer
Attribute Century.VB_Description = "Return the century value"
    Century = m_iCentury
End Property

Private Sub CalcCentury()
'Calculate the century value. The century value repeats itself every 400 years
'To find the value we take the 2 digit century value, subtract 17 and mod the
'result by 4.
' For example, if the year is 2410, the century is 24.
' 24 - 17 = 7. 7 Mod 4 = 3. Therefore, the century value is 2
    Dim iCent As Integer
    
    iCent = CInt(Left$(m_sYr, Len(m_sYr) - 2))
    m_iCentury = m_arCent(Abs((iCent - 17) Mod 4))
End Sub

Public Property Get Twelves() As Integer
Attribute Twelves.VB_Description = "Return the twelves value"
    Twelves = m_iTwelves
End Property

Private Sub CalcTwelves()
'Calculate the Twelve's value. the Twelves value is the number of times 12
'will go evenly into the 2digit year. Thus, integer division is required
'example: if the year is 2147, the 2 digit year is 47. The twelves value
'is 3 because 12 will evenly go into 47 3 times
    m_iTwelves = CInt(Right$(m_sYr, 2) \ 12)
End Sub

Public Property Get Remainder() As Integer
Attribute Remainder.VB_Description = "Return the remainders value "
    Remainder = m_iRemainder
End Property

Private Sub CalcRemainder()
'Calculate the remainders value. the remainder value is the remainder after all
'multiples of 12 have been subtracted from the 2 digit year.
'example: if the year is 2147, the 2 digit year is 47. The remainder value
'is 11, because 12 will evenly go into 47 3 times. 3*12 = 36 and 47 - 36 = 11
    m_iRemainder = CInt(Right$(m_sYr, 2) Mod 12)
End Sub

Public Property Get Fours() As Integer
Attribute Fours.VB_Description = "Return the four's value"
    Fours = m_iFours
End Property

Private Sub CalcFours()
'calculate the four's value. The fours value is the number of times 4 will evenly
'go into the remainder value.
'example: if the year is 2147, the fours value is 2 because the remainder value
'is 11 (see above) and 4 will evenly go into 11 2 times
    m_iFours = CInt((Right$(m_sYr, 2) Mod 12) \ 4)
End Sub

Public Property Get LeapYear() As Boolean
    LeapYear = m_bLeapYear
End Property

Public Function GetDayOfWeek(sDate As String) As String
Attribute GetDayOfWeek.VB_Description = "Calculate the day of week the date falls on"
'date passed in the form of "Month Day Year"
'first we extract the individual parts from the date string
'from those values and calculate the day of week
    Erase m_arDate
    m_arDate = Split(sDate, " ")
    m_sMonth = Left$(m_arDate(0), 3)    'get test month
    SetMonth
    m_iDy = CInt(m_arDate(1))           'get test day
    m_sYr = m_arDate(2)                 'get test year
    
    CalcDay
    CalcCentury
    CalcTwelves
    CalcRemainder
    CalcFours
    GetDayOfWeek = CalcDayOfWeek
    
End Function

Public Function GetDate(sDay As String) As Integer
Attribute GetDate.VB_Description = "Return the Date the desired day falls on"
'To get the date a particular day falls on we need 4 values:
'the week, the day, the month, and the year
'The date is calculated in 5 steps
'1) get the zero day for the desired week and month
'2) get the zero day for the 2 digit year
'3) get the difference between the year zero day and the test day
'4) add that difference to the weekly zero day
' example: what day does the 3rd tuesday of April, 2001 fall on
'1) The zero day for April is the 4th. Since that always falls in the
'   first week we need to add 2 weeks to find the zero day of the 3rd week:
'   4 + 2*7 = 4 + 14 = 18. So, the zero day of April, week 3 is the 18th
'2) next we need to find the zero day of the year 2001. This is done by
'   summing the century, twelves, remainder, and fours values and modling by 7:
'   century value = 2, twelves = 0, remainder = 1, fours = 0. Summing and
'   modding by 7 = 3. so the zero day for the year 2001 is day 3, or wednesday
'3) next we subtract the target day (tuesday) from the zero day (wednesday)
'   tuesday - wednesday = -1
'4)  finally we add -1 to the zero day of April, week 3.
'    18 + (-1) = 17.
' Therefore the 3rd tuesday in April, 2001 falls on the 17th.


    Dim iZeroDay As Integer     'zero day for desired week of desired month
    Dim iDiff As Integer        'difference b/t desired day and iZeroDay
    Dim iYearZeroDay As Days    'zero day for desired year
    Dim strDay As String        'desired day
    Dim iDate As Integer        'date the desired day falls on
    
    Erase m_arDate
    m_arDate = Split(sDay, " ")
    m_iWk = CInt(Left$(m_arDate(0), 1))
    strDay = m_arDate(1)
    m_sMonth = Left$(m_arDate(2), 3)
    SetMonth
    m_sYr = m_arDate(3)
    
    iZeroDay = m_arZeroDays(m_mMonth) + (7 * (m_iWk - 1))
    iYearZeroDay = CalcYearZeroDay
    m_iDy = DayAsInteger(strDay)
    iDiff = m_iDy - iYearZeroDay
    
    iDate = iDiff + iZeroDay
    If iDate < 1 Then iDate = iDate + 7
    
    If IsLeapYear() Then iDate = iDate + 1
    If CheckDate(iDate) Then
        GetDate = iDate
    Else
        GetDate = -1
    End If
    
End Function

Private Function CheckDate(iDay As Integer) As Boolean
'Checks to see if the date calculated is a valid date.
'This ensures that an invalid week and day are not entered.
'for example, there is no week 5, day 4 in february. this procedure
'ensures that this error is caught

    Select Case m_mMonth
        Case January, March, May, July, August, October, December
            If iDay > 31 Then
                CheckDate = False
                Exit Function
            End If
        Case April, June, September, November
            If iDay > 30 Then
                CheckDate = False
                Exit Function
            End If
        Case February                      'feb
            If m_bLeapYear = True And iDay > 29 Then
                CheckDate = False
                Exit Function
            ElseIf iDay > 28 Then
                CheckDate = False
                Exit Function
            End If
    End Select
    
    CheckDate = True
    
End Function

Private Function DayAsInteger(sDay As String) As Integer
'this procedure returns the day of the week as an integer so that
'it can be subtracted from another day
    Dim iI As Integer
    
    iI = Sunday
    Do While iI < Saturday
        If m_arDay(iI) = sDay Then
            DayAsInteger = iI
            Exit Function
        End If
        iI = iI + 1
    Loop
        
End Function

Private Function CalcYearZeroDay() As Days
'as discussed above, the zero day for any given year is:
' (century value + twelves value + remainder value + fours value) mod 7
    CalcCentury
    CalcTwelves
    CalcRemainder
    CalcFours
    
    CalcYearZeroDay = (m_iCentury + m_iTwelves + m_iRemainder + m_iFours) Mod 7

End Function

Private Function CalcDayOfWeek() As String
'Calculate the day of week the date falls on. The day of week corresponds
'to the sum of the day, century, twelves, remainder, and days values mod 7.
    Dim iTotal As Integer
    
    'total the 5 calculated values
    iTotal = m_iDay + m_iCentury + m_iTwelves + m_iRemainder + m_iFours
    
    'if the total is negative then add 7
    If iTotal < 0 Then
        iTotal = iTotal + 1
    Else
        'else get the remainder of the total divided by 7
        iTotal = iTotal Mod 7
    End If
    
    m_sDayOfWeek = m_arDay(iTotal)
    CalcDayOfWeek = m_sDayOfWeek
    
End Function

Public Property Get DayOfWeek() As String
Attribute DayOfWeek.VB_Description = "Return the day of week the date falls on"
    DayOfWeek = m_sDayOfWeek
End Property

Private Function IsLeapYear() As Boolean
'Calculates whether a given year is a leap year or not
'Returns True if it is a leap year, false if it is not
'A century year (2000, e.g.) is a leap year only if it is
'divisible by 400. Any other year is a leap year if it is
'divisible by 4.

    'check if year is a century year and is divisible by 400
    If ((CInt(m_sYr) Mod 100 = 0) And (CInt(m_sYr) Mod 400 = 0)) Then
        m_bLeapYear = True
    'else check if year is divisible by 4
    ElseIf (CInt(m_sYr) Mod 4 = 0) Then
        m_bLeapYear = True
    Else
        m_bLeapYear = False
    End If
    
    IsLeapYear = m_bLeapYear
    
End Function

Private Sub SetMonth()
    Select Case m_sMonth
        Case "January", "Jan"
            m_mMonth = January
        Case "February", "Feb"
            m_mMonth = February
        Case "March", "Mar"
            m_mMonth = March
        Case "April", "Apr"
            m_mMonth = April
        Case "May"
            m_mMonth = May
        Case "June", "Jun"
            m_mMonth = June
        Case "July", "Jul"
            m_mMonth = July
        Case "August", "Aug"
            m_mMonth = August
        Case "September", "Sept", "Sep"
            m_mMonth = September
        Case "October", "Oct"
            m_mMonth = October
        Case "November", "Nov"
            m_mMonth = November
        Case "December", "Dec"
            m_mMonth = December
    End Select
    
End Sub
Private Sub Class_Initialize()
'********************************************
'fill Zero_Days array with the zero day for the
'corresponding month
    m_arZeroDays(January) = 3
    m_arZeroDays(February) = 0
    m_arZeroDays(March) = 0
    m_arZeroDays(April) = 4
    m_arZeroDays(May) = 2
    m_arZeroDays(June) = 6
    m_arZeroDays(July) = 4
    m_arZeroDays(August) = 1
    m_arZeroDays(September) = 5
    m_arZeroDays(October) = 3
    m_arZeroDays(November) = 0
    m_arZeroDays(December) = 5
'********************************************


'********************************************
'initialize the Day array with the days of
'the week Sun-Sat
    m_arDay(Sunday) = "Sunday"
    m_arDay(Monday) = "Monday"
    m_arDay(Tuesday) = "Tuesday"
    m_arDay(Wednesday) = "Wednesday"
    m_arDay(Thursday) = "Thursday"
    m_arDay(Friday) = "Friday"
    m_arDay(Saturday) = "Saturday"
'********************************************


'********************************************
'initialize Cent array with the doomsday value
'for each century
    m_arCent(0) = 0     '1700's
    m_arCent(1) = 5     '1800's
    m_arCent(2) = 3     '1900's
    m_arCent(3) = 2     '2000's
'********************************************

    m_iDay = 0
    m_iCentury = 0
    m_iTwelves = 0
    m_iRemainder = 0
    m_iFours = 0

End Sub
