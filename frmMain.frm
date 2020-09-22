VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Find date for given week, day, month, year"
      Height          =   2850
      Left            =   4110
      TabIndex        =   11
      Top             =   810
      Width           =   4860
      Begin VB.CommandButton cmdClearDate 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2790
         TabIndex        =   25
         Top             =   2220
         Width           =   1080
      End
      Begin VB.CommandButton cmdGetDate 
         Caption         =   "Get Date"
         Height          =   375
         Left            =   1365
         TabIndex        =   23
         Top             =   2220
         Width           =   1080
      End
      Begin VB.ComboBox cboWkNumber 
         Height          =   315
         Left            =   315
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   780
         Width           =   660
      End
      Begin VB.ComboBox cboWkDay 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   780
         Width           =   990
      End
      Begin VB.ComboBox cboMonthB 
         Height          =   315
         Left            =   2475
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   780
         Width           =   1080
      End
      Begin VB.TextBox txtYearB 
         Height          =   345
         Left            =   3765
         TabIndex        =   13
         Top             =   750
         Width           =   780
      End
      Begin VB.TextBox txtDisplayB 
         Height          =   330
         Left            =   330
         TabIndex        =   12
         Top             =   1590
         Width           =   4200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Week"
         Height          =   195
         Left            =   315
         TabIndex        =   21
         Top             =   525
         Width           =   435
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Day"
         Height          =   195
         Left            =   1275
         TabIndex        =   20
         Top             =   525
         Width           =   285
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         Height          =   195
         Left            =   2580
         TabIndex        =   19
         Top             =   525
         Width           =   450
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Year(yyyy)"
         Height          =   195
         Left            =   3765
         TabIndex        =   18
         Top             =   525
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date for selected week, day, month, year"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1350
         Width           =   2910
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Find Day for given month, date, year"
      Height          =   2850
      Left            =   150
      TabIndex        =   2
      Top             =   810
      Width           =   3750
      Begin VB.CommandButton cmdClearDay 
         Caption         =   "Clear"
         Height          =   375
         Left            =   2085
         TabIndex        =   24
         Top             =   2220
         Width           =   1080
      End
      Begin VB.CommandButton cmdGetDay 
         Caption         =   "Get Day"
         Height          =   375
         Left            =   660
         TabIndex        =   22
         Top             =   2220
         Width           =   1080
      End
      Begin VB.TextBox txtDisplay 
         Height          =   315
         Left            =   360
         TabIndex        =   6
         Top             =   1620
         Width           =   2985
      End
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   2595
         TabIndex        =   5
         Top             =   795
         Width           =   765
      End
      Begin VB.ComboBox cboDay 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   795
         Width           =   720
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   390
         List            =   "frmMain.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   795
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Month:"
         Height          =   195
         Left            =   405
         TabIndex        =   10
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Day:"
         Height          =   195
         Left            =   1740
         TabIndex        =   9
         Top             =   555
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year (yyyy)"
         Height          =   195
         Left            =   2595
         TabIndex        =   8
         Top             =   570
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Day of week for selected date:"
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   1395
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   525
      Left            =   4170
      TabIndex        =   1
      Top             =   3930
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Date Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -120
      TabIndex        =   0
      Top             =   105
      Width           =   9240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DOW As New CDoomsday


'*****************************************************************************
'in 1982 mathematician John Conway developed a very simple algorithm for
'quickly calculating the day of week on which any given date in history falls.
'He named it the doomsday algorithmand it works as follows:

'1) definition of doomsday: Doomsday is the day of the week of the last day
'   in February. example: Feb 29, 2000 is a Tuesday, so Doomsday for the year
'   2000 is Tuesday. Each month has its own corresponding Doomsday date:
'   Jan=17, (18 in leap year); Feb=28 (29 in leap year); Mar=7; April=4; May=9;
'   June=6; July=11; Aug=8; Sept=5; Oct=10; Nov=7; Dec=12.
'   Note: in this program the term zero day is used instead of Doomsday
'   Note: different zero days are used in this program for calculation reasons

'2) 5 number values are calculated: Day, Century, 12's, Remainder, and 4's

' DAY
' Day value is the difference between the day of the month and the nearest
' zero day for the given month.

' CENTURY
' Century value is the zero day for the century the given date lies in.
' 1700 = 0, 1800 = 5, 1900 = 3, 2000 = 2. These repeat every 400 years

' 12's
' 12's value = 2-digit year value \ 12.

' REMAINDER
' Remainder value = 2-digit year value Mod 12.

' 4's
' 4's value is the number of times 4 goes into the Remainder value.


'3) To get the day of week: add the values and Mod by 7. The final value is
'   the day of week where Sun=0, Mon=2......Sat=6

'4) Example: Calculate the day of week for the date August 25th, 1963.
'   Day value: Nearest zero day in August is 22 (2 wks after Aug 8).
'   25 - 22 = 3.
'   Day value = 3.

'   Century value: 1963 is in the 1900's so
'   Century value = 3

'   12's Value: 2-digit year is 63. 12 evenly goes into 63 5 times.
'   12's value = 5.

'   Remainder value: 63 Mod 12 = 3 (63 - 60 = 3).
'   Remainder value = 3.

'   4's value: 4 evenly goes into 3 0 times.
'   4's Value = 0.

'   Add values: 3 + 3 + 5 + 3 + 0 = 14. 14 Mod 7 = 0. 0 corresponds to
'   Sunday, so August 25th, 1963 was a Sunday


'   Example: Calculate the day of week for the date February 14, 2064
'   Day value: nearest zero day is the 15th. 14 - 15 = -1

'   Century value: 2000 = 2.

'   12's value: 64 \ 12 = 5

'   Remainder value: 64 - 60 = 4

'   4's value: 4 even goes into 4 1 time

'   -1 + 2 + 5 + 4 + 1 = 11. 11 Mod 7 = 4. 4 corresponds to Thursday, so
'   February 14, 2064 will be a thursday.

'*****************************************************************************

'We can also use this algorithm to calculate the date for
'a given week, day, month, and year.
'To do this we get the nearest doomsday (zero day) value and calculate the offset.
'Step 1) Get the zero day for the given year.
'Step 2) Get the nearest zero date for the given month.
'Step 3) adjust for the desired day

'Example: What is the 3rd Friday in October, 2015.
'Step 1: zero day for 2015: Cent + 12's + Remainder + 4's = 2 + 1 + 3 + 0 = 6
'        zero day for the year 2015 is a Saturday
'Step 2: Nearest zero date for October is the 17th. The 3rd is the first zero
'        date, the 10th is the 2nd zero date.
'Step 3: Adjust for Friday: since the 17th is the 3rd Saturday, the 16th is the
'        third Friday.
'Therefore, the 3rd Friday in October, 2015 is the 16th.

'*****************************************************************************

Private Sub Load_Values()
'loads values into combo boxes and initializes the arrays
    
'********************************************
'load month names into the month combo boxes
    With cboMonth
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
    End With
    
    With cboMonthB
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
    End With
'********************************************


'********************************************
'load week numbers into the week number combo box
    With cboWkNumber
        .AddItem "1st"
        .AddItem "2nd"
        .AddItem "3rd"
        .AddItem "4th"
        .AddItem "5th"
    End With
'********************************************
    
    
'********************************************
'load the days of the week into the weekday combo box
    With cboWkDay
        .AddItem "Sunday"
        .AddItem "Monday"
        .AddItem "Tuesday"
        .AddItem "Wednesday"
        .AddItem "Thursday"
        .AddItem "Friday"
        .AddItem "Saturday"
    End With
'********************************************

End Sub

Private Sub cboMonth_Click()
'fill cboDay combo box with the correct number of valid days
    Dim intI As Integer
    
    With cboDay
        .Clear  'clear contents of combo box
        Select Case cboMonth.ListIndex
            Case April, June, September, November
                For intI = 1 To 30
                    .AddItem Str(intI)
                Next intI
                
            Case February '(29 is for leap year)
                For intI = 1 To 29
                    .AddItem CStr(intI)
                Next intI
                
            Case January, March, May, July, August, October, December
                For intI = 1 To 31
                    .AddItem CStr(intI)
                Next intI
        End Select
    End With
                
End Sub

Private Sub cmdClearDate_Click()
    txtDisplayB.Text = ""
End Sub

Private Sub cmdClearDay_Click()
    txtDisplay.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGetDate_Click()
    Dim strDate As String
    strDate = Get_Date
End Sub

Private Sub cmdGetDay_Click()
'calculate the day of week the desired day falls on
'format day and display in text box
    Dim strDay As String        'Day of week returned by GetDay function
    Dim strTestDate As String   'date being evaluated
    Dim dCurDate As Date        'current date
    Dim dDate As Date           'strTestDate in Date format
   
    'convert current date to short date format
    dCurDate = Format(Date, "Short Date")
    
    'get test date value
    strTestDate = Trim$(cboMonth.Text) & " " & Trim$(cboDay.Text) & " " & Trim$(txtYear.Text)
    dDate = Format(CDate(strTestDate), "Short Date")
    
    'calculate day of week for date entered by user
    strDay = DOW.GetDayOfWeek(strTestDate)
    
    'format output based on whether desired day is before or after current day
    If (dCurDate < dDate) Then
        txtDisplay.Text = strTestDate & " will be a " & strDay
    Else
        txtDisplay.Text = strTestDate & " was a " & strDay
    End If
End Sub

Private Function Get_Date()
'calculate the date for a given day, week, month, year
    Dim strDate As String
    Dim intTemp As Integer
    
    strDate = Trim$(cboWkNumber.Text) & " " & Trim$(cboWkDay.Text) & " " & _
              Trim$(cboMonthB.Text) & " " & Trim$(txtYearB.Text)
              
    intTemp = DOW.GetDate(strDate)
    
    If intTemp > 0 Then
        strDate = Format_Date(intTemp)
        txtDisplayB.Text = "The " & cboWkNumber.Text & " " & cboWkDay.Text & _
                           " of " & cboMonthB.Text & ", " & _
                           txtYearB.Text & " falls on the " & strDate
    Else
        txtDisplayB.Text = cboMonthB.Text & ", " & txtYearB.Text & _
                           " does not have a " & " " & cboWkNumber.Text & _
                           " " & cboWkDay.Text
    End If
                           
    
End Function

Public Function Format_Date(ByVal intDay As Integer) As String
'format the day

    Select Case intDay
        Case 1, 21, 31
            Format_Date = CStr(intDay) & "st"
            Exit Function
        Case 2, 22
            Format_Date = CStr(intDay) & "nd"
            Exit Function
        Case 3, 23
            Format_Date = CStr(intDay) & "rd"
            Exit Function
        Case Else
            Format_Date = CStr(intDay) & "th"
    End Select
    
End Function

Private Sub Form_Load()
    Call Load_Values
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set DOW = Nothing
End Sub
