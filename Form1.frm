VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UTC Time / Date Example"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "ISO 8601 Format"
      Height          =   855
      Left            =   120
      TabIndex        =   35
      Top             =   3480
      Width           =   9015
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   42
         Text            =   "Text18"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text19 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3960
         TabIndex        =   38
         Text            =   "Text19"
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Text            =   "Text17"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Date / Time"
         Height          =   255
         Left            =   3960
         TabIndex        =   41
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Time"
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Now Function"
      Height          =   1095
      Left            =   4680
      TabIndex        =   31
      Top             =   2280
      Width           =   4455
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   43
         Text            =   "Text16"
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Text            =   "Text15"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "UTC ""Now"""
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "VB ""Now"""
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Medium Date"
      Height          =   975
      Left            =   4680
      TabIndex        =   26
      Top             =   1200
      Width           =   4455
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Text            =   "Text14"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Text            =   "Text13"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Local Format"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "UTC Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Local Date"
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   25
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Long Date"
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   135
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text10"
         Top             =   975
         Width           =   4215
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Text9"
         Top             =   450
         Width           =   4215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "UTC Long Date"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   765
         Width           =   4215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Local Long Date"
         Height          =   270
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Time Zone Information"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   9015
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Text12"
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Text11"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Long Time Zone"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Short Time Zone"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Short Date"
      Height          =   975
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Text7"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Text5"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Local Format"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "UTC Date"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Local Date"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   1920
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      Caption         =   "Time"
      Height          =   1830
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3015
         TabIndex        =   46
         Text            =   "Text8"
         Top             =   1365
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1320
         TabIndex        =   45
         Text            =   "Text6"
         Top             =   1380
         Width           =   1350
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Text4"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Text2"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "or"
         Height          =   270
         Left            =   2745
         TabIndex        =   47
         Top             =   1395
         Width           =   210
      End
      Begin VB.Label Label20 
         Caption         =   "UTC Time Offset"
         Height          =   255
         Left            =   75
         TabIndex        =   44
         Top             =   1425
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "UTC Time"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Local Time"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "24 Hour Format"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Local Format"
         Height          =   255
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Menu UTCexit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu UTCabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




    '****************************************************************************
    '   Demo project for UTC time functions in module UTC_time
    '
    '   I need UTC time a lot, since I write a lot of Ham Radio programs.
    '   But there are a lot of uses for UTC (GMT), Time stamps etc.
    '   Feel free to use and change the code.  You can contact me at
    '   the email address below.
    '
    '   Module UTC_time is set of functions to return UTC Time / Date using some API calls
    '   Format of functions similar to standard Time / Date
    '   funtions in VB (Time, Time$, Date, Date$)
    '
    '   Also there are two funtions that return the name of the
    '   time zone the system is set for, one the short name (like "EST"),
    '   the second for the long name (like "Eastern Standard Time")

    '   Time zone and UTC time functions assume the correct time zone is selected
    '   in the Time/Date properties.
    '
    '   Mark Mokoski, 03-JAN-2003
    '   markm@cmtelephone.com
    '
    '***************************************************************************

Private Sub Command1_Click()

    Unload frmAbout
    Unload Me
    'End program

End Sub

Private Sub Form_Load()

    ' ***************************************************************************
    ' * Test to see if App is allready running
    ' * If App is running, terminate copy
    ' ***************************************************************************

        If App.PrevInstance Then
            MsgBox "UTC Time / Date application is already running." & vbCrLf & _
            "Only one instance (copy) of program this can be running" & vbCrLf & _
            "for proper operation.", vbCritical, "Application ERROR"
            End
        Else
            '  MsgBox "This is the first instance of your application."
        End If

    '"Kick off" the form with a timer event to populate text boxes
    Timer1_Timer

End Sub

Private Sub Timer1_Timer()

    '******************************
    '
    'See UTC_time module for full discription of functions
    '
    '******************************
    'Display times
    Text1.Text = Time       'Local format
    Text2.Text = Time$      '24 Hour format
    Text3.Text = UTCtime2   'Local format
    Text4.Text = UTCtime    '24 Hour format
    'Display dates
    Text5.Text = Date       'Local Format(region dependent)
    Text6.Text = UTCoffset & " Minutes" 'UTC to Local Time Offset in Minutes
    Text7.Text = UTCdate   'Like date function  (region dependent)
    Text8.Text = (UTCoffset / 60) & " Hours" 'UTC Offset in Hours
    Text9.Text = Format(Date, "long date")
    Text10.Text = Format(UTCdate, "long date") 'UTCdate can be used like date in formating
    Text13.Text = UCase(Format(Date, "medium date"))     'I just like the look of uppercase here
    Text14.Text = UCase(Format(UTCdate, "medium date"))  'I just like the look of uppercase here
    'Display Time Zone Info
    Text11.Text = shortTZname   'ex: "EST"
    Text12.Text = longTZname    'ex: "Eastern Standard Time"
    'Display "Now" function
    Text15.Text = Now 'VB "Now" function
    Text16.Text = UTCnow 'UTC "Now" function
    'ISO 8601 Date/Time format
    Text17.Text = ISOdate
    Text18.Text = ISOtime
    Text19.Text = ISOnow

End Sub

Private Sub UTCabout_Click()

        If frmAbout.Visible = True Then
            frmAbout.SetFocus    'Don't load second copy of form, set focus on form
            Exit Sub
        Else
            Load frmAbout
            frmAbout.Visible = True
        End If

End Sub

Private Sub UTCexit_Click()

    Command1_Click 'End App

End Sub
