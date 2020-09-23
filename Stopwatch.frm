VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HS StopWatch"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   1080
   End
   Begin VB.CommandButton cmdElapsed 
      Caption         =   "Elapsed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtTmrDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'High Speed Timer class with 3 sample programs to illustrate usage!
'
'  1)  Time execution of your code.
'  2)  Simple Game to test your reaction time.
'  3)  Fully functional high-precision stopwatch.
'
'The High Speed Timer class is fully independent and ready to be
'added to other proggies.  Also shows usage of LARGE INTEGER (8 bytes),
'demonstrates a few API calls, all wrapped up in a neat little package.
'Please vote if you think it's worth it :<|
'
'Written by JR Musselman
'   jrmintn@yahoo.com
'
Option Explicit

'make a copy of the timer.
Dim tmHowLong As clsElapsedTimer

Private Sub cmdElapsed_Click()

    tmrUpdate.Enabled = False
    TimerShowCurrentValue
    
End Sub

Private Sub cmdReset_Click()

'Reset the timer
tmHowLong.Reset
TimerShowCurrentValue

End Sub

Private Sub cmdGo_Click()

'Activate the timer
tmHowLong.Start

'Turn on the automatic update of display
tmrUpdate.Enabled = True

End Sub

Private Sub cmdStop_Click()

'Stop the timer
tmHowLong.StopTimer

'Turn off the automatic display
tmrUpdate.Enabled = False

'Update the display with latest time
TimerShowCurrentValue

End Sub

Private Sub Form_Load()

'make a local copy of the timer.
Set tmHowLong = New clsElapsedTimer
tmHowLong.Reset
TimerShowCurrentValue

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Release resources
Set tmHowLong = Nothing

End Sub

Private Sub TimerShowCurrentValue()
txtTmrDisplay.Text = Format(tmHowLong.Elapsed, "####0.000000")
End Sub

Private Sub tmrUpdate_Timer()
TimerShowCurrentValue
End Sub
