VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HS Timer Test"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Test how long it takes to launch Notepad"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
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
'  JR Musselman
'  jrmintn@yahoo.com
'

Option Explicit

Private Sub Command1_Click()

'make a copy of the timer.
Dim tmHowLong As New clsElapsedTimer

'Reset the HS counter
tmHowLong.Start

'Do something you want to time...
Shell "NotePad"

'See how long it took to execute the code.
Text1 = Format(tmHowLong.Elapsed * 1000, "###.000000") & " Milliseconds"

'Release resources
Set tmHowLong = Nothing

End Sub

Private Sub Form_Load()

End Sub
