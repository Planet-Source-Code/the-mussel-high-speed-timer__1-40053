VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HS Timer Game"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
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

'make a new copy of the timer.
Dim tmHowLong As clsElapsedTimer

Private Sub Command1_Click()

'Reset the HS counter
Text1 = "Hit the STOP button, fast...."
tmHowLong.Reset
tmHowLong.Start

End Sub

Private Sub Command2_Click()

Dim MsElapsed As Double

MsElapsed = tmHowLong.Elapsed * 1000
tmHowLong.Reset

If MsElapsed > 0 Then
    'Update the display with elapsed time
    Text1.Text = "It took you " & Format(MsElapsed, "######.###### Milliseconds")
Else
    Text1 = "Hit the start button first."
End If

End Sub

Private Sub Form_Load()

'make a local copy of the timer.
Set tmHowLong = New clsElapsedTimer

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Release resources
Set tmHowLong = Nothing

End Sub
