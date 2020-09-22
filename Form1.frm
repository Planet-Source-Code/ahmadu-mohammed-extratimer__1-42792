VERSION 5.00
Object = "*\AExtraTimerCtl.vbp"
Begin VB.Form Form1 
   Caption         =   "ET"
   ClientHeight    =   528
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   1728
   LinkTopic       =   "Form1"
   ScaleHeight     =   528
   ScaleWidth      =   1728
   StartUpPosition =   3  'Windows Default
   Begin Project1.ExtraTimer ExtraTimer1 
      Left            =   480
      Top             =   1440
      _ExtentX        =   529
      _ExtentY        =   529
      Interval        =   25
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Testing ExtraTimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1692
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ExtraTimer1_Timer()
If Label1.Visible = True Then Label1.Visible = False Else Label1.Visible = True

End Sub

Private Sub Form_Load()
Form_Resize
ExtraTimer1.Enabled = True
End Sub

Private Sub Form_Resize()
Label1.Move 0, 0
Form1.Width = Label1.Width
Form1.Height = Label1.Height
End Sub
