VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HAPPY NEW YEAR!"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   5325
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Time to LUNAR NEW YEAR:"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
d = DateDiff("s", Now(), #2/12/2021#)
'MsgBox d
End Sub

Private Sub Timer1_Timer()
d = DateDiff("s", Now(), #2/12/2021#)
If d <= 0 Then
    Timer1.Enabled = False
    Form1.Hide
    Form2.Show
    Exit Sub
End If
h = d / 3600
m = (d / 60) Mod 60
s = d Mod 60
Label2.Caption = Format(h, "00") & ":" & Format(m, "00") & ":" & Format(s, "00")
End Sub
