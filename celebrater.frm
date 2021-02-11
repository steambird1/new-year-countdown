VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Celebrater by steambird"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   ScaleHeight     =   3405
   ScaleWidth      =   5400
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Happy new year!!"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path & "\celebrate.jpeg")
End Sub

Private Sub Form_Resize()
Image1.Move 0, 0, Me.Width, Me.Height
End Sub

Private Sub Form_Terminate()
Unload Form1
End Sub

