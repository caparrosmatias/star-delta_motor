VERSION 5.00
Object = "{BD200638-829B-11D2-8FF6-F0DA46C10001}#1.0#0"; "NTPORT.OCX"
Begin VB.Form Form3 
   Caption         =   "Arranque Manual"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Enabled         =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2160
      Top             =   2640
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Apagar Todo"
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Contactor Triangulo"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Contactor Estrella"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Contactor Linea"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin NTPORT_Custom_Control.NTPORT NTPORT1 
      Left            =   3960
      Top             =   2520
      _ExtentX        =   1508
      _ExtentY        =   1085
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   1680
      Picture         =   "Form3.frx":0000
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   1680
      Picture         =   "Form3.frx":0C81
      Top             =   1080
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   1680
      Picture         =   "Form3.frx":1A3C
      Top             =   1080
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   1680
      Picture         =   "Form3.frx":279C
      Top             =   1080
      Width           =   2475
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
NTPORT1.address = &H378
NTPORT1.Value = &H2
Command2.Enabled = True
Command1.Enabled = False
Command3.Enabled = False

Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = True


End Sub

Private Sub Command2_Click()
NTPORT1.address = &H378
NTPORT1.Value = &H3
Command2.Enabled = False
Command3.Enabled = True

Image1.Visible = False
Image2.Visible = True
Image3.Visible = False
Image4.Visible = False
End Sub

Private Sub Command3_Click()

Command3.Enabled = False

Image1.Visible = False
Image2.Visible = False
Image3.Visible = True
Image4.Visible = False
Timer1.Enabled = True
NTPORT1.address = &H378
NTPORT1.Value = &H2
End Sub

Private Sub Command4_Click()
NTPORT1.address = &H378
NTPORT1.Value = &H0
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False

Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
End Sub

Private Sub Form_Load()
Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
End Sub

Private Sub Timer1_Timer()
NTPORT1.address = &H378
NTPORT1.Value = &H6
Timer1.Enabled = False

End Sub
