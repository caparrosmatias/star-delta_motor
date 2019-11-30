VERSION 5.00
Object = "{BD200638-829B-11D2-8FF6-F0DA46C10001}#1.0#0"; "NTPORT.OCX"
Begin VB.Form Form2 
   Caption         =   "Arranque Automatico"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Enabled         =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   3120
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3120
      Top             =   2760
   End
   Begin NTPORT_Custom_Control.NTPORT NTPORT1 
      Left            =   3840
      Top             =   2760
      _ExtentX        =   1508
      _ExtentY        =   873
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   2760
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apagado"
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encendido"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Apagado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1395
      Left            =   240
      Picture         =   "Form2.frx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1395
      Left            =   240
      Picture         =   "Form2.frx":0DBB
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1395
      Left            =   240
      Picture         =   "Form2.frx":1B1B
      Top             =   1320
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Timer1.Enabled = True
NTPORT1.address = &H378
NTPORT1.Value = &H1
Image1.Visible = False

Image2.Visible = True

Image3.Visible = False
Label1.Caption = "Estrella"

Command1.Enabled = False

End Sub

Private Sub Command2_Click()
NTPORT1.address = &H378
NTPORT1.Value = &H0
Image1.Visible = True

Image2.Visible = False

Image3.Visible = False

Label1.Caption = "Apagado"
Command1.Enabled = True


End Sub

Private Sub Form_Load()
Timer1.Enabled = False

Timer2.Enabled = False


End Sub

Private Sub Timer1_Timer()
NTPORT1.address = &H378
NTPORT1.Value = &H3
Timer3.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = False




End Sub

Private Sub Timer2_Timer()
NTPORT1.address = &H378
NTPORT1.Value = &H6
Image1.Visible = False

Image2.Visible = False

Image3.Visible = True
Label1.Caption = "Triangulo"
Timer2.Enabled = False
Timer3.Enabled = False





End Sub

Private Sub Timer3_Timer()
Timer1.Enabled = False

Timer2.Enabled = True

Timer3.Enabled = False

NTPORT1.address = &H378
NTPORT1.Value = &H2


End Sub
