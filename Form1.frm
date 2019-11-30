VERSION 5.00
Object = "{BD200638-829B-11D2-8FF6-F0DA46C10001}#1.0#0"; "NTPORT.OCX"
Begin VB.Form Form1 
   Caption         =   "Arranque Estrella Triangulo"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin NTPORT_Custom_Control.NTPORT NTPORT1 
      Left            =   5400
      Top             =   2760
      _ExtentX        =   1085
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Automatico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Arranque Estrella-Triangulo"
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
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form2.Enabled = True
Form2.Visible = True
Form3.Visible = False
End Sub

Private Sub Command2_Click()
Form3.Enabled = True
Form3.Visible = True
Form2.Visible = False
End Sub

Private Sub Form_Load()
NTPORT1.address = &H378
NTPORT1.Value = &H0
End Sub

