VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rotate WMF´s     eaguirre"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   9
      Left            =   6600
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1395
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Index           =   8
      Left            =   1800
      Picture         =   "Form1.frx":03E2
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Index           =   7
      Left            =   240
      Picture         =   "Form1.frx":18C4
      ScaleHeight     =   1635
      ScaleWidth      =   1395
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Index           =   6
      Left            =   3360
      Picture         =   "Form1.frx":2586
      ScaleHeight     =   3195
      ScaleWidth      =   2955
      TabIndex        =   8
      Top             =   1800
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   5
      Left            =   1800
      Picture         =   "Form1.frx":2AC8
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   4
      Left            =   240
      Picture         =   "Form1.frx":540A
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   3
      Left            =   4920
      Picture         =   "Form1.frx":5C4C
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   2
      Left            =   3360
      Picture         =   "Form1.frx":65EE
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   1
      Left            =   1800
      Picture         =   "Form1.frx":7110
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin RotateWMF.angButton angButton1 
      Height          =   1200
      Left            =   6600
      TabIndex        =   1
      Top             =   600
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
      Trace           =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":8332
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--------------------------------------------------------------
' Author: EAguirre
' Date  : 5/01/98
' Prog  : Form1.frm
' Desc  : Rotate a Windows Metafile
'--------------------------------------------------------------

Option Explicit
Dim fDraw As Boolean

Private Sub Command1_Click()
  fDraw = Not (fDraw)
  If fDraw Then
    Command1.Caption = "Stop"
   Else
    Command1.Caption = "Start"
  End If
  If fDraw Then Draw
End Sub

Private Sub Form_Load()
fDraw = False
End Sub

Public Sub Draw()
Dim i As Integer
Do While fDraw
    Angle = angButton1.Angle
    'Rotate the metafile
    For i = 0 To 9
      EnumMetaFile Picture1(i).hdc, Picture1(i).Picture, AddressOf EnumMetaRecord, 1
      Picture1(i).Refresh
    Next i
    DoEvents
Loop
End Sub

