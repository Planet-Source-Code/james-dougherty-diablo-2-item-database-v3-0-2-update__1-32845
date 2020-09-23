VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   0  'None
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5235
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSetup.frx":000C
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCloseMM2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmSetup.frx":1095
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picCloseUn2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      Picture         =   "frmSetup.frx":137F
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdClose2 
      DownPicture     =   "frmSetup.frx":1676
      Height          =   210
      Left            =   4800
      Picture         =   "frmSetup.frx":19DA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   210
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   3135
   End
   Begin prjChameleon.chameleonButton cmdFunc 
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmSetup.frx":1CD1
   End
   Begin prjChameleon.chameleonButton cmdFunc 
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   1680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   13160664
      FCOL            =   0
      FCOLO           =   0
      MPTR            =   0
      MICON           =   "frmSetup.frx":1CED
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Create New Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   960
      TabIndex        =   3
      Top             =   315
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   840
      Picture         =   "frmSetup.frx":1D09
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   1230
   End
   Begin VB.Image imgAcc 
      Height          =   480
      Left            =   240
      Picture         =   "frmSetup.frx":26EE
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose2_Click()
Call cmdFunc_Click(1)
End Sub

Private Sub cmdClose2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose2.Picture = picCloseMM2.Picture
End Sub

Private Sub cmdFunc_Click(Index As Integer)
On Local Error Resume Next
Select Case Index
 Case 0
  If txtName = "" Then
   frmMSG.lblMessage.Caption = "Please enter your account's name."
   frmMSG.Show 1
   OK = False
   Exit Sub
  End If
  AName = txtName.Text
  OK = True
  Unload Me
 Case 1
  OK = False
  Unload Me
End Select
End Sub

Private Sub Form_Load()
FloatWindow frmSetup.hWnd, Float
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose2.Picture = picCloseUn2.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmSetup.hWnd, SINK
End Sub
