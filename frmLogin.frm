VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   840
      Picture         =   "frmLogin.frx":000C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   10
      Top             =   1230
      Width           =   300
   End
   Begin VB.CommandButton cmdClose2 
      DownPicture     =   "frmLogin.frx":0341
      Height          =   210
      Left            =   4800
      Picture         =   "frmLogin.frx":06A5
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   210
   End
   Begin VB.PictureBox picCloseMM2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "frmLogin.frx":099C
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   8
      Top             =   2880
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
      Left            =   120
      Picture         =   "frmLogin.frx":0C86
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picChecked 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      Picture         =   "frmLogin.frx":0F7D
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picPlain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      Picture         =   "frmLogin.frx":1300
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSMask.MaskEdBox txtPassword 
      Height          =   255
      Left            =   2025
      TabIndex        =   0
      Top             =   720
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   450
      _Version        =   393216
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   2880
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   2025
      TabIndex        =   3
      Top             =   375
      Width           =   2550
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   1725
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      MICON           =   "frmLogin.frx":1635
   End
   Begin prjChameleon.chameleonButton cmdOk 
      Height          =   375
      Left            =   1110
      TabIndex        =   12
      Top             =   1725
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
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
      MICON           =   "frmLogin.frx":1651
   End
   Begin VB.Image imgUnlocked 
      Height          =   480
      Left            =   240
      Picture         =   "frmLogin.frx":166D
      Top             =   240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLocked 
      Height          =   480
      Left            =   240
      Picture         =   "frmLogin.frx":1AAF
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remember Password On This Computer"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   1260
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   825
      TabIndex        =   1
      Tag             =   "&Password:"
      Top             =   780
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   825
      TabIndex        =   2
      Tag             =   "&User Name:"
      Top             =   390
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   2385
      Left            =   0
      Picture         =   "frmLogin.frx":1EF1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5325
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Private Checked As Boolean
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Private Sub cmdClose2_Click()
Call cmdCancel_Click
End Sub

Private Sub cmdClose2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose2.Picture = picCloseMM2.Picture
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim sBuffer As String
Dim lSize As Long
Dim Check As String

sBuffer = Space$(255)
lSize = Len(sBuffer)
Call GetUserName(sBuffer, lSize)

If lSize > 0 Then
 txtUserName.Text = Left$(sBuffer, lSize)
Else
 txtUserName.Text = vbNullString
End If

Check = GetSetting(App.Title, "Settings", "Remember Value", False)
Checked = Check

If Checked = True Then
 picCheck.Picture = picChecked.Picture
 txtPassword.Text = GetSetting(App.Title, "Settings", "Remember Password", "")
Else
 picCheck.Picture = picPlain.Picture
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose2.Picture = picCloseUn2.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next

If Checked Then
 SaveSetting App.Title, "Settings", "Remember Value", Checked
 SaveSetting App.Title, "Settings", "Remember Password", txtPassword.Text
Else
 SaveSetting App.Title, "Settings", "Remember Value", Checked
 SaveSetting App.Title, "Settings", "Remember Password", " "
End If

End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
OK = False
Me.Hide
End Sub

Private Sub cmdOk_Click()
On Local Error Resume Next
If txtPassword.Text = "KSMD" Or txtPassword.Text = "ksmd" Then
 OK = True
 Me.Hide
Else
 frmMSG.lblMessage.Caption = "Invalid password. " & vbCrLf & "Please insert the correct password and try agian"
 frmMSG.Show 1
 txtPassword.SetFocus
 txtPassword.SelStart = 0
 txtPassword.SelLength = Len(txtPassword.Text)
End If
End Sub

Private Sub picCheck_Click()
If Checked = False Then
 picCheck.Picture = picChecked.Picture
 Checked = True
Else
 picCheck.Picture = picPlain.Picture
 Checked = False
End If
End Sub

Private Sub Timer1_Timer()
On Local Error Resume Next
If txtPassword.Text = "KSMD" Or txtPassword.Text = "ksmd" Then
 imgLocked.Visible = False
 imgUnlocked.Visible = True
 Me.Icon = imgUnlocked.Picture
Else
 imgLocked.Visible = True
 imgUnlocked.Visible = False
 Me.Icon = imgLocked.Picture
End If
End Sub
