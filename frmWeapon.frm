VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmWeapon 
   BorderStyle     =   0  'None
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   Icon            =   "frmWeapon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWeapon.frx":000C
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   4
      Left            =   3960
      Picture         =   "frmWeapon.frx":16FF
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   19
      Top             =   2040
      Width           =   300
   End
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   2640
      Picture         =   "frmWeapon.frx":1A34
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   14
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   1200
      Picture         =   "frmWeapon.frx":1D69
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   13
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   2640
      Picture         =   "frmWeapon.frx":209E
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   12
      Top             =   2040
      Width           =   300
   End
   Begin VB.PictureBox picCheck 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   1200
      Picture         =   "frmWeapon.frx":23D3
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   2040
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
      Picture         =   "frmWeapon.frx":2708
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picChecked 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      Picture         =   "frmWeapon.frx":2A3D
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picCloseMM2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1680
      Picture         =   "frmWeapon.frx":2DC0
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
      Left            =   1920
      Picture         =   "frmWeapon.frx":30AA
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdClose2 
      DownPicture     =   "frmWeapon.frx":33A1
      Height          =   210
      Left            =   5760
      Picture         =   "frmWeapon.frx":3705
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   210
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   4410
   End
   Begin prjChameleon.chameleonButton cmdFunc 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   3240
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
      MICON           =   "frmWeapon.frx":39FC
   End
   Begin prjChameleon.chameleonButton cmdFunc 
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
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
      MICON           =   "frmWeapon.frx":3A18
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crafted"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   20
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unique"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   18
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magical"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3120
      TabIndex        =   17
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rare"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1680
      TabIndex        =   16
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Regular"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1680
      TabIndex        =   15
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   " Weapon Type"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   210
      Left            =   1170
      TabIndex        =   8
      Top             =   1650
      Width           =   1920
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Add Weapon"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   435
      Width           =   4140
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   960
      Picture         =   "frmWeapon.frx":3A34
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4410
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   960
      Picture         =   "frmWeapon.frx":4419
      Stretch         =   -1  'True
      Top             =   1485
      Width           =   4410
   End
   Begin VB.Image imgChar 
      Height          =   480
      Left            =   315
      Picture         =   "frmWeapon.frx":53CA
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   510
   End
End
Attribute VB_Name = "frmWeapon"
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
 
  If txtName.Text = "" Then
   frmMSG.lblMessage.Caption = "Please enter a name for the weapon."
   frmMSG.Show 1
   Exit Sub
  End If
  
  If lblTemp.Caption = "" Then
   frmMSG.lblMessage.Caption = "Please select the type of weapon."
   frmMSG.Show 1
   Exit Sub
  End If
  
  If picCheck(0).Picture = picChecked.Picture Then
   HoldOPValue3 = 0
  ElseIf picCheck(1).Picture = picChecked.Picture Then
   HoldOPValue3 = 1
  ElseIf picCheck(2).Picture = picChecked.Picture Then
   HoldOPValue3 = 2
  ElseIf picCheck(3).Picture = picChecked.Picture Then
   HoldOPValue3 = 3
  ElseIf picCheck(4).Picture = picChecked.Picture Then
   HoldOPValue3 = 4
  End If
  
  WeaponName = txtName.Text
  WeaponOK = True
  Unload Me
 Case 1
  WeaponOK = False
  Unload Me
End Select

End Sub

Private Sub Form_Load()
FloatWindow frmWeapon.hWnd, Float
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdClose2.Picture = picCloseUn2.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmWeapon.hWnd, SINK
End Sub

Private Sub picCheck_Click(Index As Integer)
lblTemp.Caption = "OK"
Select Case Index
 Case 0
  picCheck(0).Picture = picChecked.Picture
  picCheck(1).Picture = picPlain.Picture
  picCheck(2).Picture = picPlain.Picture
  picCheck(3).Picture = picPlain.Picture
  picCheck(4).Picture = picPlain.Picture
 Case 1
  picCheck(0).Picture = picPlain.Picture
  picCheck(1).Picture = picChecked.Picture
  picCheck(2).Picture = picPlain.Picture
  picCheck(3).Picture = picPlain.Picture
  picCheck(4).Picture = picPlain.Picture
 Case 2
  picCheck(0).Picture = picPlain.Picture
  picCheck(1).Picture = picPlain.Picture
  picCheck(2).Picture = picChecked.Picture
  picCheck(3).Picture = picPlain.Picture
  picCheck(4).Picture = picPlain.Picture
 Case 3
  picCheck(0).Picture = picPlain.Picture
  picCheck(1).Picture = picPlain.Picture
  picCheck(2).Picture = picPlain.Picture
  picCheck(3).Picture = picChecked.Picture
  picCheck(4).Picture = picPlain.Picture
 Case 4
  picCheck(0).Picture = picPlain.Picture
  picCheck(1).Picture = picPlain.Picture
  picCheck(2).Picture = picPlain.Picture
  picCheck(3).Picture = picPlain.Picture
  picCheck(4).Picture = picChecked.Picture
End Select
End Sub
