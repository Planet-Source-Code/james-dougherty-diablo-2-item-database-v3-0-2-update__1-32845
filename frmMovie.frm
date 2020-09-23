VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmMovie 
   BorderStyle     =   0  'None
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   Icon            =   "frmMovie.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
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
      Left            =   840
      Picture         =   "frmMovie.frx":000C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   18
      Top             =   2160
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
      Left            =   840
      Picture         =   "frmMovie.frx":0341
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   16
      Top             =   1800
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
      Left            =   840
      Picture         =   "frmMovie.frx":0676
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   14
      Top             =   1440
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
      Left            =   840
      Picture         =   "frmMovie.frx":09AB
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   12
      Top             =   1080
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
      Left            =   840
      Picture         =   "frmMovie.frx":0CE0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      Top             =   720
      Width           =   300
   End
   Begin VB.PictureBox picPlain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1320
      Picture         =   "frmMovie.frx":1015
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   5160
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
      Left            =   240
      Picture         =   "frmMovie.frx":134A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   300
   End
   Begin prjChameleon.chameleonButton cmdExit 
      Height          =   375
      Left            =   1620
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ok/Cancel"
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
      MICON           =   "frmMovie.frx":16CD
   End
   Begin prjChameleon.chameleonButton cmdAbout 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "About"
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
      MICON           =   "frmMovie.frx":16E9
   End
   Begin prjChameleon.chameleonButton cmdApply 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Apply"
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
      MICON           =   "frmMovie.frx":1705
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable one movie"
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
      Left            =   1320
      TabIndex        =   19
      Top             =   735
      Width           =   1485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable two movies"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   1095
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable three movies"
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
      Left            =   1320
      TabIndex        =   15
      Top             =   1455
      Width           =   1725
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable four movies"
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
      Left            =   1320
      TabIndex        =   13
      Top             =   1815
      Width           =   1620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Failing to do so can cause your comp. to hang."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   360
      TabIndex        =   11
      Top             =   4680
      Width           =   3795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Only apply this if you just reinstalled diablo II."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   360
      TabIndex        =   10
      Top             =   4440
      Width           =   3705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Important :"
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
      Left            =   240
      TabIndex        =   9
      Top             =   4200
      Width           =   915
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enable all diablo II movies"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   2175
      Width           =   2160
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Movie Trainer"
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
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   3870
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   135
      Picture         =   "frmMovie.frx":1721
      Stretch         =   -1  'True
      Top             =   225
      Width           =   4110
   End
   Begin VB.Image Image1 
      Height          =   3435
      Left            =   0
      Picture         =   "frmMovie.frx":2106
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4365
   End
End
Attribute VB_Name = "frmMovie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'I redid this a lilttle.
'it checks to make sure they already dont have all movies enabled

Private sMSetting(1 To 5) As String
Private Checked As Boolean

Private Sub cmdAbout_Click()
frmMSG.lblMessage.Caption = "Diablo II Movie Trainer" & vbCrLf & "Written by Max Raskin. So thanks go to him"
frmMSG.Show
End Sub

Private Sub cmdApply_Click()
Dim tmpStr As String
Dim i As Integer

If lblTemp.Caption = "" Then
  frmMSG.lblMessage.Caption = "Please select option to apply."
  frmMSG.Show
  Exit Sub
End If
  
For i = 0 To 4
 If picCheck(i).Picture = picChecked.Picture Then
  UpdateKey HKEY_CURRENT_USER, "Software\Blizzard Entertainment\Diablo II", "Aux Battle.Net", sMSetting(i + 1)
  frmMSG.lblMessage.Caption = lblTemp.Caption + 1 & " movies have been enabled."
  frmMSG.Show
 End If
Next
'tmpStr = RGGetKeyValue(HKEY_CURRENT_USER, "Software\Blizzard Entertainment\Diablo II", "Aux Battle.Net")
'If tmpStr = "216.148.246.50" Then
'  frmMSG.lblMessage.Caption = "All movies are already enabled on this computer."
'  frmMSG.Show
'  Exit Sub
'Else
'  UpdateKey HKEY_CURRENT_USER, "Software\Blizzard Entertainment\Diablo II", "Aux Battle.Net", sMSetting
'End If
    
End Sub

Private Sub cmdExit_Click()
FloatWindow frmMovie.hWnd, SINK
Unload frmMovie
End Sub

Private Sub Form_Load()
FloatWindow frmMovie.hWnd, Float
sMSetting(1) = "216.148.246.34"
sMSetting(2) = "216.148.246.38"
sMSetting(3) = "216.148.246.98"
sMSetting(4) = "216.148.246.40"
sMSetting(5) = "216.148.246.50"
Checked = False
End Sub

Private Sub picCheck_Click(Index As Integer)
'If Checked Then
' lblTemp.Caption = ""
' picCheck(0).Picture = picPlain.Picture
' Checked = False
'Else
' lblTemp.Caption = 0
' picCheck(0).Picture = picChecked.Picture
' Checked = True
'End If
lblTemp.Caption = Index
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
