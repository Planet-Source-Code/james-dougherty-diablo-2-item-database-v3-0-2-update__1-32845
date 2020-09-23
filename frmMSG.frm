VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   Icon            =   "frmMSG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   110
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdOk 
      Height          =   375
      Left            =   2085
      TabIndex        =   0
      Top             =   1080
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
      MICON           =   "frmMSG.frx":000C
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message!"
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
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmMSG.frx":0028
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   1650
      Left            =   0
      Picture         =   "frmMSG.frx":046A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5250
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Sub cmdOk_Click()
On Local Error Resume Next
FloatWindow frmMSG.hWnd, SINK
Unload Me
End Sub

Private Sub Form_Load()
FloatWindow frmMSG.hWnd, Float
Beep 44100, 1
End Sub
