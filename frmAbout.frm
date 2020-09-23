VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdOk 
      Height          =   375
      Left            =   3615
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmAbout.frx":28A2
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beta Tester - Jeff Deaton"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   480
      Width           =   2805
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(C)2002 James E. Dougherty"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   4800
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diablo 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   270
      TabIndex        =   1
      Top             =   240
      Width           =   3750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Database"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   270
      TabIndex        =   0
      Top             =   1200
      Width           =   3780
   End
   Begin VB.Image Image1 
      Height          =   4710
      Left            =   270
      Picture         =   "frmAbout.frx":28BE
      Stretch         =   -1  'True
      Top             =   390
      Width           =   8280
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
FloatWindow frmAbout.hWnd, Float
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmAbout.hWnd, SINK
End Sub
