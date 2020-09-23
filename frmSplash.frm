VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblLoad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   960
   End
   Begin VB.Label Label3 
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
      TabIndex        =   1
      Top             =   1200
      Width           =   3780
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
      TabIndex        =   0
      Top             =   240
      Width           =   3750
   End
   Begin VB.Image Image1 
      Height          =   4710
      Left            =   270
      Picture         =   "frmSplash.frx":28A2
      Stretch         =   -1  'True
      Top             =   390
      Width           =   8280
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Single
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long

Private Sub Form_Load()
On Local Error Resume Next
Dim ValueOnce As Integer
Dim fLogin As New frmLogin

frmSplash.Show
Do Until i >= 2000
 DoEvents
 i = i + 0.01
Loop
lblLoad.Caption = "Loading Complete"
fLogin.Show 1
If Not fLogin.OK Then End
Unload fLogin
frmSplash.Refresh
frmMain.Show

ValueOnce = GetSetting(App.Title, "Settings", "Once", -1)
If ValueOnce = -1 Then
 DoOnce
 ValueOnce = 0
 SaveSetting App.Title, "Settings", "Once", ValueOnce
End If
Unload frmSplash
End Sub

Private Sub DoOnce()
Dim Program As String
Dim Icon As String
Program = App.Path & "\Diablo2 Item Database.exe"
Icon = App.Path & "\didDoc2.ico"
Associate Program, "did", "Diablo2 Item Database File", Icon
End Sub

