VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00808080&
   Caption         =   "Diablo 2 Item Database [ ]"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0442
   Begin VB.PictureBox picMenu 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":3111
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
      Begin VB.Shape mnuShape 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   360
         Left            =   1920
         Top             =   15
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Help"
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
         Left            =   2040
         TabIndex        =   26
         Top             =   90
         Width           =   360
      End
      Begin VB.Label lblWindow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Window"
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
         TabIndex        =   25
         Top             =   90
         Width           =   660
      End
      Begin VB.Label lblView 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&View"
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
         Left            =   600
         TabIndex        =   24
         Top             =   90
         Width           =   420
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&File"
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
         Left            =   120
         TabIndex        =   23
         Top             =   90
         Width           =   285
      End
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   1800
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3B13
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4037
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimeTimer 
      Interval        =   1000
      Left            =   5400
      Top             =   3960
   End
   Begin VB.PictureBox picOutputH 
      Align           =   2  'Align Bottom
      Height          =   990
      Left            =   0
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   788
      TabIndex        =   15
      Top             =   7020
      Width           =   11880
      Begin VB.TextBox txtOutput 
         BackColor       =   &H00E0E0E0&
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
         Height          =   810
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Text            =   "frmMain.frx":455B
         Top             =   60
         Width           =   11280
      End
      Begin prjChameleon.chameleonButton cmdClose 
         Height          =   225
         Left            =   90
         TabIndex        =   20
         Top             =   45
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         BTYPE           =   3
         TX              =   "X"
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
         MICON           =   "frmMain.frx":4571
      End
      Begin prjChameleon.chameleonButton cmdClear 
         Height          =   225
         Left            =   90
         TabIndex        =   21
         Top             =   630
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         BTYPE           =   3
         TX              =   "C"
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
         MICON           =   "frmMain.frx":458D
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   5
         X1              =   13
         X2              =   13
         Y1              =   16
         Y2              =   52
      End
      Begin VB.Line Line6 
         X1              =   18
         X2              =   18
         Y1              =   8
         Y2              =   49
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   12
         X2              =   12
         Y1              =   16
         Y2              =   57
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   8
         X2              =   8
         Y1              =   16
         Y2              =   57
      End
      Begin VB.Line Line1 
         X1              =   11
         X2              =   11
         Y1              =   8
         Y2              =   57
      End
      Begin VB.Image InfoWindowBack 
         Height          =   915
         Left            =   0
         Picture         =   "frmMain.frx":45A9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   11925
      End
   End
   Begin VB.PictureBox StatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   14
      Top             =   8010
      Width           =   11880
      Begin VB.Label lblTime 
         BackColor       =   &H00000000&
         Caption         =   " Time Goes Here"
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
         Height          =   180
         Left            =   10485
         TabIndex        =   18
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00000000&
         Caption         =   " Date Goes Here"
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
         Height          =   180
         Left            =   8850
         TabIndex        =   17
         Top             =   60
         Width           =   1500
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00000000&
         Caption         =   " Status Goes Here"
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
         Height          =   180
         Left            =   315
         TabIndex        =   16
         Top             =   60
         Width           =   8220
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   10410
         Picture         =   "frmMain.frx":4BC2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1590
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   8790
         Picture         =   "frmMain.frx":57C5
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1620
      End
      Begin VB.Image Image2 
         Height          =   300
         Left            =   0
         Picture         =   "frmMain.frx":63C8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8790
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7603
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":791F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":83B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":86CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":89EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D07
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9023
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":933F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":965B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCont 
      Align           =   3  'Align Left
      Height          =   6645
      Left            =   0
      ScaleHeight     =   439
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   0
      Top             =   375
      Width           =   1740
      Begin prjChameleon.chameleonButton PicVis 
         Height          =   255
         Left            =   1395
         TabIndex        =   13
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "X"
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
         MICON           =   "frmMain.frx":997F
      End
      Begin VB.TextBox txtAName 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   1590
      End
      Begin VB.PictureBox Window01 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   6
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   111
         TabIndex        =   2
         Top             =   0
         Width           =   1695
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Main Window"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   210
            Left            =   45
            TabIndex        =   3
            Top             =   45
            Width           =   1095
         End
      End
      Begin VB.CheckBox Dummy 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -300
         TabIndex        =   1
         Top             =   1800
         Width           =   255
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add Character"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":999B
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add Jewelry"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":99B7
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add Armor"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":99D3
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   9
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add Weapon"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":99EF
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   495
         Index           =   4
         Left            =   60
         TabIndex        =   10
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "      Add Runes,       Gems, Jewels"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":9A0B
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   11
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Add Charms"
         ENAB            =   0   'False
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
         MICON           =   "frmMain.frx":9A27
      End
      Begin prjChameleon.chameleonButton cmdAddItem 
         Height          =   375
         Index           =   6
         Left            =   60
         TabIndex        =   12
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "New Account"
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
         MICON           =   "frmMain.frx":9A43
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   112
         Y1              =   28
         Y2              =   28
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   112
         Y1              =   29
         Y2              =   29
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   112
         Y1              =   112
         Y2              =   112
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   112
         Y1              =   113
         Y2              =   113
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNT NAME"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   210
         TabIndex        =   4
         Top             =   510
         Width           =   1215
      End
      Begin VB.Image MainWindowBack 
         Height          =   9000
         Left            =   0
         Picture         =   "frmMain.frx":9A5F
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Account"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Account..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close Account"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "C&lose All Accounts"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Account"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Account &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileDummy 
         Caption         =   "       ---Most Recently Used List---"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile4 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewNewFeat 
         Caption         =   "&New Features..."
      End
      Begin VB.Menu mnuViewSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewInfoWin 
         Caption         =   "&Information Window"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewMainWin 
         Caption         =   "&Main Window"
         Checked         =   -1  'True
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNew 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuD2 
      Caption         =   "Diablo II Hacks!!"
      Begin VB.Menu mnuD2Socket 
         Caption         =   "&Socket Items!"
      End
      Begin VB.Menu mnuD2Gem 
         Caption         =   "&Gem Editor!"
      End
      Begin VB.Menu mnuD2Movie 
         Caption         =   "&Movie Hack!"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpBasic 
         Caption         =   "Learn To &Use..."
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddItem_Click(Index As Integer)
On Local Error Resume Next
Dim i As Long

Select Case Index
 Case 0
  frmChar.Show vbModal

  If CharOK = True Then
   ActiveForm.DView.ImageList = imgList
   ActiveForm.DView.Nodes.Add AName, tvwChild, CharName, CharName, 1
   ActiveForm.DView.Nodes.Add CharName, tvwChild, "Jewelry" & CharName, "Jewelry", 2
   ActiveForm.DView.Nodes.Add CharName, tvwChild, "Armor" & CharName, "Armor", 3
   ActiveForm.DView.Nodes.Add CharName, tvwChild, "Weapons" & CharName, "Weapons", 4
   ActiveForm.DView.Nodes.Add CharName, tvwChild, "Jewels/Runes" & CharName, "Jewels/Runes/Gems", 5
   ActiveForm.DView.Nodes.Add CharName, tvwChild, "Charms" & CharName, "Charms", 7
   ActiveForm.DView.Nodes.Item(AName).Expanded = True
   ActiveForm.DView.Nodes.Item(CharName).Expanded = True
   frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & CharName & " Added To Database..."
  
   For i = 1 To 5
    cmdAddItem(i).Enabled = True
   Next
  End If
  
 Case 1
  If ActiveForm.DView.SelectedItem.Image <> 1 Then
   frmMSG.lblMessage.Caption = "Please select a character to add the item to..."
   frmMSG.Show 1
   Exit Sub
  End If
  CharName = ActiveForm.DView.SelectedItem.Text
  
  frmJewelry.Show vbModal
  
  If JewelryOK = True Then
   ActiveForm.DView.ImageList = imgList
   If FromFile = True Then
    If HoldOPValue = 0 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & JewelryName & CharName, JewelryName, 9
    ElseIf HoldOPValue = 1 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & JewelryName & CharName, JewelryName, 10
    ElseIf HoldOPValue = 2 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & JewelryName & CharName, JewelryName, 11
    ElseIf HoldOPValue = 3 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & JewelryName & CharName, JewelryName, 12
    ElseIf HoldOPValue = 4 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & JewelryName & CharName, JewelryName, 13
    Else
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & JewelryName & CharName, JewelryName
    End If
   Else
    If HoldOPValue = 0 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & CharName & JewelryName, JewelryName, 9
    ElseIf HoldOPValue = 1 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & CharName & JewelryName, JewelryName, 10
    ElseIf HoldOPValue = 2 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & CharName & JewelryName, JewelryName, 11
    ElseIf HoldOPValue = 3 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & CharName & JewelryName, JewelryName, 12
    ElseIf HoldOPValue = 4 Then
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & CharName & JewelryName, JewelryName, 13
    Else
     ActiveForm.DView.Nodes.Add "Jewelry" & CharName, tvwChild, "Jewelry1" & CharName & JewelryName, JewelryName
    End If
   End If
   frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & JewelryName & " Added To Database..."
  End If
 
 Case 2
  If ActiveForm.DView.SelectedItem.Image <> 1 Then
   frmMSG.lblMessage.Caption = "Please select a character to add the item to..."
   frmMSG.Show 1
   Exit Sub
  End If
  CharName = ActiveForm.DView.SelectedItem.Text
  
  frmArmor.Show vbModal
  
  If ArmorOK = True Then
   ActiveForm.DView.ImageList = imgList
   If FromFile = True Then
    If HoldOPValue2 = 0 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & ArmorName & CharName, ArmorName, 9
    ElseIf HoldOPValue2 = 1 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & ArmorName & CharName, ArmorName, 10
    ElseIf HoldOPValue2 = 2 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & ArmorName & CharName, ArmorName, 11
    ElseIf HoldOPValue2 = 3 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & ArmorName & CharName, ArmorName, 12
    ElseIf HoldOPValue2 = 4 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & ArmorName & CharName, ArmorName, 13
    Else
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & ArmorName & CharName, ArmorName
    End If
   Else
    If HoldOPValue2 = 0 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & CharName & ArmorName, ArmorName, 9
    ElseIf HoldOPValue2 = 1 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & CharName & ArmorName, ArmorName, 10
    ElseIf HoldOPValue2 = 2 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & CharName & ArmorName, ArmorName, 11
    ElseIf HoldOPValue2 = 3 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & CharName & ArmorName, ArmorName, 12
    ElseIf HoldOPValue2 = 4 Then
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & CharName & ArmorName, ArmorName, 13
    Else
     ActiveForm.DView.Nodes.Add "Armor" & CharName, tvwChild, "Armor1" & CharName & ArmorName, ArmorName
    End If
   End If
   frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & ArmorName & " Added To Database..."
  End If
  Exit Sub
  
 Case 3
  If ActiveForm.DView.SelectedItem.Image <> 1 Then
   frmMSG.lblMessage.Caption = "Please select a character to add the item to..."
   frmMSG.Show 1
   Exit Sub
  End If
  CharName = ActiveForm.DView.SelectedItem.Text
  
  frmWeapon.Show vbModal
  
  If WeaponOK = True Then
   ActiveForm.DView.ImageList = imgList
   If FromFile = True Then
    If HoldOPValue3 = 0 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & WeaponName & CharName, WeaponName, 9
    ElseIf HoldOPValue3 = 1 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & WeaponName & CharName, WeaponName, 10
    ElseIf HoldOPValue3 = 2 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & WeaponName & CharName, WeaponName, 11
    ElseIf HoldOPValue3 = 3 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & WeaponName & CharName, WeaponName, 12
    ElseIf HoldOPValue3 = 4 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & WeaponName & CharName, WeaponName, 13
    Else
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & WeaponName & CharName, WeaponName
    End If
   Else
    If HoldOPValue3 = 0 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & CharName & WeaponName, WeaponName, 9
    ElseIf HoldOPValue3 = 1 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & CharName & WeaponName, WeaponName, 10
    ElseIf HoldOPValue3 = 2 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & CharName & WeaponName, WeaponName, 11
    ElseIf HoldOPValue3 = 3 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & CharName & WeaponName, WeaponName, 12
    ElseIf HoldOPValue3 = 4 Then
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & CharName & WeaponName, WeaponName, 13
    Else
     ActiveForm.DView.Nodes.Add "Weapons" & CharName, tvwChild, "Weapon1" & CharName & WeaponName, WeaponName
    End If
   End If
   frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & WeaponName & " Added To Database..."
  End If
  Exit Sub
  
 Case 4
  If ActiveForm.DView.SelectedItem.Image <> 1 Then
   frmMSG.lblMessage.Caption = "Please select a character to add the item to..."
   frmMSG.Show 1
   Exit Sub
  End If
  CharName = ActiveForm.DView.SelectedItem.Text
  
  frmJewels.Show vbModal
  
  If JewelOK = True Then
   ActiveForm.DView.ImageList = imgList
   If FromFile = True Then
    ActiveForm.DView.Nodes.Add "Jewels/Runes" & CharName, tvwChild, "Jewel1" & CharName & JewelName & AName, JewelName, 5
   Else
    ActiveForm.DView.Nodes.Add "Jewels/Runes" & CharName, tvwChild, "Jewel1" & CharName & JewelName & AName, JewelName, 5
   End If
   frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & JewelName & " Added To Database..."
  End If
  Exit Sub
  
 Case 5
  If ActiveForm.DView.SelectedItem.Image <> 1 Then
   frmMSG.lblMessage.Caption = "Please select a character to add the item to..."
   frmMSG.Show 1
   Exit Sub
  End If
  CharName = ActiveForm.DView.SelectedItem.Text
  
  frmCharm.Show vbModal
  
  If CharmOK = True Then
   ActiveForm.DView.ImageList = imgList
   If FromFile = True Then
    ActiveForm.DView.Nodes.Add "Charms" & CharName, tvwChild, "Charm1" & CharName & CharmName & AName, CharmName, 7
   Else
    ActiveForm.DView.Nodes.Add "Charms" & CharName, tvwChild, "Charm1" & CharName & CharmName & AName, CharmName, 7
   End If
   frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & CharmName & " Added To Database..."
  End If
  Exit Sub
  
 Case 6
   Dim NewFrm As frmDocument
   frmSetup.Show vbModal

   If OK = True Then
    frmMain.txtAName.Text = AName
    frmMain.picCont.Enabled = True
 
    Set NewFrm = New frmDocument
    NewFrm.Caption = AName
    NewFrm.Show
    NewFrm.DView.ImageList = imgList
    NewFrm.DView.Nodes.Add , , AName, AName, 8
    frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & AName & " Loaded..."
    cmdAddItem(0).Enabled = True
    FromFile = False
   End If
End Select
End Sub

Private Sub cmdAddItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
 Case 0
  lblStatus.Caption = " Add a new character to an account"
 Case 1
  lblStatus.Caption = " Add a jewelry to a selected character"
 Case 2
  lblStatus.Caption = " Add an armor to a selected character"
 Case 3
  lblStatus.Caption = " Add a weapon to a selected character"
 Case 4
  lblStatus.Caption = " Add a jewel, rune, or gem to a selected character"
 Case 5
  lblStatus.Caption = " Add a charm to a selected character"
 Case 6
  lblStatus.Caption = " Create a new account"
End Select
End Sub

Private Sub cmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus.Caption = " Clear information list"
End Sub

Private Sub cmdClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus.Caption = " Close's the information window. Click [View->Information window] to bring back if wanted"
End Sub

Private Sub InfoWindowBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus.Caption = " Displays information of what actions are or were being processed"
End Sub

Private Sub lblFile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuShape.Move 1, 1, 32
mnuShape.Visible = True
lblStatus.Caption = " File->"
End Sub

Private Sub lblHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuShape.Move 128, 1, 40
mnuShape.Visible = True
lblStatus.Caption = " Help->"
End Sub

Private Sub lblView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuShape.Move 34, 1, 40
mnuShape.Visible = True
lblStatus.Caption = " View->"
End Sub

Private Sub lblWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuShape.Move 73, 1, 56
mnuShape.Visible = True
lblStatus.Caption = " Window->"
End Sub

Private Sub MainWindowBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus.Caption = " "
End Sub

Private Sub MDIForm_Load()
On Local Error Resume Next
Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 30)
Me.Top = GetSetting(App.Title, "Settings", "MainTop", -30)
Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 11940)
Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 9045)
txtOutput.Text = "Information Window:"
SetMenuBitmaps
If Command$ <> "" Then CommLineFile
LoadRescentFiles
mnuViewInfoWin.Checked = GetSetting(App.Title, "Settings", "View1", True)
mnuViewMainWin.Checked = GetSetting(App.Title, "Settings", "View2", True)
picOutputH.Visible = mnuViewInfoWin.Checked
picCont.Visible = mnuViewMainWin.Checked
Draw_Gradient_Title_Bar_Horizontal Window01, &HC00000, &HFFFF00
lblDate.Caption = "   Date - " & Format$(Now, "mm-dd-yy")
lblTime.Caption = "  Time - " & Time$
FloatWindow Me.hWnd, Float
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
lblStatus.Caption = " "
mnuShape.Visible = False
If Forms.Count > 1 Then
 frmMain.txtAName.Text = ActiveForm.Caption
Else
 Exit Sub
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Local Error Resume Next
If Me.WindowState <> vbMinimized Then
 SaveSetting App.Title, "Settings", "MainLeft", Me.Left
 SaveSetting App.Title, "Settings", "MainTop", Me.Top
 SaveSetting App.Title, "Settings", "MainWidth", Me.Width
 SaveSetting App.Title, "Settings", "MainHeight", Me.Height
End If
FloatWindow Me.hWnd, Sink
End
End Sub

Private Sub mnuD2Gem_Click()
frmGems.Show
End Sub

Private Sub mnuD2Movie_Click()
frmMovie.Show
End Sub

Private Sub mnuD2Socket_Click()
frmSocket.Show
FloatWindow frmSocket.hWnd, Float
End Sub

Private Sub mnuFile1_Click()
LoadMRUFile mnuFile1.Tag
End Sub

Private Sub mnuFile2_Click()
LoadMRUFile mnuFile2.Tag
End Sub

Private Sub mnuFile3_Click()
LoadMRUFile mnuFile3.Tag
End Sub

Private Sub mnuFile4_Click()
LoadMRUFile mnuFile4.Tag
End Sub

Private Sub mnuFileClose_Click()
On Local Error Resume Next
Dim i As Long

If ActiveForm Is Nothing Then
 frmMSG.lblMessage.Caption = "No open accounts to close..."
 frmMSG.Show 1
 Exit Sub
End If

txtOutput.Text = txtOutput.Text & vbCrLf & ActiveForm.Caption & " Unloaded"
Unload ActiveForm

If Forms.Count > 1 Then
 txtAName.Text = ActiveForm.Caption
 Exit Sub
Else
 For i = 0 To 5
  cmdAddItem(i).Enabled = False
 Next
 txtAName.Text = ""
End If
End Sub

Private Sub mnuFileCloseAll_Click()
On Local Error Resume Next
Dim i As Long

If ActiveForm Is Nothing Then
 frmMSG.lblMessage.Caption = "No open accounts to close..."
 frmMSG.Show 1
 Exit Sub
End If

While Forms.Count > 1
 Unload ActiveForm
Wend

For i = 0 To 5
 cmdAddItem(i).Enabled = False
Next

txtOutput.Text = txtOutput.Text & vbCrLf & "All Accounts Unloaded"
txtAName.Text = ""
End Sub

Private Sub mnuFileNew_Click()
On Local Error Resume Next
Dim NewFrm As frmDocument
frmSetup.Show vbModal

If OK = True Then
 frmMain.txtAName.Text = AName
 frmMain.picCont.Enabled = True
 
 Set NewFrm = New frmDocument
 NewFrm.Caption = AName
 NewFrm.Show
 NewFrm.DView.ImageList = imgList
 NewFrm.DView.Nodes.Add , , AName, AName, 8
 frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & AName & " Loaded..."
 cmdAddItem(0).Enabled = True
 FromFile = False
End If

End Sub

Private Sub mnuFileOpen_Click()
LoadRegularFile
End Sub

Private Sub mnuFileSave_Click()
On Local Error Resume Next
Dim Nd As Node, CharNd As Node, JewelryNd As Node, ArmorNd As Node
Dim Char As Integer
Dim Jewelry As Integer
Dim Jewelrys As Integer
Dim Armor As Integer
Dim Armors As Integer
Dim NdValues As Integer
Dim Path As String
Dim tmpStr As String

If HoldFileTitle = "" Then
 Call mnuFileSaveAs_Click
 Exit Sub
End If

If AName = "" Then
 frmMSG.lblMessage.Caption = "Nothing to save at this time..."
 frmMSG.Show 1
 Exit Sub
End If

frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & tmpFileName & " Saved..."
Path = tmpFileName & HoldFileTitle
Set OutStream = FSys.CreateTextFile(Path, True, False)

Set Nd = ActiveForm.DView.Nodes.Item(AName)
NdValues = Nd.Children
OutStream.WriteLine Nd.Text
Set CharNd = Nd.Child

For Char = 1 To NdValues
 OutStream.WriteLine " " & CharNd.Text
 Jewelrys = CharNd.Children
 Set JewelryNd = CharNd.Child
  For Jewelry = 1 To Jewelrys
   OutStream.WriteLine "  " & JewelryNd.Text
   Armors = JewelryNd.Children
   Set ArmorNd = JewelryNd.Child
    For Armor = 1 To Armors
     OutStream.WriteLine "   " & ArmorNd.Image & ArmorNd.Text
     Set ArmorNd = ArmorNd.Next
    Next
   Set JewelryNd = JewelryNd.Next
  Next
 Set CharNd = CharNd.Next
Next

Set OutStream = Nothing
End Sub

Private Sub mnuFileSaveAs_Click()
On Local Error Resume Next
Dim Nd As Node, CharNd As Node, JewelryNd As Node, ArmorNd As Node
Dim Char As Integer
Dim Jewelry As Integer
Dim Jewelrys As Integer
Dim Armor As Integer
Dim Armors As Integer
Dim NdValues As Integer
Dim Path As String
Dim tmpStr As String

With CD
 .Filter = "D2ID (*.did)|*.did"
 .CancelError = False
 .InitDir = App.Path
 .ShowSave
 HoldFileTitle = .FileTitle
 tmpFileName = Left$(.FileName, Len(.FileName) - Len(.FileTitle))
 AddToList .FileName
 Me.Caption = "Diablo 2 Item Database [ " & .FileTitle & " ]"
End With

frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & tmpFileName & " Saved..."
Path = tmpFileName & HoldFileTitle
Set OutStream = FSys.CreateTextFile(Path, True, False)

Set Nd = ActiveForm.DView.Nodes.Item(AName)
NdValues = Nd.Children
OutStream.WriteLine Nd.Text
Set CharNd = Nd.Child

For Char = 1 To NdValues
 OutStream.WriteLine " " & CharNd.Text
 Jewelrys = CharNd.Children
 Set JewelryNd = CharNd.Child
  For Jewelry = 1 To Jewelrys
   OutStream.WriteLine "  " & JewelryNd.Text
   Armors = JewelryNd.Children
   Set ArmorNd = JewelryNd.Child
    For Armor = 1 To Armors
     OutStream.WriteLine "   " & ArmorNd.Image & ArmorNd.Text
     Set ArmorNd = ArmorNd.Next
    Next
   Set JewelryNd = JewelryNd.Next
  Next
 Set CharNd = CharNd.Next
Next

Set OutStream = Nothing
End Sub

Private Sub mnuHelpAbout_Click()
On Local Error Resume Next
frmAbout.Show vbModal
End Sub

Private Sub mnuHelpBasic_Click()
On Local Error Resume Next
frmUse.Show vbModal
End Sub

Private Sub mnuViewInfoWin_Click()
On Local Error Resume Next
mnuViewInfoWin.Checked = Not mnuViewInfoWin.Checked
picOutputH.Visible = mnuViewInfoWin.Checked
If mnuViewMainWin.Checked Then
 txtOutput.Text = txtOutput.Text & vbCrLf & "Information Window Opened"
End If
SaveSetting App.Title, "Settings", "View1", mnuViewInfoWin.Checked
End Sub

Private Sub mnuViewMainWin_Click()
On Local Error Resume Next
mnuViewMainWin.Checked = Not mnuViewMainWin.Checked
picCont.Visible = mnuViewMainWin.Checked
If mnuViewMainWin.Checked Then
 txtOutput.Text = txtOutput.Text & vbCrLf & "Main Window Opened"
End If
SaveSetting App.Title, "Settings", "View2", mnuViewMainWin.Checked
End Sub

Private Sub mnuViewNewFeat_Click()
frmNewFeatures.Show
End Sub

Private Sub mnuWindowArrangeIcons_Click()
On Local Error Resume Next
Me.Arrange vbArrangeIcons
txtOutput.Text = txtOutput.Text & vbCrLf & "Window's Arranged"
End Sub

Private Sub mnuWindowNew_Click()
On Local Error Resume Next
Dim NewFrm As frmDocument
frmSetup.Show vbModal

If OK = True Then
 frmMain.txtAName.Text = AName
 frmMain.picCont.Enabled = True
 
 Set NewFrm = New frmDocument
 NewFrm.Caption = AName
 NewFrm.Show
 NewFrm.DView.ImageList = imgList
 NewFrm.DView.Nodes.Add , , AName, AName, 8
 frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & AName & " Loaded..."
 cmdAddItem(0).Enabled = True
 FromFile = False
End If
End Sub

Private Sub mnuWindowTileVertical_Click()
On Local Error Resume Next
Me.Arrange vbTileVertical
txtOutput.Text = txtOutput.Text & vbCrLf & "Window's Tiled Vertically"
End Sub

Private Sub mnuWindowTileHorizontal_Click()
On Local Error Resume Next
Me.Arrange vbTileHorizontal
txtOutput.Text = txtOutput.Text & vbCrLf & "Window's Tiled Horizontally"
End Sub

Private Sub mnuWindowCascade_Click()
On Local Error Resume Next
Me.Arrange vbCascade
txtOutput.Text = txtOutput.Text & vbCrLf & "Window's Cascade"
End Sub

Private Sub mnuFileExit_Click()
On Local Error Resume Next
txtOutput.Text = "Goodbye"
Unload Me
End Sub

Private Sub picCont_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Forms.Count > 1 Then
 frmMain.txtAName.Text = ActiveForm.Caption
Else
 Exit Sub
End If
End Sub

Private Sub picCont_Resize()
picCont.Refresh
End Sub

Private Sub picMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuShape.Visible = False
lblStatus.Caption = " "
End Sub

Private Sub picOutputH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
If Forms.Count > 1 Then
 frmMain.txtAName.Text = ActiveForm.Caption
Else
 Exit Sub
End If
End Sub

Private Sub PicVis_Click()
On Local Error Resume Next
picCont.Visible = False
mnuViewMainWin.Checked = False
txtOutput.Text = txtOutput.Text & vbCrLf & "Main Window Close Complete"
End Sub

Private Sub cmdClear_Click()
On Local Error Resume Next
txtOutput.Text = "Information Window:"
End Sub

Private Sub cmdClose_Click()
On Local Error Resume Next
picOutputH.Visible = False
mnuViewInfoWin.Checked = False
txtOutput.Text = txtOutput.Text & vbCrLf & "Information Window Close Complete"
End Sub

Private Function CommLineFile()
On Local Error Resume Next
Dim NewFrm As frmDocument
Dim tmpInStr As String
Dim Key As String
Dim ThisFile As String
Dim tmpTrimed As String
Dim DoOneTime As Boolean
Dim i As Long
Dim sFile As String
Dim fType As String
    
If Command$ <> "" Then
 sFile = Command$
End If
  
On Error GoTo CommResume
Set InStream = FSys.OpenTextFile(sFile)

Set NewFrm = New frmDocument
NewFrm.Show
NewFrm.DView.ImageList = imgList
   
DoOneTime = False

While InStream.AtEndOfStream = False
 tmpInStr = InStream.ReadLine
 
 If DoOneTime = False Then
  AName = tmpInStr
  NewFrm.Caption = AName
  txtAName.Text = AName
  NewFrm.DView.Nodes.Add , , AName, AName, 8
  DoOneTime = True
 End If
 
 If Left(tmpInStr, 3) = "   " Then
  If Left(tmpInStr, 4) = "   9" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 9
  ElseIf Left(tmpInStr, 5) = "   10" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 10
  ElseIf Left(tmpInStr, 5) = "   11" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 11
  ElseIf Left(tmpInStr, 5) = "   12" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 12
  ElseIf Left(tmpInStr, 5) = "   13" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 13
  ElseIf Left(tmpInStr, 4) = "   7" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 7
  ElseIf Left(tmpInStr, 4) = "   5" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 5
  End If
 ElseIf Left(tmpInStr, 2) = "  " Then
  If Right$(tmpInStr, 7) = "Jewelry" Then
   Key = "Jewelry" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Jewelry" & CharName, Trim(tmpInStr), 2
  ElseIf Right$(tmpInStr, 5) = "Armor" Then
   Key = "Armor" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Armor" & CharName, Trim(tmpInStr), 3
  ElseIf Right$(tmpInStr, 7) = "Weapons" Then
   Key = "Weapons" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Weapons" & CharName, Trim(tmpInStr), 4
  ElseIf Right$(tmpInStr, 17) = "Jewels/Runes/Gems" Then
   Key = "Jewels/Runes" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Jewels/Runes" & CharName, Trim(tmpInStr), 5
  ElseIf Right$(tmpInStr, 6) = "Charms" Then
   Key = "Charms" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Charms" & CharName, Trim(tmpInStr), 7
  End If
 ElseIf Left(tmpInStr, 1) = " " Then
  tmpTrimed = Left(tmpInStr, 1)
  CharName = Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed)))
  NewFrm.DView.Nodes.Add AName, tvwChild, CharName, Trim(tmpInStr), 1
 End If
Wend
   
Set InStream = Nothing

For i = 0 To 5
 cmdAddItem(i).Enabled = True
Next
ActiveForm.DView.Nodes.Item(AName).Expanded = True
ActiveForm.DView.Nodes.Item(CharName).Expanded = True
FromFile = True
Exit Function
     
CommResume:
End Function

Private Function FileExists(sFileName As String) As Boolean
On Error GoTo FExistsError
Dim f As String
f = FreeFile
Open sFileName For Input As #f
Close #f
    
FExistsError:
If Err.Number = 53 Then
 FileExists = False
ElseIf Err.Number = 0 Then
 FileExists = True
End If
End Function

Private Sub LoadMRUFile(FileName As String)
On Local Error Resume Next
Dim NewFrm As frmDocument
Dim tmpInStr As String
Dim Key As String
Dim ThisFile As String
Dim tmpTrimed As String
Dim tmpTrimed2 As String
Dim DoOneTime As Boolean
Dim sFile As String
Dim i As Long
Dim j As Long

If FileExists(FileName) = True Then
 Set InStream = FSys.OpenTextFile(FileName)

Set NewFrm = New frmDocument
NewFrm.Show
NewFrm.DView.ImageList = imgList
   
DoOneTime = False

While InStream.AtEndOfStream = False
 tmpInStr = InStream.ReadLine
 
 If DoOneTime = False Then
  AName = tmpInStr
  NewFrm.Caption = AName
  txtAName.Text = AName
  NewFrm.DView.Nodes.Add , , AName, AName, 8
  DoOneTime = True
 End If
 
 If Left(tmpInStr, 3) = "   " Then
  If Left(tmpInStr, 4) = "   9" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 9
  ElseIf Left(tmpInStr, 5) = "   10" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 10
  ElseIf Left(tmpInStr, 5) = "   11" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 11
  ElseIf Left(tmpInStr, 5) = "   12" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 12
  ElseIf Left(tmpInStr, 5) = "   13" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 13
  ElseIf Left(tmpInStr, 4) = "   7" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 7
  ElseIf Left(tmpInStr, 4) = "   5" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 5
  End If
 ElseIf Left(tmpInStr, 2) = "  " Then
  If Right$(tmpInStr, 7) = "Jewelry" Then
   Key = "Jewelry" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Jewelry" & CharName, Trim(tmpInStr), 2
  ElseIf Right$(tmpInStr, 5) = "Armor" Then
   Key = "Armor" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Armor" & CharName, Trim(tmpInStr), 3
  ElseIf Right$(tmpInStr, 7) = "Weapons" Then
   Key = "Weapons" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Weapons" & CharName, Trim(tmpInStr), 4
  ElseIf Right$(tmpInStr, 17) = "Jewels/Runes/Gems" Then
   Key = "Jewels/Runes" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Jewels/Runes" & CharName, Trim(tmpInStr), 5
  ElseIf Right$(tmpInStr, 6) = "Charms" Then
   Key = "Charms" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Charms" & CharName, Trim(tmpInStr), 7
  End If
 ElseIf Left(tmpInStr, 1) = " " Then
  tmpTrimed = Left(tmpInStr, 1)
  CharName = Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed)))
  NewFrm.DView.Nodes.Add AName, tvwChild, CharName, Trim(tmpInStr), 1
 End If
Wend
   
Set InStream = Nothing

For i = 0 To 5
 cmdAddItem(i).Enabled = True
Next

ActiveForm.DView.Nodes.Item(AName).Expanded = True
ActiveForm.DView.Nodes.Item(CharName).Expanded = True
txtOutput.Text = txtOutput.Text & vbCrLf & sFile & " loaded..."
FromFile = True
Else
 frmMSG.lblMessage.Caption = "File does not exist." & vbCrLf & " Make sure it wasn't moved or deleted."
 frmMSG.Show 1
 txtOutput.Text = txtOutput.Text & vbCrLf & "File Load Failed..."
 FromFile = False
End If
End Sub

Private Sub LoadRegularFile()
Dim NewFrm As frmDocument
Dim tmpInStr As String
Dim Key As String
Dim ThisFile As String
Dim tmpTrimed As String
Dim DoOneTime As Boolean
Dim i As Long
Dim Jewe As String

On Local Error GoTo ErrOut

With CD
 .Filter = "D2ID (*.did)|*.did"
 .CancelError = False
 .InitDir = App.Path
 .ShowOpen
 ThisFile = .FileName
End With

If ThisFile = "" Then Exit Sub
AddToList ThisFile
Me.Caption = "Diablo 2 Item Database [ " & ThisFile & " ]"
Set InStream = FSys.OpenTextFile(ThisFile)

Set NewFrm = New frmDocument
NewFrm.Show
NewFrm.DView.ImageList = imgList
   
DoOneTime = False

While InStream.AtEndOfStream = False
 tmpInStr = InStream.ReadLine
 
 If DoOneTime = False Then
  If Left(tmpInStr, 1) <> " " Then
   AName = tmpInStr
   NewFrm.Caption = AName
   txtAName.Text = AName
   NewFrm.DView.Nodes.Add , , AName, AName, 8
   DoOneTime = True
  Else
   frmMSG.lblMessage.Caption = "File seems to be corrupt." & vbCrLf & " Make sure the file hasn't been tampered with."
   frmMSG.Show 1
   txtOutput.Text = txtOutput.Text & vbCrLf & "File Load Failed..."
   FromFile = False
   Exit Sub
  End If
 End If
 
 If Left(tmpInStr, 3) = "   " Then
  If Left(tmpInStr, 4) = "   9" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 9
  ElseIf Left(tmpInStr, 5) = "   10" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 10
  ElseIf Left(tmpInStr, 5) = "   11" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 11
  ElseIf Left(tmpInStr, 5) = "   12" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 12
  ElseIf Left(tmpInStr, 5) = "   13" Then
   tmpTrimed = Left(tmpInStr, 5)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 13
  ElseIf Left(tmpInStr, 4) = "   7" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 7
  ElseIf Left(tmpInStr, 4) = "   5" Then
   tmpTrimed = Left(tmpInStr, 4)
   NewFrm.DView.Nodes.Add Key, tvwChild, Key & tmpInStr, Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed))), 5
  End If
 ElseIf Left(tmpInStr, 2) = "  " Then
  If Right$(tmpInStr, 7) = "Jewelry" Then
   Key = "Jewelry" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Jewelry" & CharName, Trim(tmpInStr), 2
  ElseIf Right$(tmpInStr, 5) = "Armor" Then
   Key = "Armor" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Armor" & CharName, Trim(tmpInStr), 3
  ElseIf Right$(tmpInStr, 7) = "Weapons" Then
   Key = "Weapons" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Weapons" & CharName, Trim(tmpInStr), 4
  ElseIf Right$(tmpInStr, 17) = "Jewels/Runes/Gems" Then
   Key = "Jewels/Runes" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Jewels/Runes" & CharName, Trim(tmpInStr), 5
  ElseIf Right$(tmpInStr, 6) = "Charms" Then
   Key = "Charms" & CharName
   NewFrm.DView.Nodes.Add CharName, tvwChild, "Charms" & CharName, Trim(tmpInStr), 7
  End If
 ElseIf Left(tmpInStr, 1) = " " Then
  tmpTrimed = Left(tmpInStr, 1)
  CharName = Trim(Right(tmpInStr, Len(tmpInStr) - Len(tmpTrimed)))
  NewFrm.DView.Nodes.Add AName, tvwChild, CharName, Trim(tmpInStr), 1
 End If
Wend
   
Set InStream = Nothing

For i = 0 To 5
 cmdAddItem(i).Enabled = True
Next
ActiveForm.DView.Nodes.Item(AName).Expanded = True
ActiveForm.DView.Nodes.Item(CharName).Expanded = True
frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & ThisFile & " Loaded..."
FromFile = True
Exit Sub

ErrOut:
 frmMSG.lblMessage.Caption = "File seems to be corrupt." & vbCrLf & " Make sure the file hasn't been tampered with."
 frmMSG.Show 1
 txtOutput.Text = txtOutput.Text & vbCrLf & "File Load Failed..."
 FromFile = False
 Exit Sub
End Sub

Private Sub PicVis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus.Caption = " Close's the main window. Click [View->Main window] to bring back if wanted"
End Sub

Private Sub TimeTimer_Timer()
lblTime.Caption = "  Time - " & Time$
End Sub

Private Sub txtAName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtAName = "" Then
 lblStatus.Caption = " Shows your current accounts name"
Else
 lblStatus.Caption = " Shows your current accounts name. Currently your account name is " & txtAName.Text
End If
End Sub

Private Sub txtOutput_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblStatus.Caption = " Displays information of what actions are or were being processed"
End Sub

Private Sub SetMenuBitmaps()
Dim lngMenu, lngSubMenu As Long

lngMenu = GetMenu(Me.hWnd)
lngSubMenu = GetSubMenu(lngMenu, 0)

Call SetMenuItemBitmaps(lngSubMenu, 0, MF_BYPOSITION, IL.ListImages(1).Picture, IL.ListImages(1).Picture)
Call SetMenuItemBitmaps(lngSubMenu, 1, MF_BYPOSITION, IL.ListImages(2).Picture, IL.ListImages(2).Picture)
Call SetMenuItemBitmaps(lngSubMenu, 6, MF_BYPOSITION, IL.ListImages(3).Picture, IL.ListImages(3).Picture)
End Sub

Private Sub Window01_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mnuShape.Visible = False
End Sub
