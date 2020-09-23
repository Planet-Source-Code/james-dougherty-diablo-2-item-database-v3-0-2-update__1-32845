VERSION 5.00
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmUse 
   BorderStyle     =   0  'None
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   Icon            =   "frmUse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjChameleon.chameleonButton cmdClose 
      Height          =   375
      Left            =   300
      TabIndex        =   1
      ToolTipText     =   "Close Help Window"
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
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
      MICON           =   "frmUse.frx":000C
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   0
      Picture         =   "frmUse.frx":0028
      ScaleHeight     =   363
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   586
      TabIndex        =   2
      Top             =   0
      Width           =   8790
      Begin VB.PictureBox Step 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4275
         Index           =   1
         Left            =   240
         Picture         =   "frmUse.frx":28BE
         ScaleHeight     =   285
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   8265
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1500
            Left            =   3720
            Picture         =   "frmUse.frx":4371
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   250
            TabIndex        =   12
            Top             =   1440
            Width           =   3750
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2010
            Left            =   1080
            Picture         =   "frmUse.frx":5C83
            ScaleHeight     =   134
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   114
            TabIndex        =   13
            Top             =   1200
            Width           =   1710
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Click ""Add Character"" to begin"
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
            Left            =   840
            TabIndex        =   15
            Top             =   3240
            Width           =   2505
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step 2 - Creating Your Character"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            TabIndex        =   16
            Top             =   360
            Width           =   4500
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insert the name of your character and press ""Ok"""
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
            Left            =   3600
            TabIndex        =   14
            Top             =   3000
            Width           =   4080
         End
      End
      Begin VB.PictureBox Step 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4275
         Index           =   0
         Left            =   240
         Picture         =   "frmUse.frx":713B
         ScaleHeight     =   285
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   5
         Top             =   360
         Width           =   8265
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1500
            Left            =   3720
            Picture         =   "frmUse.frx":8BEE
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   250
            TabIndex        =   9
            Top             =   1440
            Width           =   3750
         End
         Begin VB.PictureBox picStep1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1830
            Left            =   1080
            Picture         =   "frmUse.frx":A5EF
            ScaleHeight     =   122
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   115
            TabIndex        =   6
            Top             =   1200
            Width           =   1725
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insert the name of your account and press ""Ok"""
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
            Left            =   3600
            TabIndex        =   10
            Top             =   3000
            Width           =   3945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Click ""New Account"" to begin"
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
            Left            =   840
            TabIndex        =   8
            Top             =   3000
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step 1 - Creating An Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   7
            Top             =   360
            Width           =   3975
         End
      End
      Begin VB.PictureBox Step 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4275
         Index           =   3
         Left            =   240
         Picture         =   "frmUse.frx":BC64
         ScaleHeight     =   285
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   8265
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2430
            Left            =   1920
            Picture         =   "frmUse.frx":D717
            ScaleHeight     =   162
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   304
            TabIndex        =   28
            Top             =   960
            Width           =   4560
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step 4 - Visual After Item Is Added"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1800
            TabIndex        =   31
            Top             =   360
            Width           =   4650
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "After your item is added it will show the item's"
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
            Left            =   2280
            TabIndex        =   30
            Top             =   3480
            Width           =   3885
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "type (if any) and the items name"
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
            Left            =   2880
            TabIndex        =   29
            Top             =   3720
            Width           =   2655
         End
      End
      Begin VB.PictureBox Step 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4275
         Index           =   4
         Left            =   240
         Picture         =   "frmUse.frx":10585
         ScaleHeight     =   285
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   32
         Top             =   600
         Visible         =   0   'False
         Width           =   8265
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2595
            Left            =   6480
            Picture         =   "frmUse.frx":12038
            ScaleHeight     =   173
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   48
            Top             =   840
            Width           =   1005
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2490
            Left            =   3960
            Picture         =   "frmUse.frx":13395
            ScaleHeight     =   166
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   41
            Top             =   840
            Width           =   1005
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2490
            Left            =   1080
            Picture         =   "frmUse.frx":146F9
            ScaleHeight     =   166
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   67
            TabIndex        =   33
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C2 -"
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
            Left            =   5760
            TabIndex        =   54
            Top             =   3720
            Width           =   315
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deletes a selected node"
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
            Left            =   6120
            TabIndex        =   53
            Top             =   3720
            Width           =   1995
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C1 -"
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
            Left            =   5760
            TabIndex        =   52
            Top             =   3480
            Width           =   315
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C2"
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
            Left            =   7560
            TabIndex        =   51
            Top             =   2160
            Width           =   210
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C1"
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
            Left            =   7560
            TabIndex        =   50
            Top             =   960
            Width           =   210
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Refreshes the database"
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
            Left            =   6120
            TabIndex        =   49
            Top             =   3480
            Width           =   1980
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B2 -"
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
            Left            =   3120
            TabIndex        =   47
            Top             =   3720
            Width           =   300
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Collaps's a selected node"
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
            Left            =   3480
            TabIndex        =   46
            Top             =   3720
            Width           =   2130
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B1 -"
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
            Left            =   3120
            TabIndex        =   45
            Top             =   3480
            Width           =   300
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B2"
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
            Left            =   5040
            TabIndex        =   44
            Top             =   2160
            Width           =   195
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B1"
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
            Left            =   5040
            TabIndex        =   43
            Top             =   960
            Width           =   195
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expands a selected node"
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
            Left            =   3480
            TabIndex        =   42
            Top             =   3480
            Width           =   2055
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A2 -"
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
            TabIndex        =   40
            Top             =   3720
            Width           =   315
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Collaps's all nodes in the tree"
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
            Left            =   480
            TabIndex        =   39
            Top             =   3720
            Width           =   2460
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A1 -"
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
            TabIndex        =   38
            Top             =   3480
            Width           =   315
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A2"
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
            Left            =   2160
            TabIndex        =   37
            Top             =   2160
            Width           =   210
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A1"
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
            Left            =   2160
            TabIndex        =   36
            Top             =   960
            Width           =   210
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step 5 - Overview Of The Database Buttons"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   35
            Top             =   360
            Width           =   5970
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expands all nodes in the tree"
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
            Left            =   480
            TabIndex        =   34
            Top             =   3480
            Width           =   2385
         End
      End
      Begin VB.PictureBox Step 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4275
         Index           =   2
         Left            =   -5520
         Picture         =   "frmUse.frx":1577D
         ScaleHeight     =   285
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   551
         TabIndex        =   19
         Top             =   5280
         Visible         =   0   'False
         Width           =   8265
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2550
            Index           =   0
            Left            =   1080
            Picture         =   "frmUse.frx":17230
            ScaleHeight     =   170
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   110
            TabIndex        =   21
            Top             =   960
            Width           =   1650
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2250
            Left            =   3720
            Picture         =   "frmUse.frx":194C8
            ScaleHeight     =   150
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   250
            TabIndex        =   20
            Top             =   1200
            Width           =   3750
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note - Some items require you select the item type"
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
            Left            =   3525
            TabIndex        =   26
            Top             =   3720
            Width           =   4230
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "your items to the database"
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
            Left            =   840
            TabIndex        =   25
            Top             =   3840
            Width           =   2220
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Insert the name of your item and press ""ok"""
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
            Left            =   3780
            TabIndex        =   24
            Top             =   3480
            Width           =   3660
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Step 3 - Adding Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2520
            TabIndex        =   23
            Top             =   360
            Width           =   2925
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Click these buttons to add "
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
            Left            =   840
            TabIndex        =   22
            Top             =   3600
            Width           =   2220
         End
      End
      Begin prjChameleon.chameleonButton StepFor 
         Height          =   375
         Left            =   8025
         TabIndex        =   3
         ToolTipText     =   "Step Forward"
         Top             =   4680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ">>"
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
         MICON           =   "frmUse.frx":1B99C
      End
      Begin prjChameleon.chameleonButton StepBack 
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         ToolTipText     =   "Step Backward"
         Top             =   4680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "<<"
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
         MICON           =   "frmUse.frx":1B9B8
      End
      Begin prjChameleon.chameleonButton StepBegin 
         Height          =   375
         Left            =   7440
         TabIndex        =   0
         ToolTipText     =   "Start Over"
         Top             =   4680
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "¥"
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
         MICON           =   "frmUse.frx":1B9D4
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   18
         Top             =   4800
         Width           =   525
      End
      Begin VB.Label lblCounter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   17
         Top             =   4800
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmUse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
FloatWindow frmUse.hWnd, Float

For i = 0 To 4
 Step(i).Move 18, 24
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmUse.hWnd, SINK
End Sub

Private Sub StepBack_Click()
If lblCounter <= 2 Then: StepBack.Enabled = False: StepBegin.Enabled = False
If lblCounter <= 1 Then Exit Sub
StepFor.Enabled = True
lblCounter = lblCounter - 1
Step(lblCounter - 1).Visible = True
Step(lblCounter).Visible = False
End Sub

Private Sub StepBegin_Click()
Dim i As Integer

lblCounter = 1
StepBegin.Enabled = False
StepBack.Enabled = False
If StepFor.Enabled = False Then StepFor.Enabled = True

For i = 1 To 4
 Step(i).Visible = False
Next
Step(0).Visible = True
End Sub

Private Sub StepFor_Click()
If lblCounter.Caption >= 4 Then StepFor.Enabled = False
If lblCounter.Caption >= 5 Then Exit Sub
StepBack.Enabled = True
StepBegin.Enabled = True
lblCounter = lblCounter + 1
Step(lblCounter - 1).Visible = True
Step(lblCounter - 2).Visible = False
End Sub
