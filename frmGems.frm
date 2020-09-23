VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmGems 
   BorderStyle     =   0  'None
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   Icon            =   "frmGems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picChecked 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   0
      Picture         =   "frmGems.frx":000C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   12
      Top             =   3480
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
      Picture         =   "frmGems.frx":038F
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
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
      Left            =   3405
      Picture         =   "frmGems.frx":06C4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   10
      Top             =   1980
      Width           =   300
   End
   Begin VB.ComboBox cmbDest 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGems.frx":0A47
      Left            =   360
      List            =   "frmGems.frx":0A60
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2280
      Width           =   1440
   End
   Begin VB.ComboBox cmbSource 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmGems.frx":0AA6
      Left            =   360
      List            =   "frmGems.frx":0AB0
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1665
      Width           =   1440
   End
   Begin prjChameleon.chameleonButton cmdConvert 
      Height          =   495
      Left            =   300
      TabIndex        =   5
      Top             =   840
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Convert Potions Into Gems"
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
      MICON           =   "frmGems.frx":0AD2
   End
   Begin prjChameleon.chameleonButton cmdOpen 
      Height          =   495
      Left            =   3375
      TabIndex        =   1
      Top             =   900
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Open Character"
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
      MICON           =   "frmGems.frx":0AEE
   End
   Begin prjChameleon.chameleonButton cmdExit 
      Height          =   375
      Left            =   3375
      TabIndex        =   2
      Top             =   2280
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
      MICON           =   "frmGems.frx":0B0A
   End
   Begin prjChameleon.chameleonButton cmdAbout 
      Height          =   375
      Left            =   3375
      TabIndex        =   3
      Top             =   1440
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
      MICON           =   "frmGems.frx":0B26
   End
   Begin prjChameleon.chameleonButton cmdUpgrade 
      Height          =   615
      Left            =   1995
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Upgrade All Gems To Perfect"
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
      MICON           =   "frmGems.frx":0B42
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   32000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup"
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
      Left            =   3795
      TabIndex        =   13
      Top             =   1980
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   2115
      Left            =   3240
      Picture         =   "frmGems.frx":0B5E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1365
   End
   Begin VB.Image Image5 
      Height          =   2115
      Left            =   1920
      Picture         =   "frmGems.frx":1BE7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Note: This freezes the program on my computer so I'm not sure if it works. Suppose to?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   4155
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Type:"
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
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   2055
      Width           =   1425
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Type:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Image Image3 
      Height          =   2115
      Left            =   240
      Picture         =   "frmGems.frx":2C70
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Gem Hack"
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
      Top             =   315
      Width           =   4350
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   150
      Picture         =   "frmGems.frx":3C21
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4560
   End
   Begin VB.Image Image1 
      Height          =   3585
      Left            =   0
      Picture         =   "frmGems.frx":4606
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4845
   End
End
Attribute VB_Name = "frmGems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GEMEDIT v1.10 SOURCE CODE                                                       '
' Written by Disk2 (disktwo@yahoo.com)                                            '
'                                                                                 '
' This is the source code to the GemEdit program I made for Diablo 2.             '
' It's nothing fancy, but it works. I've tried to comment the code fairly well.   '
' If you have any questions, PLEASE don't email me :) Figure it out for yourself. '
' I don't have the time to help, and I think you learn better by trying.          '
'                                                                                 '
' I released this code to help anyone who wants to make an editor. I shows        '
' basic binary I/O. It's a start :)                                               '
'                                                                                 '
' If you use this code, I request a mention somewhere in the editor :) Thanks...  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Used to find the beginning and end of the inventory section...
Private Type ItemHeader
    szFirstJM As String * 2
    iItemCount As Byte
    iEmpty As Byte
    szLastJM As String * 2
End Type

' This isn't complete :) It does what I need it to though.
Private Type Item
    iSubType As Byte
    iType As Byte
End Type

' Holds the filename and declares the item header...
Dim strFileName As String
Dim ItemHead As ItemHeader
Private Checked As Boolean
Dim sFile As String
Dim sBKFile As String

Private Sub cmdAbout_Click()
frmMSG.lblMessage.Caption = "Socketer v1.10" & vbCrLf & "Written by Disk2. So thanks go to him"
frmMSG.Show
End Sub

Private Sub BackUp()
On Error Resume Next
Kill sBKFile
FileCopy sFile, sBKFile
End Sub

Private Sub cmdExit_Click()
Unload frmGems
End Sub

Private Sub cmdConvert_Click()
' Check to see if the user select a source and destination type...
If lblTemp.Caption <> "" Then
  BackUp
End If
If cmbSource.Text = "" Or cmbDest.Text = "" Then
 MsgBox "You must select a source and destination type before converting!", vbCritical + vbOKOnly, "Error"
 frmMSG.lblMessage.Caption = "You must select a source and destination" & vbCrLf & "type before converting!"
 frmMSG.Show
Else
 ' If the user DID, then convert the items.
 Convert
End If
End Sub

Private Sub cmdOpen_Click()
    ' Set the common dialog properties
    dlgMain.Flags = &H1000
    dlgMain.InitDir = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\Diablo II\", "Save Path")
    dlgMain.DialogTitle = "Open Character"
    dlgMain.Filter = "Diablo 2 Saved Games (*.d2s)|*.d2s|"
    dlgMain.CancelError = False
    dlgMain.MaxFileSize = 32000
    dlgMain.ShowOpen
    
    ' If the user selected a file, open it and enable the controls...
    If Len(dlgMain.FileName) > 0 Then
        strFileName = dlgMain.FileName
        sFile = dlgMain.FileName
        sBKFile = dlgMain.FileName & ".bak"
        cmbSource.Enabled = True
        cmbDest.Enabled = True
        cmdConvert.Enabled = True
        cmdUpgrade.Enabled = True
    End If
End Sub

Private Sub cmdUpgrade_Click()
' Fix the user's gems... This is done in the FixGems sub
If lblTemp.Caption <> "" Then
  BackUp
End If
FixGems
End Sub

Private Sub FixGems()
    On Local Error Resume Next
    
    Dim iPos As Integer ' Holds the position in the file
    Dim xItem As Item ' Temp item
    Dim TheString As String * 4 ' will be JMJM if we're at the end of the items...
    Dim TheEnd As ItemHeader ' Used to find the end of the items

    ' Reset ItemHead
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    iPos = &H1 ' Start at the beginning of the file

    ' Open the file
    Open strFileName For Binary As #1
        ' Read from the file until we find the "JM  JM". This means we've found the beginning of the item data...
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1 ' Increase the position
        Loop
    
        iPos = iPos + 3 ' Go to the REAL start of the item information.

        ' If the user has no items, don't continue...
        If ItemHead.iItemCount = 0 Then
            frmMSG.lblMessage.Caption = "This character doesn't appear" & vbCrLf & "to have any items!"
            frmMSG.Show
            Close #1 ' Close the file
            Exit Sub
        End If

        ' Read items. Compare them with known gem codes. Convert them if they aren't perfect.
        Do Until TheString = "JMJM"
            Get #1, iPos, TheEnd
            
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            iPos = iPos + 2
        
            Get #1, iPos + 6, xItem.iSubType
            Get #1, iPos + 7, xItem.iType
            
            ' Hehe. This is confusing.
            ' The actual concept isn't. It's just the way I coded it :)
            ' Here's a table of all the gem codes...
            '
            '          | Chipped | Flawed | Regular | Flawless | Perfect |
            ' ---------|--------------------------------------------------
            ' Diamond  | 5015    | 6015   | 7015    | 8015     | 9015    |
            ' ---------|--------------------------------------------------
            ' Ruby     | 1015    | 0015   | 2015    | 3015     | 4015    |
            ' ---------|--------------------------------------------------
            ' Topaz    | 1014    | 2014   | 3014    | 4014     | 5014    |
            ' ---------|--------------------------------------------------
            ' Sapphire | 6014    | 7014   | 8014    | 9014     | A014    |
            ' ---------|--------------------------------------------------
            ' Amethyst | C013    | D013   | E013    | F013     | 0014    |
            ' ---------|--------------------------------------------------
            ' Emerald  | B014    | C014   | D014    | E014     | F014    |
            ' ---------|--------------------------------------------------
            ' Skull    | 4016    | 5016   | 6016    | 7016     | 8016    |
            ' ---------|--------------------------------------------------
            
            ' With that in mind, you can figure out this code.
            Select Case xItem.iType
                Case &H13
                    Select Case xItem.iSubType
                        Case &HD0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HC0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HF0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                        Case &HE0
                            Put #1, iPos + 6, &H0
                            Put #1, iPos + 7, &H14
                    End Select
                Case &H14
                    Select Case xItem.iSubType
                        Case &H20
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H10
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H40
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H30
                            Put #1, iPos + 6, &H50
                            Put #1, iPos + 7, &H14
                        Case &H60
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H70
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H80
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &H90
                            Put #1, iPos + 6, &HA0
                            Put #1, iPos + 7, &H14
                        Case &HC0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HB0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HE0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                        Case &HD0
                            Put #1, iPos + 6, &HF0
                            Put #1, iPos + 7, &H14
                    End Select
                    Case &H15
                    Select Case xItem.iSubType
                        Case &H60
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H50
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H80
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H70
                            Put #1, iPos + 6, &H90
                            Put #1, iPos + 7, &H15
                        Case &H0
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H10
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H20
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                        Case &H30
                            Put #1, iPos + 6, &H40
                            Put #1, iPos + 7, &H15
                    End Select
                    Case &H16
                    Select Case xItem.iSubType
                        Case &H40
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H50
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H60
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                        Case &H70
                            Put #1, iPos + 6, &H80
                            Put #1, iPos + 7, &H16
                    End Select
            End Select
            
            ' Increase the position so we can read the next item.
            iPos = iPos + 25
        Loop
    Close #1
    
    ' Tell the user the gems were perfected
    frmMSG.lblMessage.Caption = "All your gems are now perfect."
    frmMSG.Show
End Sub

Private Sub Convert()
    On Local Error Resume Next

    Dim iPos As Integer
    Dim xItem As Item
    Dim dItem As Item
    Dim TheEnd As ItemHeader
    Dim TheString As String * 4
    
    ' Depending of the destination type, set the temp item type to a perfect gem.
    ' Refer to the table in FixGems for the gem codes...
    Select Case cmbDest.Text
        Case "Diamonds"
            dItem.iType = &H15
            dItem.iSubType = &H90
        Case "Rubys"
            dItem.iType = &H15
            dItem.iSubType = &H40
        Case "Topazes"
            dItem.iType = &H14
            dItem.iSubType = &H50
        Case "Sapphires"
            dItem.iType = &H14
            dItem.iSubType = &HA0
        Case "Amethysts"
            dItem.iType = &H14
            dItem.iSubType = &H0
        Case "Emeralds"
            dItem.iType = &H14
            dItem.iSubType = &HF0
        Case "Skulls"
            dItem.iType = &H16
            dItem.iSubType = &H80
    End Select

    ' Reset ItemHead
    ItemHead.szFirstJM = ""
    ItemHead.szLastJM = ""
    ItemHead.iItemCount = 0
    ItemHead.iEmpty = 0

    ' Start at the beginning of the file
    iPos = &H1

    ' Open the file
    Open strFileName For Binary As #1
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1
        Loop
    
         ' Go to the REAL item data start
        iPos = iPos + 3

        ' If the user has no items, there's no point in continuing.
        If ItemHead.iItemCount = 0 Then
            frmMSG.lblMessage.Caption = "This character doesn't appear" & vbCrLf & "to have any items!"
            frmMSG.Show
            Close #1
            Exit Sub
        End If

        ' Read items until we reach the end of the file.
        Do Until TheString = "JMJM"
            Get #1, iPos, TheEnd
            
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            iPos = iPos + 2
            
            ' Get the item type
            Get #1, iPos + 6, xItem.iSubType
            Get #1, iPos + 7, xItem.iType
            
            ' Depending on the source type, look for health or mana potions.
            ' You can figure out the potion codes by looking at the code below.
            If cmbSource.Text = "Health Potions" Then
                If xItem.iType = &H15 Then
                    If xItem.iSubType = &HA0 Or xItem.iSubType = &HB0 Or xItem.iSubType = &HC0 Or xItem.iSubType = &HD0 Or xItem.iSubType = &HE0 Then
                        Put #1, iPos + 6, dItem
                    End If
                End If
            End If
            If cmbSource.Text = "Mana Potions" Then
                If xItem.iType = &H16 Then
                    If xItem.iSubType = &H0 Or xItem.iSubType = &H10 Or xItem.iSubType = &H20 Or xItem.iSubType = &H30 Then
                        Put #1, iPos + 6, dItem
                    End If
                ElseIf xItem.iType = &H15 Then
                    If xItem.iSubType = &HF0 Then
                        Put #1, iPos + 6, dItem
                    End If
                End If
            End If
            
            ' Increase the position so we can read the next file
            iPos = iPos + 25
        Loop
    Close #1

    ' Show the message...
    frmMSG.lblMessage.Caption = "All your " & cmbSource.Text & " are now " & cmbDest.Text & "."
    frmMSG.Show
End Sub

Private Sub Form_Load()
cmbSource.ListIndex = 0
cmbDest.ListIndex = 0
FloatWindow frmGems.hWnd, Float
Checked = True
lblTemp.Caption = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmGems.hWnd, SINK
End Sub

Private Sub picCheck_Click(Index As Integer)
If Checked Then
 lblTemp.Caption = ""
 picCheck(0).Picture = picPlain.Picture
 Checked = False
Else
 lblTemp.Caption = 0
 picCheck(0).Picture = picChecked.Picture
 Checked = True
End If
End Sub
