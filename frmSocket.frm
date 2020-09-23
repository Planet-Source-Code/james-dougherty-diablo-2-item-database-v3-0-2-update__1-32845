VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BB31661F-0587-11D6-9DD0-00C04F0BD97C}#1.0#0"; "PRJCHAMELEON.OCX"
Begin VB.Form frmSocket 
   BorderStyle     =   0  'None
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   Icon            =   "frmSocket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   230
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPlain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   600
      Picture         =   "frmSocket.frx":000C
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   3720
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
      Left            =   1320
      Picture         =   "frmSocket.frx":0341
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   10
      Top             =   3720
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
      Left            =   1080
      Picture         =   "frmSocket.frx":06C4
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   8
      Top             =   2280
      Width           =   300
   End
   Begin VB.ComboBox cmbList 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSocket.frx":0A47
      Left            =   285
      List            =   "frmSocket.frx":0A54
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   2670
   End
   Begin prjChameleon.chameleonButton cmdOpen 
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   480
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
      MICON           =   "frmSocket.frx":0A7F
   End
   Begin prjChameleon.chameleonButton cmdSocketAll 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1050
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Socket All Items"
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
      MICON           =   "frmSocket.frx":0A9B
   End
   Begin prjChameleon.chameleonButton cmdExit 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2070
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
      MICON           =   "frmSocket.frx":0AB7
   End
   Begin prjChameleon.chameleonButton cmdAbout 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1620
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
      MICON           =   "frmSocket.frx":0AD3
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   0
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjChameleon.chameleonButton cmdSocket 
      Height          =   375
      Left            =   1095
      TabIndex        =   7
      Top             =   1470
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Socket Item"
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
      MICON           =   "frmSocket.frx":0AEF
   End
   Begin VB.Image Image5 
      Height          =   2370
      Left            =   3105
      Picture         =   "frmSocket.frx":0B0B
      Stretch         =   -1  'True
      Top             =   270
      Width           =   1365
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
      Left            =   1440
      TabIndex        =   9
      Top             =   2295
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   450
      Left            =   180
      Picture         =   "frmSocket.frx":1B94
      Stretch         =   -1  'True
      Top             =   2190
      Width           =   2880
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Item To Socket:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   285
      TabIndex        =   6
      Top             =   840
      Width           =   1365
   End
   Begin VB.Image Image3 
      Height          =   1290
      Left            =   180
      Picture         =   "frmSocket.frx":2C1D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2880
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
      TabIndex        =   13
      Top             =   2760
      Width           =   4155
   End
   Begin VB.Label lblTemp 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      Caption         =   " Socket Your Items"
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
      Left            =   225
      TabIndex        =   0
      Top             =   285
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   150
      Picture         =   "frmSocket.frx":3CA6
      Stretch         =   -1  'True
      Top             =   210
      Width           =   2880
   End
   Begin VB.Image Image2 
      Height          =   3450
      Left            =   0
      Picture         =   "frmSocket.frx":468B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4650
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Socketer v1.10 Source Code
' Written by Disk2 (disktwo@yahoo.com)
' =============================================================='
' This code is probably very inefficient and sloppy,            '
' but I released it to give any newbie game hacker              '
' an example of a simple saved game editor. This program        '
' edits items in Diablo 2 (www.blizzard.com). It sockets        '
' them.                                                         '
'                                                               '
' The program works by cycling through the items and replacing  '
' a byte with 0x08 (the marker of a socketed item :). Well,     '
' it's written in VB so it can't be TOO hard :), so here it is. '
'                                                               '
' BTW: If you use this code to make an editor, kindly give me   '
' some credit. Thanks ;)                                        '
' Also, the comments I provided might be confusing. Sorry, I'm  '
' not a good writer :)                                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This is a flag used in the Open button. It makes sure the filename selected exists...
Private Const OFN_FILEMUSTEXIST = &H1000

' The inventory is marked by a "JM  JM". It's also ended by that...
Private Type ItemHeader
    szFirstJM As String * 2 ' The first JM
    iItemCount As Byte ' This is a FAKE item count. This caused errors in v1.00 of Socketer
    iEmpty As Byte ' NULL byte
    szLastJM As String * 2 ' The final JM
End Type

' This is obviously not a complete item type :)
' This only holds the inventory position and the Equipped position.
' It's all I need for this purpose...
Private Type Item
    iSocketed As Byte
    iInvPos As Byte
End Type

' Hold the filename
Dim strFileName As String
' Used to find the beginning of the file...
Dim ItemHead As ItemHeader
Dim sFile As String
Dim sBKFile As String
Dim Checked As Boolean

Private Sub cmdExit_Click()
FloatWindow frmSocket.hWnd, SINK
Unload frmSocket
End Sub

Private Sub cmdAbout_Click()
frmMSG.lblMessage.Caption = "Socketer v1.10" & vbCrLf & "Written by Disk2. So thanks go to him"
frmMSG.Show
End Sub

Private Sub cmdOpen_Click()
    ' Set the properties...
    dlgMain.DialogTitle = "Open Character"
    dlgMain.Filter = "Diablo 2 Saved Games (*.d2s)|*.d2s|"
    dlgMain.InitDir = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\Diablo II", "Save Path")
    dlgMain.Flags = &H1000
    dlgMain.CancelError = False
    
    ' Show the window...
    dlgMain.ShowOpen
    
    ' If the user chose a file, open it and enable all the controls
    If Len(dlgMain.FileName) > 0 Then
        strFileName = dlgMain.FileName ' Set the filename
        sFile = dlgMain.FileName
        sBKFile = dlgMain.FileName & ".bak"
        cmbList.Enabled = True
        cmdSocket.Enabled = True
        cmdSocketAll.Enabled = True
    End If
End Sub

Private Sub cmdSocket_Click()
If lblTemp.Caption <> "" Then
  BackUp
End If
    ' Ok, depending on which item was chosen socket the item and show the appropriate message
    If cmbList.Text = "Left-Hand Item" Then
        Socket &H4, "Your left-hand item is now socketed."
    ElseIf cmbList.Text = "Right-Hand Item" Then
        Socket &H5, "Your right-hand item is now socketed."
    ElseIf cmbList.Text = "Helm" Then
        Socket &H1, "Your helm is now socketed."
    End If
End Sub

Private Sub BackUp()
On Error Resume Next
Kill sBKFile
FileCopy sFile, sBKFile
End Sub

Private Sub cmdSocketAll_Click()
If lblTemp.Caption <> "" Then
  BackUp
End If
Socket &H0, "All of your equipped socketable items are now socketed!"
End Sub

Private Sub Form_Load()
FloatWindow frmSocket.hWnd, Float
cmbList.ListIndex = 0
Checked = True
lblTemp.Caption = 0
End Sub

Private Sub Socket(Position As Integer, Message As String)
On Error Resume Next ' If we encounter an error, resume next :)
    
Dim iPos As Integer ' IMPORTANT: This holds our position in the file...
Dim xItem As Item ' The temp item. Used to compare item positions, etc...
Dim TheEnd As ItemHeader ' Used to check if we're at the end of the inventory
Dim TheString As String * 4 ' Should be "JMJM" if we're at the end of the inventory

    ' See declaration of ItemHeader type (at top of code)
TheEnd.iEmpty = &H0
TheEnd.iItemCount = &H0
TheEnd.szFirstJM = ""
TheEnd.szLastJM = ""

    ' Clear ItemHead (it's a global variable so we need to clear it each time...)
ItemHead.szFirstJM = ""
ItemHead.szLastJM = ""
ItemHead.iItemCount = 0
ItemHead.iEmpty = 0

    ' Start at the beginning of the file
    iPos = &H1

    ' Open the filename (strFileName)
    Open strFileName For Binary As #1
        ' Get the position of the start of the inventory data
        Do Until ItemHead.szFirstJM = "JM" And ItemHead.szLastJM = "JM"
            Get #1, iPos, ItemHead
            
            iPos = iPos + 1
        Loop
    
        ' OK. We found it. Now we have to increase our position by 3 to get to the first item
        iPos = iPos + 3

        ' If the item count is zero then there's no point in continuing :)
        If ItemHead.iItemCount = 0 Then
            MsgBox "This character doesn't appear to have any items! If this is an error please email me at cregistry@yahoo.com and attach the saved game file. Thanks!", vbOKOnly + vbInformation, "Notice"
            
            ' Close the file and exit the sub...
            Close #1
            Exit Sub
        End If

        ' The ItemHead.iItemCount is a fake value for the number of items (I guess).
        ' The number doesn't account for gems that are in socketed items.
        ' So now we have to read items until we find the closing "JM  JM" in the file...
        Do Until TheString = "JMJM"
            ' First of all, make sure we aren't at the end of the inventory
            Get #1, iPos, TheEnd
            ' If TheString equals "JMJM" then we are at the end
            TheString = TheEnd.szFirstJM & TheEnd.szLastJM
            
            ' Increase our position by 2 to get to the item data
            ' BTW: Each item is 25 bytes long...
            iPos = iPos + 2
            
            ' Read the position of the item.
            Get #1, iPos + 4, xItem.iInvPos
            
            ' Depending on the value of Position when the function was called,
            ' socket the appropriate item(s).
            ' BTW: &H0 means socket all the items that are equipped...
            Select Case Position
                Case &H0
                    If xItem.iInvPos = &H1 Or xItem.iInvPos = &H4 Or xItem.iInvPos = &H5 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H1
                    If xItem.iInvPos = &H1 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H4
                    If xItem.iInvPos = &H4 Then
                        Put #1, iPos + 1, &H8
                    End If
                Case &H5
                    If xItem.iInvPos = &H5 Then
                        Put #1, iPos + 1, &H8
                    End If
            End Select
            
            ' Increase the position by 25 so that we can read the next item...
            iPos = iPos + 25
            
            ' Then loop back to the beginning :)
        Loop
        
    ' Close the file
    Close #1
    
    ' Show the message (this is shown regardless of wether the item was socketed
    ' successfully :)
    frmMSG.lblMessage.Caption = "Done! " & vbCrLf & Message
    frmMSG.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
FloatWindow frmSocket.hWnd, SINK
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
