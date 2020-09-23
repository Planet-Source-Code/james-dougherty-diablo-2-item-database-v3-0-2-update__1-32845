VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Account"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExpand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Expand All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      Picture         =   "frmDocument.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   240
      Width           =   900
   End
   Begin VB.CommandButton cmdCollapse 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Collapse All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      Picture         =   "frmDocument.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   900
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   6240
      Picture         =   "frmDocument.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete Selected Item"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   6240
      Picture         =   "frmDocument.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   900
   End
   Begin VB.CommandButton cmdCollapseSel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Collapse Selected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      Picture         =   "frmDocument.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   900
   End
   Begin VB.CommandButton cmdExpandSel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Expand Selected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   0
      Picture         =   "frmDocument.frx":1374
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   900
   End
   Begin VB.Timer SizeTimer 
      Interval        =   1
      Left            =   2880
      Top             =   2400
   End
   Begin MSComctlLib.TreeView DView 
      Height          =   6375
      Left            =   480
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11245
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgBack 
      Height          =   9000
      Left            =   7200
      Picture         =   "frmDocument.frx":167E
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   9000
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCollapse_Click()
On Local Error Resume Next
Dim i As Long

For i = 1 To DView.Nodes.Count
 If DView.Nodes(i).Children > 0 Then DView.Nodes(i).Expanded = False
Next
frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & "Database Collapsed..."
End Sub

Private Sub cmdCollapse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " Collapse every node"
End Sub

Private Sub cmdCollapseSel_Click()
On Local Error Resume Next
DView.SelectedItem.Expanded = False
frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & DView.SelectedItem.Text & " Collapsed..."
End Sub

Private Sub cmdCollapseSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " Collapse a selected node"
End Sub

Private Sub cmdDelete_Click()
On Local Error Resume Next

If frmMain.cmdAddItem(1).Enabled = False Then
 frmMSG.lblMessage.Caption = "Please create a character."
 frmMSG.Show 1
 Exit Sub
End If

If DView.SelectedItem.Text = "" Then
 frmMSG.lblMessage.Caption = "Please select an item to delete."
 frmMSG.Show 1
 Exit Sub
End If

If DView.SelectedItem.Text = AName Or _
   DView.SelectedItem.Text = CharName Or _
   DView.SelectedItem.Text = "Jewelry" Or _
   DView.SelectedItem.Text = "Armor" Or _
   DView.SelectedItem.Text = "Weapons" Or _
   DView.SelectedItem.Text = "Jewels/Runes/Gems" Or _
   DView.SelectedItem.Text = "Charms" Then
 frmMSG.lblMessage.Caption = "Can not delete predefined nodes." & vbCrLf & "Only items you have inserted."
 frmMSG.Show 1
 Exit Sub
End If

frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & DView.SelectedItem.Text & " was deleted..."
DView.Nodes.Remove (DView.SelectedItem.Index)
End Sub

Private Sub cmdDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " Delete a selected node"
End Sub

Private Sub cmdExpand_Click()
On Local Error Resume Next
Dim i As Long

For i = 1 To DView.Nodes.Count
 If DView.Nodes(i).Children > 0 Then DView.Nodes(i).Expanded = True
Next
frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & "Database Expanded..."
End Sub

Private Sub cmdExpand_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " Expand every node"
End Sub

Private Sub cmdExpandSel_Click()
On Local Error Resume Next
DView.SelectedItem.Expanded = True
frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & DView.SelectedItem.Text & " Expanded..."
End Sub

Private Sub cmdExpandSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " Expand a selected node"
End Sub

Private Sub cmdRefresh_Click()
On Local Error Resume Next
DView.Refresh
frmMain.txtOutput.Text = frmMain.txtOutput.Text & vbCrLf & "Database Refreshed..."
End Sub

Private Sub cmdRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " Refresh the database"
End Sub

Private Sub DView_Click()
Debug.Print DView.SelectedItem.Key
End Sub

Private Sub DView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " "
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " "
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
imgBack.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
DView.Left = Int((Me.ScaleWidth - DView.Width) / 2)
DView.Top = Int((Me.ScaleHeight - DView.Height) / 2)
cmdExpand.Move (DView.Left - cmdExpand.Width) - 5, DView.Top
cmdCollapse.Move (DView.Left - cmdCollapse.Width) - 5, (cmdExpand.Top + cmdExpand.Height) + 15
cmdExpandSel.Move (DView.Left - cmdExpandSel.Width) - 5, (cmdCollapse.Top + cmdCollapse.Height) + 30
cmdCollapseSel.Move (DView.Left - cmdCollapseSel.Width) - 5, (cmdExpandSel.Top + cmdExpandSel.Height) + 15
cmdRefresh.Move (DView.Left + DView.Width) + 5, DView.Top
cmdDelete.Move (DView.Left + DView.Width) + 5, (cmdRefresh.Top + cmdRefresh.Height) + 15
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.txtAName.Text = Me.Caption
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.lblStatus.Caption = " "
End Sub
