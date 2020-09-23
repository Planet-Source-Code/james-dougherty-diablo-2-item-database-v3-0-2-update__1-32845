Attribute VB_Name = "mMRU"
Option Explicit

Private Const Max_File_Length As Integer = 35

Public MRUpath() As String
Public MRUlimit As Integer
Public FileMRU1 As String
Public FileMRU2 As String
Public FileMRU3 As String
Public FileMRU4 As String
Public FileName As String
Public Modified As Boolean
Private m_FileName As String
Private m_FileTitle As String
Private Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

Private Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40

Public Sub LoadRescentFiles()
Dim Readln As String
Dim StrName As String
Dim StrName2 As String
Dim StrName3 As String
Dim StrName4 As String
On Error GoTo 10
FileMRU1 = ""
FileMRU2 = ""
FileMRU3 = ""
FileMRU4 = ""
Open App.Path & "\Resource.ini" For Input As #1
    Line Input #1, Readln
    FileMRU1 = Readln
    Line Input #1, Readln
    FileMRU2 = Readln
    Line Input #1, Readln
    FileMRU3 = Readln
    Line Input #1, Readln
    FileMRU4 = Readln
10:
    Close #1
    
    If FileMRU1 <> "" Then
        frmMain.mnuFile1.Visible = True
        frmMain.mnuFile1.Tag = FileMRU1
        If Len(FileMRU1) > Max_File_Length Then
         StrName = Left(FileMRU1, 3) & "..." & Right(FileMRU1, Max_File_Length - 10)
        Else
         StrName = FileMRU1
        End If
        frmMain.mnuFile1.Caption = "&1: " & StrName
        frmMain.mnuFileBar4.Visible = True
        frmMain.mnuFileDummy.Visible = True
    Else
        frmMain.mnuFile1.Visible = False
        frmMain.mnuFile1.Caption = ""
        frmMain.mnuFileBar4.Visible = False
        frmMain.mnuFileDummy.Visible = False
    End If
    
    If FileMRU2 <> "" Then
        frmMain.mnuFile2.Visible = True
        frmMain.mnuFile2.Tag = FileMRU2
        If Len(FileMRU2) > Max_File_Length Then
         StrName2 = Left(FileMRU2, 3) & "..." & Right(FileMRU2, Max_File_Length - 10)
        Else
         StrName2 = FileMRU2
        End If
        frmMain.mnuFile2.Caption = "&2: " & StrName2
    Else
        frmMain.mnuFile2.Visible = False
        frmMain.mnuFile2.Caption = ""
    End If
    
    If FileMRU3 <> "" Then
        frmMain.mnuFile3.Visible = True
        frmMain.mnuFile3.Tag = FileMRU3
        If Len(FileMRU3) > Max_File_Length Then
         StrName3 = Left(FileMRU3, 3) & "..." & Right(FileMRU3, Max_File_Length - 10)
        Else
         StrName3 = FileMRU3
        End If
        frmMain.mnuFile3.Caption = "&3: " & StrName3
    Else
        frmMain.mnuFile3.Visible = False
        frmMain.mnuFile3.Caption = ""
    End If

    If FileMRU4 <> "" Then
        frmMain.mnuFile4.Visible = True
        frmMain.mnuFile4.Tag = FileMRU4
        If Len(FileMRU4) > Max_File_Length Then
         StrName4 = Left(FileMRU4, 3) & "..." & Right(FileMRU4, Max_File_Length - 10)
        Else
         StrName4 = FileMRU4
        End If
        frmMain.mnuFile4.Caption = "&4: " & StrName4
    Else
        frmMain.mnuFile4.Visible = False
        frmMain.mnuFile4.Caption = ""
    End If
End Sub

Public Sub AddToList(FileName As String)
    FileMRU4 = FileMRU3
    FileMRU3 = FileMRU2
    FileMRU2 = FileMRU1
    FileMRU1 = FileName
    Open App.Path & "\Resource.ini" For Output As #1
        Print #1, FileMRU1
        If FileMRU2 <> "" Then Print #1, FileMRU2
        If FileMRU3 <> "" Then Print #1, FileMRU3
        If FileMRU4 <> "" Then Print #1, FileMRU4
    Close #1
    LoadRescentFiles
End Sub

Public Function GetFTitle(strFileName As String)
On Error GoTo GFTError
Dim cbBuf As String
    
cbBuf = String(250, vbNullChar)
GetFileTitle strFileName, cbBuf, Len(cbBuf)
GetFTitle = Left(cbBuf, InStr(1, cbBuf, vbNullChar) - 1)
GFTError:
End Function
