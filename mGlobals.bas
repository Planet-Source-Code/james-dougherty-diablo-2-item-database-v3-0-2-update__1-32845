Attribute VB_Name = "mGlobals"
Option Explicit

Public Enum FloatType
 Float = 0
 SINK = 1
End Enum

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public OK As Boolean
Public AName As String
Public CharOK As Boolean
Public CharName As String
Public JewelryOK As Boolean
Public JewelryName As String
Public HoldOPValue As Long
Public ArmorOK As Boolean
Public ArmorName As String
Public HoldOPValue2 As Long
Public WeaponOK As Boolean
Public WeaponName As String
Public HoldOPValue3 As Long
Public JewelOK As Boolean
Public JewelName As String
Public CharmOK As Boolean
Public CharmName As String
Public FSys As New FileSystemObject
Public OutStream As TextStream
Public InStream As TextStream
Public HoldFileTitle As String
Public tmpFileName As String
Public FromFile As Boolean
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Const MF_BYPOSITION = &H400&

Public Sub FloatWindow(hWnd As Long, Action As FloatType)
Dim wFlags As Integer, result As Integer

wFlags = SWP_NOMOVE Or SWP_NOSIZE

If Action = Float Then
    Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags)
Else
    Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags)
End If

End Sub
