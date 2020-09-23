Attribute VB_Name = "mRegistery"
Option Explicit

Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or READ_CONTROL
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_SUB_KEY Or KEY_CREATE_LINK Or KEY_SET_VALUE
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_OPTION_VOLATILE = 1
Public Const KEY_EXECUTE = KEY_READ
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_DWORD = 4
Public Const ERROR_NONE = 0
Public Const ERROR_BADKEY = 2
Public Const ERROR_ACCESS_DENIED = 8
Public Const ERROR_SUCCESS = 0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public cEnumValues As Collection
Public cEnumKeys As Collection

Public CurGold As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Function RGCreateKey(hKey As Long, SubKey As String)
Dim lngRet As Long
Dim lngResult As Long
Dim lngDis As Long
Dim tmp As SECURITY_ATTRIBUTES
    
tmp.nLength = 50
tmp.lpSecurityDescriptor = 0
tmp.bInheritHandle = True
    
lngRet = RegCreateKeyEx(hKey, SubKey, 0&, 0&, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, tmp, lngResult, lngDis)
lngRet = RegCloseKey(lngResult)
End Function

Public Function RGDeleteKey(hKey As Long, SubKey As String)
RegDeleteKey hKey, SubKey
End Function

Public Function RGSetKeyValue(hKey As Long, SubKey As String, ValueName As String, sValue As String)
Dim lngRet As Long
Dim lngResult As Long
    
lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
If lngRet = ERROR_SUCCESS Then
 RegSetValueEx lngResult, ValueName, 0, REG_SZ, ByVal sValue, Len(sValue)
 RegFlushKey lngResult
 RegCloseKey lngResult
End If
End Function

Public Function RGGetKeyValue(hKey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
Dim lngRet As Long
Dim lngResult As Long
Dim sData As String
    
lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
If lngRet = ERROR_SUCCESS Then
 sData = String(128, vbNullChar)
 lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
 If Not lngRet = ERROR_SUCCESS Then RGGetKeyValue = Default: Exit Function
 RGGetKeyValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
 RegCloseKey lngResult
Else
 RGGetKeyValue = Default
End If
End Function

Public Function RGDeleteKeyValue(hKey As Long, SubKey As String, ValueName As String)
Dim lngRet As Long
Dim lngResult As Long
       
lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
RegDeleteValue lngResult, ValueName
End Function

Public Function RGEnumKeyValues(hKey As Long, SubKey As String)
Dim lngRet As Long
Dim lngResult As Long
Dim sData As String
Dim intIndex As Integer

lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
If lngRet = ERROR_SUCCESS Then
 Set cEnumValues = New Collection
 Do
  sData = String(128, vbNullChar)
  lngRet = RegEnumValue(lngResult, intIndex, sData, Len(sData), 0, ByVal 0&, ByVal 0&, ByVal 0&)
  If lngRet <> 0 Then Exit Do
  cEnumValues.Add Left(sData, InStr(1, sData, vbNullChar) - 1)
  intIndex = intIndex + 1
 Loop
 RegCloseKey lngResult
End If
End Function

Public Function RGEnumKeys(hKey As Long, SubKey As String)
Dim lngRet As Long
Dim lngResult As Long
Dim sData As String
Dim intIndex As Integer

lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
If lngRet = ERROR_SUCCESS Then
Set cEnumKeys = New Collection
Do
  sData = String(128, vbNullChar)
  lngRet = RegEnumKey(lngResult, intIndex, sData, Len(sData))
  If lngRet <> 0 Then Exit Do
  cEnumKeys.Add Left(sData, InStr(1, sData, vbNullChar) - 1)
  intIndex = intIndex + 1
 Loop
 RegCloseKey lngResult
End If
End Function

Public Function Associate(Program As String, Extension As String, Description As String, Optional Icon As String)
RGCreateKey HKEY_CLASSES_ROOT, "." & Extension
RGSetKeyValue HKEY_CLASSES_ROOT, "." & Extension, "", Extension & "file"
    
RGCreateKey HKEY_CLASSES_ROOT, Extension & "file"
RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell"

RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open"
RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\shell\open\command"
RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\shell\open\command", "", Program & " " & "%1"

RGCreateKey HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon"
    
RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file", "", Description
RGSetKeyValue HKEY_CLASSES_ROOT, Extension & "file\DefaultIcon", "", Icon
End Function

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    On Local Error Resume Next
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim r As Long
    Dim lValueType As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        sKeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = sKeyVal                                   ' Return Value
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type
    
    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...
    
    '------------------------------------------------------------
    '- Create/Modify Key Value...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' A Space Is Needed For RegSetValueEx() To Work...
    
    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
    '------------------------------------------------------------
    '- Close Registry Key...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Close Key
    
    UpdateKey = True                                    ' Return Success
    Exit Function                                       ' Exit
CreateKeyError:
    UpdateKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
End Function

