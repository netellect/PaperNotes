Attribute VB_Name = "modFileOps"
Option Explicit
'----------------------------------------------------------------------------------
'       (c) 2000-2003 by Vlad Kozin
'==================================================================================

Public Const MAX_PATH = 1024&
Public Const MAXDWORD = &HFFFF
Public Const MAX_ARGS = 10
Public Const INVALID_HANDLE_VALUE = -1&
Public Const MOVEFILE_COPY_ALLOWED = &H2&
Public Const MOVEFILE_DELAY_UNTIL_REBOOT = &H4&
Public Const MOVEFILE_REPLACE_EXISTING = &H1&

Public Const MOVEFILE_WRITE_THROUGH = &H8&

Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" _
    (ByVal hInst As Long, _
    ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" _
    (ByVal hInst As Long, _
    ByVal lpIconPath As String, _
    lpiIcon As Long) As Long

Public gFindFirstInPathProductName As String

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, _
    lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, _
    lpFindFileData As WIN32_FIND_DATA) As Long
    
Public Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
    
    
' crap. the following function is not supported in Windows 98 !!!!
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" _
    (lpExistingFileName As String, _
    lpNewFileName As String, _
    ByVal dwFlags As Long) As Long  'was ByVal before
    
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" _
    (ByVal lpExistingFileName As String, _
    ByVal lpNewFileName As String) As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" _
    (ByVal lpPathName As String, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
    (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" _
    (ByVal lpSectionName As String, _
    ByVal lpKeyName As String, _
    ByVal nDefault As Long, _
    ByVal lpINIFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpSectionName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpINIFileName As String) As Long

Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" _
    (ByVal lpSectionName As String, _
    ByVal lpString As String, _
    ByVal lpINIFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpSectionName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpINIFileName As String) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'Registry Functions
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long         'if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Constants for Windows 32-bit Registry API
Private Enum enumHKEY_CLASS_LOCAL
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
'Reg Key Security Options
Private Const DELETE = &H10000
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))

'registry entry types
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_RESOURCE_LIST = 8
'Private Const REG_OPTION_NON_VOLATILE = 0

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234&
Private Const ERROR_NO_MORE_ITEMS = 259&

' from modEncryption
Private Const C0DEBASE As Byte = &H5F   ' Do not change.

Public Function FileSize(MaskOrName As String) As Long
    Dim retdata As WIN32_FIND_DATA
    Dim handle As Long
    Dim fsize As Long
    fsize = 0
    handle = FindFirstFile(MaskOrName, retdata)
    If handle <> INVALID_HANDLE_VALUE Then
        fsize = (retdata.nFileSizeHigh * (MAXDWORD + 1)) + retdata.nFileSizeLow
        handle = FindClose(handle)
   End If
    FileSize = fsize
End Function

Public Function FileIsReadOnly(MaskOrName As String) As Boolean
    Dim retdata As WIN32_FIND_DATA
    Dim handle As Long
    FileIsReadOnly = False
    handle = FindFirstFile(MaskOrName, retdata)
    If handle <> INVALID_HANDLE_VALUE Then
        If retdata.dwFileAttributes = (retdata.dwFileAttributes Or FILE_ATTRIBUTE_READONLY) Then
            FileIsReadOnly = True
        End If
        handle = FindClose(handle)
    End If
End Function

Public Function FileDateTime(MaskOrName As String) As FILETIME
    Dim retdata As WIN32_FIND_DATA
    Dim handle As Long
    handle = FindFirstFile(MaskOrName, retdata)
    If handle = INVALID_HANDLE_VALUE Then
        FileDateTime.dwHighDateTime = 0
        FileDateTime.dwLowDateTime = 0
    Else
        FileDateTime = retdata.ftLastWriteTime
        handle = FindClose(handle)
    End If
End Function

Public Function FileExists(MaskOrName As String) As Boolean
    Dim retdata As WIN32_FIND_DATA
    Dim handle As Long
    handle = FindFirstFile(MaskOrName, retdata)
    If handle = INVALID_HANDLE_VALUE Then
        FileExists = False
    Else
        FileExists = True
        handle = FindClose(handle)
    End If
End Function
    
Public Function IsDirectory(MaskOrName As String) As Boolean
    Dim retdata As WIN32_FIND_DATA
    Dim handle As Long
    
    IsDirectory = False
    handle = FindFirstFile(MaskOrName, retdata)
    If handle <> INVALID_HANDLE_VALUE Then
        If ((retdata.dwFileAttributes Or FILE_ATTRIBUTE_DIRECTORY) = retdata.dwFileAttributes) Then
            IsDirectory = True
        End If
        handle = FindClose(handle)
    End If
End Function

Public Function FileDelete(FileName As String) As Boolean
    If DeleteFile(FileName) <> 0 Then
        FileDelete = True
    Else
        FileDelete = False
    End If
End Function

Public Function CreateDir(dirname As String) As Boolean
    Dim a As SECURITY_ATTRIBUTES
    a.bInheritHandle = 0
    a.lpSecurityDescriptor = 0
    a.nLength = 0
    If CreateDirectory(dirname, a) <> 0 Then
        CreateDir = True
    Else
        CreateDir = False
    End If
End Function

Public Function TakeFilename(ByVal fullname As String, Optional Separator As String = "\", Optional IncludeExtension As Boolean = True) As String
    Dim pos As Long
    Dim posext As Long
    Dim tmpStr As String
    pos = Len(fullname) - InStrRev(fullname, Separator)
    tmpStr = VBA.Right(fullname, pos)
    If Not IncludeExtension Then
        posext = InStrRev(tmpStr, ".")
        tmpStr = VBA.Left(tmpStr, posext - 1)
    End If
    TakeFilename = tmpStr
End Function

Public Function TakePath(ByVal fullname As String, Optional Separator As String = "\") As String
    'make SURE THE FILENAME IS INCLUDED, otherwise the function will work incorrectly
    Dim pos As Long
    pos = InStrRev(fullname, Separator)
    TakePath = VBA.Left(fullname, IIf(pos > 0, pos - 1, 0))
End Function

Public Function TakeNthDirName(ByVal fullname As String, Nth As Long, Optional Separator As String = "\") As String
    Dim pos As Long
    Dim prevpos As Long
    Dim tmpStr As String
    Dim tmpstrLeft As String
    Dim k As Long
    pos = 1
    prevpos = 0
    TakeNthDirName = ""
    If Nth > 10 Then Exit Function
    For k = 1 To Nth
        prevpos = pos
        pos = InStr(pos, fullname, Separator)
        If pos = 0 Then
            Exit For
        End If
    Next
    If pos < prevpos Then Exit Function
    tmpStr = VBA.Mid(fullname, prevpos, pos - prevpos)
    TakeNthDirName = tmpStr
End Function


Public Function GetTempFolder() As String
'returns path to Windows temporary folder
' len(TEMP) < 256
    Dim tmpStr As String
    Dim reslen As Long
    tmpStr = VBA.Space(MAX_PATH)
    reslen = GetTempPath(Len(tmpStr), tmpStr)
    GetTempFolder = VBA.Left(tmpStr, reslen)
End Function

Public Function GetWinDir() As String
'returns Windows folder path
    Dim tmpStr As String
    Dim reslen As Long
    tmpStr = VBA.Space(MAX_PATH)
    reslen = GetWindowsDirectory(tmpStr, Len(tmpStr))
    GetWinDir = VBA.Left(tmpStr, reslen)
End Function

Public Function GetSysDir() As String
'returns Windows system folder path
    Dim tmpStr As String
    Dim reslen As Long
    tmpStr = VBA.Space(MAX_PATH)
    reslen = GetSystemDirectory(tmpStr, Len(tmpStr))
    GetSysDir = VBA.Left(tmpStr, reslen)
End Function

Public Function GetPathStrings() As Collection
'returns a collection of strings extracted from PATH system variable.
'suppose that len(PATH) =< 512
    Dim pathvar As String
    Dim tmpStr As String
    Dim reslen As Long
    Dim k As Long
    Dim a() As String
    Dim pathcol As New Collection
    tmpStr = VBA.Space(MAX_PATH)
    pathvar = "%PATH%"
    reslen = ExpandEnvironmentStrings(pathvar, tmpStr, Len(tmpStr))
    tmpStr = VBA.Left(tmpStr, reslen)
    tmpStr = MyTrim(tmpStr)
    a = Split(tmpStr, ";")
    For k = 0 To UBound(a)
        pathcol.Add VBA.Trim(a(k))
    Next
    Set GetPathStrings = pathcol
End Function

Private Function TakeExtension(FileName As String) As String
    Dim pos As Long
    Dim strlen As Long
    strlen = Len(FileName)
    pos = InStrRev(FileName, ".")
    If pos > 0 Then
        TakeExtension = VBA.Right(FileName, strlen - pos + 1)
    Else
        TakeExtension = ""
    End If
End Function

' timestamp parameter appends "timestamp" to the filename only if the same file already exists.
Public Function RenameFile(OldName As String, NewName As String, Optional TimeStamp As Boolean = False) As Boolean
    Dim res As Long
    Dim NewNameTime As String
    RenameFile = False
    NewNameTime = ""
    On Error GoTo ErrHand
    If FileExists(OldName) Then
        
        If NewName = "" Then
            NewName = "001.TMP"
        End If
        
        If FileExists(NewName) Then
            If Not TimeStamp Then
                Kill NewName
            Else
                ' use "seconds since midnight" function
                NewNameTime = TakePath(NewName) & "\" & TakeFilename(NewName) & CStr(Timer()) & TakeExtension(NewName)
                ' if file still exists, delete. No way.
                If FileExists(NewNameTime) Then
                    Kill NewNameTime
                End If
            End If
        End If
        ' that means we have TimeStamp ON
        If NewNameTime <> "" Then
            NewName = NewNameTime
        End If
        res = MoveFile(OldName, NewName)
        If res <> 0 Then
            RenameFile = True
        End If
    End If
ErrHand:
    If Err Then
        MsgBox Err.Description, vbExclamation, App.Title
    End If
End Function

'========= private functions ==============

Private Function MyTrim(strtotrim As String) As String
    Dim tmpStr As String
    Dim pos As Long
    pos = InStr(1, strtotrim, VBA.Chr(0))
    tmpStr = VBA.Left(strtotrim, pos)
    tmpStr = Replace(tmpStr, VBA.Chr(0), " ")
    MyTrim = VBA.Trim(tmpStr)
End Function


Public Function GetTempFN() As String
    GetTempFN = VBA.Right(CStr(Year(VBA.Date)), 2) & CStr(Month(VBA.Date)) & CStr(Day(VBA.Date)) & CStr(Hour(VBA.Time)) & CStr(Minute(VBA.Time)) & CStr(Second(VBA.Time))
End Function


Public Function StripNulls(strItem As String) As String
    Dim nPos As Integer
    
    nPos = InStr(strItem, VBA.Chr$(0))
    If nPos Then
        strItem = VBA.Left$(strItem, nPos - 1)
    End If
    StripNulls = strItem
End Function
 

Public Function IsFileOpen(FileName As String) As Boolean
    Dim filevar As Long
    filevar = FreeFile()
    On Error GoTo ErrHand
    Open FileName For Binary Access Write Lock Write As #filevar
    Close #filevar
    IsFileOpen = False
ErrHand:
    If Err Then
        IsFileOpen = True
        Err.Clear
    End If
End Function

Private Function GetRegValue(ByVal enumKey As enumHKEY_CLASS_LOCAL, ByVal Section As String, Optional ByVal SubKey As String = "", Optional Index As Variant) As String
    Dim nRet As Long
    Dim nType As Long
    Dim nBytes As Long
    Dim hKey As Long
    Dim tmpStr As String
    
    On Error GoTo ErrHand
    GetRegValue = ""
    nRet = RegOpenKeyEx(enumKey, Section, 0&, KEY_READ, hKey)
    If nRet = ERROR_SUCCESS Then
        nRet = RegQueryValueEx(hKey, SubKey, 0&, nType, ByVal tmpStr, nBytes)
        If nRet = ERROR_SUCCESS Then
            If nBytes > 0 Then
                tmpStr = VBA.Space(nBytes)
                nRet = RegQueryValueEx(hKey, SubKey, 0&, nType, ByVal tmpStr, Len(tmpStr))
                If nRet = ERROR_SUCCESS Then
                    GetRegValue = VBA.Left(tmpStr, nBytes - 1)
                End If
            End If
        End If
        Call RegCloseKey(hKey)
    End If
ErrHand:
    If Err Then
        Err.Clear
    End If
End Function

Public Function Substrings(ByVal Source As String, ByVal Substring As String) As Long
    Dim Count As Long
    Dim a() As String
    a = Split(Source, Substring)
    Count = UBound(a)
    Substrings = Count
End Function
