Attribute VB_Name = "modTime"
Option Explicit

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal NoSecurity As Long, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function GetFileTime Lib "kernel32" ( _
    ByVal hFile As Long, lpCreationTime As FILETIME, _
    lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
    
Private Declare Function SetFileTime Lib "kernel32" ( _
    ByVal hFile As Long, lpCreationTime As FILETIME, _
    lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Private Declare Function SetFileCreatedTime Lib "kernel32" Alias "SetFileTime" ( _
    ByVal hFile As Long, lpCreationTime As FILETIME, _
    ByVal NullLastAccessTime As Long, ByVal NullLastWriteTime As Long) As Long

Private Declare Function SetFileAccessTime Lib "kernel32" Alias "SetFileTime" ( _
    ByVal hFile As Long, ByVal NullCreationTime As Long, _
    lpLastAccessTime As FILETIME, ByVal NullWriteTime As Long) As Long

Private Declare Function SetFileModifiedTime Lib "kernel32" Alias "SetFileTime" ( _
    ByVal hFile As Long, ByVal NullCreationTime As Long, _
    ByVal NullLastAccessTime As Long, lpLastWriteTime As FILETIME) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
    lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long

Private Declare Function LocalFileTimeToFileTime Lib "kernel32" ( _
    lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" ( _
    lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Function SystemTimeToDate(System_Time As SYSTEMTIME) As Date
    
    '
    ' Convert a SYSTEMTIME into a Date.
    '
    
    With System_Time
        SystemTimeToDate = CDate( _
            Format$(.wMonth) & "/" & _
            Format$(.wDay) & "/" & _
            Format$(.wYear) & " " & _
            Format$(.wHour) & ":" & _
            Format$(.wMinute, "00") & ":" & _
            Format$(.wSecond, "00"))
    End With
    
End Function
Public Function DateToSystemTime(ByVal dteDate As Date) As SYSTEMTIME

    '
    ' Convert a Date into a SYSTEMTIME.
    '
    
    With DateToSystemTime
        .wYear = Year(dteDate)
        .wMonth = Month(dteDate)
        .wDay = Day(dteDate)
        .wHour = Hour(dteDate)
        .wMinute = Minute(dteDate)
        .wSecond = Second(dteDate)
    End With
    
End Function
Public Function FileTimeToDate(File_Time As FILETIME) As Date
Dim System_Time As SYSTEMTIME

    '
    ' Convert the FILETIME structure into a Date.
    '
    
    ' Convert the FILETIME into a SYSTEMTIME.
    Call FileTimeToSystemTime(File_Time, System_Time)

    ' Convert the SYSTEMTIME into a Date.
    FileTimeToDate = SystemTimeToDate(System_Time)
    
End Function
Public Function DateToFileTime(ByVal dteDate As Date) As FILETIME
Dim File_Time   As FILETIME
Dim System_Time As SYSTEMTIME

    '
    ' Convert a Date into a FILETIME structure.
    '
    
    ' Convert the Date into a SYSTEMTIME.
    System_Time = DateToSystemTime(dteDate)

    ' Convert the SYSTEMTIME into a FILETIME.
    Call SystemTimeToFileTime(System_Time, File_Time)
    
    ' Return the FILETIME structure.
    DateToFileTime = File_Time
    
End Function
Public Function GetFileTimes( _
                ByVal strFile As String, _
                ByRef dteCreate As Date, _
                ByRef dteAccess As Date, _
                ByRef dteModify As Date, _
                ByVal blnUseLocalTime As Boolean) As Boolean

Dim lngHandle         As Long
Dim Create_Filetime   As FILETIME
Dim Access_Filetime   As FILETIME
Dim Modified_Filetime As FILETIME
Dim TempFileTime As FILETIME

    GetFileTimes = False

    '
    ' Open the file.
    '
    lngHandle = CreateFile(strFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If lngHandle = 0 Then Exit Function

    '
    ' Get the times.
    '
    If GetFileTime(lngHandle, Create_Filetime, Access_Filetime, Modified_Filetime) = 0 Then
        Call CloseHandle(lngHandle)
        Exit Function
    End If

    '
    ' Close the file.
    '
    If CloseHandle(lngHandle) = 0 Then Exit Function

    '
    ' Convert to the local file system time.
    '
    If blnUseLocalTime Then
        Call FileTimeToLocalFileTime(Create_Filetime, TempFileTime)
        Create_Filetime = TempFileTime

        Call FileTimeToLocalFileTime(Access_Filetime, TempFileTime)
        Access_Filetime = TempFileTime

        Call FileTimeToLocalFileTime(Modified_Filetime, TempFileTime)
        Modified_Filetime = TempFileTime
    End If

    '
    ' Convert into dates.
    '
    dteCreate = FileTimeToDate(Create_Filetime)
    dteAccess = FileTimeToDate(Access_Filetime)
    dteModify = FileTimeToDate(Modified_Filetime)

    GetFileTimes = True
    
End Function
Public Function SetFileTimes( _
                ByVal strFile As String, _
                ByVal dteCreate As Date, _
                ByVal dteAccess As Date, _
                ByVal dteModify As Date, _
                ByVal blnUseLocalTime As Boolean) As Boolean

Dim lngHandle         As Long
Dim lngAttributes     As Long
Dim Create_Filetime   As FILETIME
Dim Access_Filetime   As FILETIME
Dim Modified_Filetime As FILETIME
Dim TempFileTime      As FILETIME

    SetFileTimes = False

    '
    ' Open the file after removing the read only flag.
    '
    lngAttributes = GetAttr(strFile)
    Call SetAttr(strFile, vbNormal)

    lngHandle = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If lngHandle = 0 Then
        Call SetAttr(strFile, lngAttributes)
        Exit Function
    End If

    '
    ' Convert the dates into FILETIMEs.
    '
    Create_Filetime = DateToFileTime(dteCreate)
    Access_Filetime = DateToFileTime(dteAccess)
    Modified_Filetime = DateToFileTime(dteModify)

    '
    ' Convert the file times into system file times.
    '
    If blnUseLocalTime Then
        Call LocalFileTimeToFileTime(Create_Filetime, TempFileTime)
        Create_Filetime = TempFileTime

        Call LocalFileTimeToFileTime(Access_Filetime, TempFileTime)
        Access_Filetime = TempFileTime

        Call LocalFileTimeToFileTime(Modified_Filetime, TempFileTime)
        Modified_Filetime = TempFileTime
    End If

    '
    ' Set the times.
    '
    If SetFileTime(lngHandle, Create_Filetime, Access_Filetime, Modified_Filetime) = 0 Then
        Call SetAttr(strFile, lngAttributes)
        Call CloseHandle(lngHandle)
        Exit Function
    End If

    '
    ' Close the file. Reset its attributes.
    '
    If CloseHandle(lngHandle) = 0 Then Exit Function
    Call SetAttr(strFile, lngAttributes)
    SetFileTimes = True
    
End Function
Public Function SetFileModifiedDate( _
                ByVal strFile As String, _
                ByVal dteNewDate As Date, _
                ByVal blnUseLocalTime As Boolean) As Boolean

Dim lngHandle     As Long
Dim lngAttributes As Long
Dim TempFileTime  As FILETIME
Dim FileTimeToUse As FILETIME

    SetFileModifiedDate = False

    '
    ' Convert the date to use to a FILETIME.
    '
    FileTimeToUse = DateToFileTime(dteNewDate)

    '
    ' Convert the file time into a system file time.
    '
    If blnUseLocalTime Then
        Call LocalFileTimeToFileTime(FileTimeToUse, TempFileTime)
        FileTimeToUse = TempFileTime
    End If

    '
    ' Open the file after removing the read only flag.
    '
    lngAttributes = GetAttr(strFile)
    Call SetAttr(strFile, vbNormal)
    
    lngHandle = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If lngHandle = 0 Then Exit Function

    '
    ' Set the time.
    '
    If SetFileModifiedTime(lngHandle, ByVal 0&, ByVal 0&, FileTimeToUse) = 0 Then
        Call CloseHandle(lngHandle)
        Exit Function
    End If

    '
    ' Close the file. Reset its attributes.
    '
    If CloseHandle(lngHandle) = 0 Then Exit Function
    Call SetAttr(strFile, lngAttributes)
    SetFileModifiedDate = True
    
End Function
Public Function SetFileAccessedDate( _
                ByVal strFile As String, _
                ByVal dteNewDate As Date, _
                ByVal blnUseLocalTime As Boolean) As Boolean

Dim lngHandle     As Long
Dim lngAttributes As Long
Dim TempFileTime  As FILETIME
Dim FileTimeToUse As FILETIME

    SetFileAccessedDate = False

    '
    ' Convert the date to use to a FILETIME.
    '
    FileTimeToUse = DateToFileTime(dteNewDate)

    '
    ' Convert the file time into a system file time.
    '
    If blnUseLocalTime Then
        Call LocalFileTimeToFileTime(FileTimeToUse, TempFileTime)
        FileTimeToUse = TempFileTime
    End If

    '
    ' Open the file after removing the read only flag.
    '
    lngAttributes = GetAttr(strFile)
    Call SetAttr(strFile, vbNormal)
    
    lngHandle = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If lngHandle = 0 Then Exit Function

    '
    ' Set the time.
    '
    If SetFileAccessTime(lngHandle, ByVal 0&, FileTimeToUse, ByVal 0&) = 0 Then
        CloseHandle lngHandle
        Exit Function
    End If

    '
    '
    ' Close the file. Reset its attributes.
    '
    If CloseHandle(lngHandle) = 0 Then Exit Function
    Call SetAttr(strFile, lngAttributes)
    SetFileAccessedDate = True

End Function
Public Function SetFileCreatedDate( _
                ByVal strFile As String, _
                ByVal dteNewDate As Date, _
                ByVal blnUseLocalTime As Boolean) As Boolean

Dim lngHandle     As Long
Dim lngAttributes As Long
Dim TempFileTime  As FILETIME
Dim FileTimeToUse As FILETIME

    SetFileCreatedDate = True

    '
    ' Convert the date to use to a FILETIME.
    '
    FileTimeToUse = DateToFileTime(dteNewDate)

    '
    ' Convert the file time into a system file time.
    '
    If blnUseLocalTime Then
        Call LocalFileTimeToFileTime(FileTimeToUse, TempFileTime)
        FileTimeToUse = TempFileTime
    End If

    '
    ' Open the file after removing the read only flag.
    '
    lngAttributes = GetAttr(strFile)
    Call SetAttr(strFile, vbNormal)
    
    lngHandle = CreateFile(strFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If lngHandle = 0 Then Exit Function

    '
    ' Set the time.
    '
    If SetFileCreatedTime(lngHandle, FileTimeToUse, ByVal 0&, ByVal 0&) = 0 Then
        Call CloseHandle(lngHandle)
        Exit Function
    End If

    '
    ' Close the file. Reset its attributes.
    '
    If CloseHandle(lngHandle) = 0 Then Exit Function
    Call SetAttr(strFile, lngAttributes)
    SetFileCreatedDate = True
    
End Function



