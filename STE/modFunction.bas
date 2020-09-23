Attribute VB_Name = "modFunction"
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_END = 2
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1

Public Function GetLNKTarget(filename As String) As String
    Dim hFile As Long
    Dim tmpbyte     As Byte
    Dim tmpint      As Integer
    Dim dword       As Long
    Dim tmpPlace    As Long
    
    
    Dim i As Long
    hFile = CreateFile(filename, ByVal (GENERIC_READ Or GENERIC_WRITE), FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0)
    
    ReDim basebytes(75)
    If hFile <> INVALID_HANDLE_VALUE Then
        'move to the ShellItemId list offset
        'read two bytes. this is the number of
        'bytes we want to skip to get to the
        'start of the file location info.  Skip
        'ahead another 0x10, and we can get the
        'offset of the target path.
        SetFilePointer hFile, &H4C, 0, FILE_BEGIN
        ReadFile hFile, tmpint, 2, ret, ByVal 0&

        tmpPlace = CLng(tmpint + &H10)
        ret = SetFilePointer(hFile, tmpPlace, 0, FILE_CURRENT)
        
        'get that offset
        ReadFile hFile, dword, 4, ret, ByVal 0&

        'add that offset to the offset where the
        'file location info starts
        SetFilePointer hFile, &H4E + tmpint + dword, 0, FILE_BEGIN
        
        'set the counter to 0
        i = 0
        'loop until we get to a null byte (0x00)
        Do
            ReadFile hFile, tmpbyte, 1, ret, ByVal 0&
            'If it is a null string then get out of the loop
            If tmpbyte = &H0 Then Exit Do
            'if it isn't, appended the charachter to the end
            'of the path
            GetLNKTarget = GetLNKTarget & Chr(tmpbyte)
            'increase the counter
            i = i + 1
            DoEvents
        Loop Until tmpbyte = &H0
    End If
    CloseHandle hFile
End Function
