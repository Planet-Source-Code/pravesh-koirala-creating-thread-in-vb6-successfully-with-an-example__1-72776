Attribute VB_Name = "ModVersion"
Option Explicit

' API declarations.
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
    dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
    dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
    dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
    dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
    dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
    dwFileFlagsMask As Long        '  = &h3F for version "0.42"
    dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             '  e.g. VFT_DRIVER
    dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           '  e.g. 0
    dwFileDateLS As Long           '  e.g. 0
End Type

Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal length As Long)

' Return version information strings for a file.
Public Function VersionInformation(ByVal file_name As String) As String
Dim buffer() As Byte
Dim info_size As Long
Dim info_address As Long
Dim fixed_file_info As VS_FIXEDFILEINFO
Dim fixed_file_info_size As Long
Dim Dummy_Handle As Long
Dim Result As String
    ' Get the version information buffer size.
    info_size = GetFileVersionInfoSize(file_name, Dummy_Handle)
    If info_size = 0 Then
        MsgBox "No version information available"
        Exit Function
    End If

    ' Load the fixed file information into a buffer.
    ReDim buffer(1 To info_size)
    If GetFileVersionInfo(file_name, 0&, info_size, buffer(1)) = 0 Then
        MsgBox "Error getting version information"
        Exit Function
    End If
    If VerQueryValue(buffer(1), "\", info_address, fixed_file_info_size) = 0 Then
        MsgBox "Error getting fixed file version information"
        Exit Function
    End If

    ' Copy the information from the buffer into a
    ' usable structure.
    MoveMemory fixed_file_info, info_address, Len(fixed_file_info)

    ' Get the version information.
    With fixed_file_info
        Result = _
            Format$(.dwFileVersionMSh) & "." & _
            Format$(.dwFileVersionMSl) & "." & _
            Format$(.dwFileVersionLSh) & "." & _
            Format$(.dwFileVersionLSl)
End With
VersionInformation = Result

End Function

