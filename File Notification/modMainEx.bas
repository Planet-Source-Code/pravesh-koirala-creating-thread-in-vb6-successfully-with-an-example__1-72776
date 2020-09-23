Attribute VB_Name = "modMainEx"
Private Declare Function ReadDirectoryChanges Lib "kernel32.dll" Alias "ReadDirectoryChangesW" (ByVal hDirectory As Long, ByVal lpBuffer As Long, ByVal nBufferLength As Long, ByVal bWatchSubTree As Long, ByVal dwNotifyFiler As Long, ByVal lpBytesReturned As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByRef lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForMultipleObjects Lib "kernel32.dll" (ByVal nCount As Long, ByRef lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long

Private Const INFINITE = &HFFFFFFFF

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type MonitorParams
    bWatchSubTree As Boolean
    Flags As NotificationFlags
    ThreadParams As Long
    DirectoryHandle As Long
    hEvents(1) As Long
End Type

Private Type OVERLAPPED
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type FILE_NOTIFY_INFORMATION
    NextEntryOffSet As Long
    Action As Long
    FileNameLength As Long
    FileName(255 - 1) As Byte
End Type

Private Const FILE_LIST_DIRECTORY As Long = &H1
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_FLAG_BACKUP_SEMANTICS As Long = &H2000000 'for opening a directory
Private Const FILE_FLAG_OVERLAPPED = &H40000000

Private Const FILE_ACTION_ADDED As Long = &H1
Private Const FILE_ACTION_MODIFIED As Long = &H3
Private Const FILE_ACTION_REMOVED As Long = &H2
Private Const FILE_ACTION_RENAMED_NEW_NAME As Long = &H5
Private Const FILE_ACTION_RENAMED_OLD_NAME As Long = &H4

Dim Res As Long, DirectoryHandle As Long
Dim SA As SECURITY_ATTRIBUTES
Dim MPS(9) As MonitorParams

Public Enum NotificationFlags
Attributes = &H4
Creation = &H40
Dir_Name = &H2
file_name = &H1
Last_Access = &H20
Last_Write = &H10
Size = &H8
End Enum

Public Function AddMonitor(ByVal Path As String, bWatchSubTree As Boolean, nFlags As NotificationFlags, Param As Long) As Integer
'Here we first create directory pointed by path param with FILE_FLAG_BACKUP_SEMANTICS
'and FILE_FLAG_OVERLAPPED for async operation. Async operation is used so that the
'thread could get stop signal that we generate.
'By Param argument the index of the thread is passed to it. Which it uses to retrieve
'its corresponding MP from the array of MP i.e MPS.

    Dim THandle As Long, Indx As Long
    
    If nFlags = 0 Then
        Form1.ADDtext "Can't Create Monitor. You must at least select one of the options"
        Exit Function
    End If
    
    DirectoryHandle = CreateFile(Path, FILE_LIST_DIRECTORY, FILE_SHARE_READ _
                        Or FILE_SHARE_WRITE, SA, OPEN_EXISTING, _
                        FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, 0)
    If DirectoryHandle = 0 Then
        Form1.ADDtext "Cannot create Directory Handle. You must check at least one checkbox"
        Exit Function
    End If
    
    Indx = Param - 1
    MPS(Indx).bWatchSubTree = IIf(bWatchSubTree, 1, 0)
    MPS(Indx).Flags = nFlags
    MPS(Indx).DirectoryHandle = DirectoryHandle
    MPS(Indx).hEvents(1) = CreateEvent(SA, 0, 0, "Completion Event" & Str(Indx))
    MPS(Indx).hEvents(0) = CreateEvent(SA, 0, 0, "Stop Monitor" & Str(Indx))
    MPS(Indx).ThreadParams = Param
    
    THandle = modThreading.CreateNewThread(AddressOf WaitForChange, True, Param)
    If THandle = 0 Then
        Form1.ADDtext "Thread Creation Failed"
        Exit Function
    Else
        Form1.ADDtext "Thread Created Successfully. It's handle is " & Str(THandle)
    End If
    
    AddMonitor = 1
    End Function

Public Sub Start()
    modThreading.StartThread
End Sub

Public Sub DisposeMonitor(ByVal TIndx As Integer)
 If MPS(TIndx - 1).hEvents(0) = 0 Then Exit Sub
 SetEvent MPS(TIndx - 1).hEvents(0)
 
End Sub


Public Function WaitForChange(ByVal Param As Long) As Long
    modThreading.InitThread
    WaitHere Param
End Function


Private Function WaitHere(ByVal Param As Long) As Long
Dim FName As String, bBytesReturned As Long, Pos As Long, WaitResult As Long
Static LogString As String
Dim MP As MonitorParams
Dim Buffer(1024 * 2 - 1) As Byte        'buffer size can be increased.
Dim Y As OVERLAPPED
Dim FNI As FILE_NOTIFY_INFORMATION

    MP = MPS(Param - 1)
    Debug.Print MP.hEvents(0), MP.hEvents(1)
    FName = Space$(255)
    Y.hEvent = MP.hEvents(1)        'Assign the completion event to OVERLAPPED structure

Start:
    Res = ReadDirectoryChanges(MP.DirectoryHandle, VarPtr(Buffer(0)), 1024 * 2, _
                        MP.bWatchSubTree, MP.Flags, VarPtr(bBytesReturned), Y, 0)

'Wait for IO operation to complete or the stop event to be signaled
    
    WaitResult = WaitForMultipleObjects(2, MP.hEvents(0), 0, INFINITE)
 
 Select Case WaitResult
  Case 0            'Wait was canceled
  
  Case 1            'A change has occured
      While True
        CopyMemory FNI, Buffer(Pos), Len(FNI)
        
        
        'Extract file name from the FNI.Filename array
        FName = FNI.FileName
        FName = Left(FName, FNI.FileNameLength \ 2)
        
        
        Select Case FNI.Action
            Case FILE_ACTION_ADDED
                LogString = FName & " was Added"
            Case FILE_ACTION_REMOVED
                LogString = FName & " was Deleted"
            Case FILE_ACTION_MODIFIED
                LogString = FName & " was modified"
            Case FILE_ACTION_RENAMED_NEW_NAME
                LogString = LogString & " into " & FName
            Case FILE_ACTION_RENAMED_OLD_NAME
                LogString = FName & " was renamed"
        End Select
                
        If FNI.Action <> FILE_ACTION_RENAMED_OLD_NAME Then
            Form1.ADDtext LogString & " On Thread " & Str(MP.ThreadParams)
            LogString = ""
        End If
        
        Pos = Pos + FNI.NextEntryOffSet
        
        If FNI.NextEntryOffSet = 0 Then     'it will be 0 only if the structure is the last in Buffer
            Pos = 0
            GoTo Start:
        End If
    Wend
  
  Case Else
    
  End Select
  
Form1.ADDtext "Thread ending " & "On Thread " & Str(MP.ThreadParams)

CloseHandle MP.hEvents(0)
CloseHandle MP.hEvents(1)

CloseHandle MP.DirectoryHandle
End Function


