Attribute VB_Name = "modThreading"
'Thread Functions
Private Declare Function CreateThread Lib "kernel32.dll" (ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadId As Long) As Long
Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Private Declare Sub ExitThread Lib "kernel32.dll" (ByVal dwExitCode As Long)

'Tls Functions
Private Declare Function TlsAlloc Lib "kernel32.dll" () As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
Private Declare Function TlsSetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long, ByRef lpTlsValue As Any) As Long

'Memory Functions
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long

Private Const CREATE_SUSPENDED As Long = &H4
Public Const USER_ERROR = 1051          'Application Defined error starts from 1051

'Structures
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private THandle As Long

Private MemAddress As Long, TlsIndex As Long, TlsAddress As Long

Public Function CreateNewThread(ByVal ThreadProcedure As Long, ByVal CreateSuspended As _
                                Boolean, Optional ByVal Param As Long = 0) As Long
                                
'And we allocate a memory for a new thread which will later
'be set in the Tls of the thread. The Tls index that VB creates was stored in the
'memory location &H6610EE7C in my version of MSVBVM60. So it may differ in another
'versions. Then we create thread which will be suspended and only run after a call to Start.

Dim Y As SECURITY_ATTRIBUTES, CreationFlags As Long


'Now Get The address where Tlsindex is stored
GetTlsIndex
If TlsAddress = 0 Then GoTo EX:
CopyMemory TlsIndex, ByVal TlsAddress, Len(TlsIndex)     'Retrieve TlsIndx from TlsAddress
MemAddress = TlsGetValue(TlsIndex)

EX:
If MemAddress = 0 Then
    Err.Clear
    Err.Raise USER_ERROR + 108, "ModThreading", "Cannot retrieve TlsIndex"
    Exit Function
End If
CreationFlags = IIf(CreateSuspended, CREATE_SUSPENDED, 0)
THandle = CreateThread(Y, 0, ThreadProcedure, ByVal Param, CreationFlags, Tid)        'Create thread.
CreateNewThread = THandle
End Function

Public Sub StartThread()
    ResumeThread THandle                                                 'Resume thread
End Sub

Public Sub InitThread()
'We will set the address of memory we created in the Tlsindex retrieved from the
'MSVBVM60.DLL. VB will use this address to store DLL error information and etcs.
    TlsSetValue TlsIndex, ByVal MemAddress
End Sub

Private Sub GetTlsIndex()
'This block of the code is quite messy. One need to observe carefully.
'Here we first load MSVBVM60.dll in memory and find the address of __vbaSetSystemError
'procedure. Then we copy the whole block of the procedure so that we can find the
'Tls Index's address of our thread.
    Dim Loc(40) As Byte, St As String
    Dim Add As Long, Hnd As Long, i As Integer, j As Integer

    Hnd = LoadLibrary("MSVBVM60.dll")                'Load MSVBVM60.dll
    Add = GetProcAddress(Hnd, "__vbaSetSystemError") 'Find Address of the procedure

    CopyMemory Loc(0), ByVal (Add), 40
    
    While Loc(i) <> &HC3                '&HC3 is equivalent to RETN
    
        If Loc(i) = &HFF And Loc(i + 1) = &H35 Then
        
            For j = i + 2 To i + 5
                St = Hex(Loc(j)) & St
            Next
            
           TlsAddress = Val("&H" & St)
            
        End If
        
        i = i + 1
        
    Wend
    
    FreeLibrary Hnd
    
End Sub

Public Sub TerminateThread(ByVal dwExitCode As Long)
'Using this procedure while debugging can crash your IDE
        ExitThread dwExitCode
End Sub

