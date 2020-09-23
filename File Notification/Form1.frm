VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Monitor"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Log"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Monitor SubDirectory?"
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   1680
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Specify Flags"
      Height          =   1335
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   2895
      Begin VB.CheckBox CL_Write 
         Caption         =   "Last write"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox CL_Access 
         Caption         =   "Last Access"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox CF_size 
         Caption         =   "File size"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox CD_name 
         Caption         =   "Dir name"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox CF_name 
         Caption         =   "File name"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   3000
      Width           =   5895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start! Monitor"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Monitor"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'A couple of crashes and a lots of coding; and I present you the File monitor.
'It tries to give a VB.NET's "filesystem watcher" functionality to VB
'There's a lot of API in there, so begginers may find it difficult to understand at
'one shot. So I recommend you to look at the MSDN library for indepth
'description of APIs.
'Also, I've tried to comment the hard part as far as possible.
'However, Using this code may crash your VB IDE. SO I HAVE ONE THING TO SAY
'                      !!!!PROCEED AT YOUR OWN RISK!!!!!
Option Explicit
Dim M_Path As String, Flags As NotificationFlags
Dim No As Integer
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long



Private Sub Command1_Click()
    DisposeAll
End Sub

Private Sub Command2_Click()
'Set the flags, path and create a monitor and then start it
On Error GoTo ERrs:

Static MonitorCounter As Long

    If MonitorCounter = 0 Then MonitorCounter = 1

    Flags = Dir_Name * CD_name.Value Or _
            file_name * CF_name.Value Or _
            Last_Access * CL_Access.Value Or _
            Last_Write * CL_Write.Value Or _
            Size * CF_size.Value

    M_Path = Dir1.Path

    If AddMonitor(M_Path, Option1.Value, Flags, MonitorCounter) = 0 Then Exit Sub

    MonitorCounter = MonitorCounter + 1

    Dim St As String * 30

    GetShortPathName M_Path, M_Path, Len(M_Path)    'The textbox is quite small so use short paths
    
    M_Path = Left(M_Path, InStr(1, M_Path, vbNullChar) - 1)
    
    ADDtext "Monitor Added and started. Path is" & vbCrLf & M_Path

    modMainEx.Start     'Start the thread
    
    Exit Sub

ERrs:
If Err.Number = USER_ERROR + 108 Then
    MsgBox "Error: " & Err.Description & ", try another version of MSVBVM60.dll" _
            , vbExclamation Or vbOKOnly
End If
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    No = 0
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Public Sub ADDtext(txt As String)
    No = No + 1
    Text1.Text = Text1.Text & Str(No) & ".)  " & txt & vbCrLf
End Sub

Private Sub Form_Load()
'Check the Version.
Dim Version As String
'If app.exename returns the name of the project then the project is being debugged
'else if app.exename returns the name of the exe then project is running after being compiled
        
    Version = VersionInformation("MSVBVM60.dll")
    If Version <> "6.0.97.82" Then
        MsgBox "This program was made for MSVBVM60.dll of version 6.0.97.82. But you have a different one." _
             & vbCrLf & "However, you can try it. But if you get further errors, then you must discontinue it.", vbInformation Or vbOKOnly
    End If

    ADDtext "App started on ==> " & Now
    ADDtext "MSVBVM60.dll analysed. Version is ==> " & Version
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisposeAll
End Sub


Private Sub DisposeAll()
Dim i As Integer
For i = 1 To 10         'Disposes all Monitors
    DisposeMonitor i
Next
End Sub

