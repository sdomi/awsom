VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If mshttp <> 1 Then
Private Declare Function OpenProcess _
            Lib "kernel32" _
            (ByVal dwDesiredAccess As Long, _
            ByVal bInheritHandle As Long, _
            ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long

Public Function fetch(ByVal url As String, ByVal Path As String)
    ' for fetching avatars and the like
    ' should probably get abstracted into the other API class. TBD

    Debug.Print App.Path & "\curl.exe -k " & url & " -o " & App.Path & Path
    Status = Shell(App.Path & "\curl.exe -k " & url & " -o " & App.Path & Path, vbHide)
    hProcess = OpenProcess(&H400, False, Status) '&H400 = process query information
    
    Do
        GetExitCodeProcess hProcess, lExit
        DoEvents
    Loop While lExit = &H103 '&H103 = still active
End Function

#Else
Public Function fetch(ByVal url As String, ByVal Path As String)
    Dim objReq
    Dim objStream
    
    On Error Resume Next ' fallible
    
    Set objReq = CreateObject("Microsoft.XMLHTTP")
    objReq.Open "GET", url, False
    objReq.Send
    
    If objReq.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
    
        objStream.Write objReq.ResponseBody
        objStream.Position = 0
    
        objStream.SaveToFile App.Path & Path, 2
        objStream.Close
    End If
End Function
#End If
