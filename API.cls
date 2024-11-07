VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function OpenProcess _
            Lib "kernel32" _
            (ByVal dwDesiredAccess As Long, _
            ByVal bInheritHandle As Long, _
            ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long

Dim token As String
Dim instance As String
Dim authHeader As String

Public Function init() As Boolean
    Dim token_, token_t, token_s, instance_, instance_t, instance_s

    Set token_ = CreateObject("Scripting.FileSystemObject")
    If Dir(App.Path & "\token.txt") <> "" Then ' naive check + we should store this in a JSON anyways
        Set token_t = token_.GetFile(App.Path & "\token.txt")
        Set token_s = token_t.OpenAsTextStream(1, -2)
        token = token_s.ReadLine

        Set instance_ = CreateObject("Scripting.FileSystemObject")
        Set instance_t = instance_.GetFile(App.Path & "\instance.txt")
        Set instance_s = instance_t.OpenAsTextStream(1, -2)
        instance = "https://" & instance_s.ReadLine & "/"

        authHeader = "Cookie: __Host-pleroma_key="

        init = True
    Else
        ' Config files don't exist yet, return failure
        init = False
    End If

End Function

Public Function request(endpoint As String, Optional data As String) As String
    ' TODO: do we want separate functions for different methods? TBD
    ' TODO: unfuck the CA bundle
    Dim payload As String

    If data <> "" Then
        Dim handle As Integer
        handle = FreeFile
        Open App.Path & "\data_in" For Output As #handle
        Print #handle, data
        Close #handle
        payload = "--data @" & App.Path & "\data_in "
        If Mid$(data, 1, 1) = "{" Then ' HACK, FIXME
            payload = payload & " -H ""content-type: application/json"" "
        End If
    End If

    Status = Shell(App.Path & "\curl.exe -k -H """ & authHeader & token & """ " & instance & endpoint & " " & payload & "-o " & App.Path & "\data", vbHide)
    hProcess = OpenProcess(&H400, False, Status) '&H400 = process query information
    
    Do
        GetExitCodeProcess hProcess, lExit
        DoEvents
    Loop While lExit = &H103 '&H103 = still active

    ' BAD BAD BAD TODO
    Set data_api_object = CreateObject("Scripting.FileSystemObject")
    Set data_api_file = data_api_object.GetFile(App.Path & "\data")
    Set data_api_readfile = data_api_file.OpenAsTextStream(1, -2)
    request = data_api_readfile.ReadLine
    data_api_readfile.Close

End Function
