VERSION 5.00
Begin VB.Form awsomPost 
   Caption         =   "New Toot"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Toot!"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "awsomPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess _
                Lib "kernel32" _
                (ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwProcessId As Long) As Long
    Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long
' variable names in here should be changed
'basically half of code is here because i'm lazy and don't want to make a module
Private Sub Command1_Click()
        Dim token_t, token_s, token_read, token, instance, instance_t, instance_s, path
        path = App.path
        Set token = CreateObject("Scripting.FileSystemObject")
        Set token_t = token.GetFile(path & "\token.txt")
        Set token_s = token_t.OpenAsTextStream(1, -2)
        token_read = token_s.ReadLine
        Set instance_ = CreateObject("Scripting.FileSystemObject")
        Set instance_t = instance_.GetFile(path & "\instance.txt")
        Set instance_s = instance_t.OpenAsTextStream(1, -2)
        instance = instance_s.ReadLine
        Dim hProcess As Long
        Dim lExit
        Status = Shell(path & "\curl.exe --cacert " & path & "\cacert.pem --header ""Authorization: Bearer " & token_read & """ --data ""status=" & Text1.Text & """ https://" & instance & "/api/v1/statuses -o " & path & "\data", vbHide)
        hProcess = OpenProcess(&H400, False, Status)
        
        Do
            GetExitCodeProcess hProcess, lExit
            DoEvents
        Loop While lExit = &H103

        Unload Me
End Sub
