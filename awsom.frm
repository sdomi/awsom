VERSION 5.00
Begin VB.Form awsom 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "notSoAwsom"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "MS UI Gothic"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   3915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton refreshbt 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.Frame buttonframe 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton addbt 
         Caption         =   "New Toot"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5415
      Left            =   3720
      Max             =   1000
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "awsom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim old As Integer
Dim magic As Integer
Dim msgUserName() As Object, msgText() As Object, msgLike() As Object
Dim alreadydefined As Boolean ' this will break stuff in the future
Public postNo As Integer
Dim Status, hProcess As Long
Dim lExit, token_t, token_s, token_read, data_api_object, data_api_file, _
    data_api_readfile, userName, content, JB, read, token, instance, path, _
    instance_t, instance_s, instance_

Dim apiClient As API

Private Declare Function OpenProcess _
            Lib "kernel32" _
            (ByVal dwDesiredAccess As Long, _
            ByVal bInheritHandle As Long, _
            ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpexitcode As Long) As Long

Private Sub Form_Load()
    path = App.path
    magic = 0
    With VScroll1
        .Height = Me.ScaleHeight - 200
        .Min = 0
        .Max = 20000
        .SmallChange = Screen.TwipsPerPixelY * 10
        .LargeChange = .SmallChange
    End With
End Sub

Sub newItemSet(loops As Integer)
    While loops > 0
        If Not alreadydefined Then
            ReDim Preserve msgUserName(postNo)
            Set msgUserName(postNo) = Controls.Add("vb.label", "msgUserName" & postNo)
            ReDim Preserve msgText(postNo)
            Set msgText(postNo) = Controls.Add("vb.label", "msgText" & postNo)
            ReDim Preserve msgLike(postNo)
            Set msgLike(postNo) = Controls.Add("vb.commandbutton", "msgLike" & postNo)
        End If
        ' username
        msgUserName(postNo).Width = awsom.Width - VScroll1.Width - 75
        msgUserName(postNo).Top = 1500 * postNo
        msgUserName(postNo).Left = 0
        msgUserName(postNo).Height = 200
        msgUserName(postNo).Alignment = 2
        msgUserName(postNo).Caption = "Dynamic name"
        msgUserName(postNo).Font = "MS UI Gothic"
        msgUserName(postNo).Visible = True
        ' text
        msgText(postNo).Width = awsom.Width - VScroll1.Width - 200
        msgText(postNo).Top = 350 + 1500 * postNo
        msgText(postNo).Left = 100
        msgText(postNo).Height = 1000
        msgText(postNo).Caption = "Dynamic text"
        msgText(postNo).Font = "MS UI Gothic"
        msgText(postNo).Visible = True
        ' like button
        msgLike(postNo).Width = 500
        msgLike(postNo).Top = 900 + 1500 * postNo
        msgLike(postNo).Left = 100
        msgLike(postNo).Height = 300
        msgLike(postNo).Caption = "Like"
        msgLike(postNo).Font = "MS UI Gothic"
        msgLike(postNo).Visible = True

        ' todo: add avatar, images (if present), retoot, reply
        postNo = postNo + 1
        loops = loops - 1
    Wend
End Sub
     
Private Sub addbt_Click()
    awsomPost.Show
End Sub

    Private Sub refreshbt_Click()
    postNo = 1
    VScroll1.Value = 0
        Dim counter As Integer
        counter = 1
        If magic = 0 Then
            magic = 1
            
            Set apiClient = New API
            If apiClient.init() Then ' TODO: we shouldn't initialize it here, i'm just testing stuff
                Set JB = New JsonBag
                JB.JSON = apiClient.request("/api/v1/timelines/home")
                newItemSet (JB.Count)
                alreadydefined = True
                While JB.Count >= counter
                userName = JB.Item(counter).Item("account").Item("acct")
                content = JB.Item(counter).Item("content")
                Dim test_start As Long
                Dim test_end As Long
                Dim content_before, content_after
            
                Dim I
                I = 100
                While I > 0
                    test_start = InStr(content, "<")
                    test_end = InStr(content, ">")
                    If test_start = 0 Or test_end = 0 Then
                     I = 0
                    Else
                        content_before = Mid$(content, 1, test_start - 1)
                        content_after = Mid$(content, test_end + 1, 33333)
                    End If
                    content = content_before + content_after
                    I = I - 1
                Wend
                msgUserName(counter).Caption = userName
                msgText(counter).Caption = content
                counter = counter + 1
                Wend
        Else
            loginform.Show
            awsom.Hide
        End If

'test_image = InStr(content, "<a href=")  ' basically move this to the end
'test_image_end = InStr(content, """ ")   ' and add magic through parsing
'If test_image <> 0 And test_image_end <> 0 Then ' <a> tag, not through
'img_url = Mid$(content, test_image + 8, test_image_end)
'Picture1.Picture = LoadPicture(img_url)
'End If

        magic = 0
      End If
    End Sub

Private Sub VScroll1_Change() ' hack for scrolling, basically move everything (expect scroll and buttons) UP
    Dim eachctl As Control
    For Each eachctl In Me.Controls
        If Not (TypeOf eachctl Is VScrollBar) And Not eachctl.Name = "addbt" And Not eachctl.Name = "refreshbt" And Not eachctl.Name = "buttonframe" Then
            'MsgBox eachctl.Name
            eachctl.Top = eachctl.Top + old - VScroll1.Value
        End If
Next
old = VScroll1.Value

    
End Sub
