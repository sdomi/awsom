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
Private old As Integer
Private magic As Integer
Private postList() As Object
Private alreadydefined As Boolean ' this will break stuff in the future
Public postNo As Integer
Private Status, hProcess As Long

Private httpClient As HTTP
Private apiClient As API

Private Sub Form_Load()
    magic = 0
    With VScroll1
        .Height = Me.ScaleHeight - 200
        .Min = 0
        .Max = 20000
        .SmallChange = Screen.TwipsPerPixelY * 10
        .LargeChange = .SmallChange
    End With
    
    Set httpClient = New HTTP
End Sub

Private Sub Form_Resize()
    Dim eachctl As Control
    For Each eachctl In Me.Controls
        If Not (TypeOf eachctl Is VScrollBar) And Not eachctl.Name = "addbt" And Not eachctl.Name = "refreshbt" And Not eachctl.Name = "buttonframe" Then
            eachctl.Width = awsom.Width - VScroll1.Width - 75
        End If
    Next
    VScroll1.Left = awsom.Width - VScroll1.Width - 100
    VScroll1.Height = awsom.Height - 370 ' twigs are wacky
End Sub

Private Sub VScroll1_Change() ' hack for scrolling, basically move everything (except scroll and buttons) UP
    Dim eachctl As Control
    For Each eachctl In Me.Controls
        If Not (TypeOf eachctl Is VScrollBar) And Not eachctl.Name = "addbt" And Not eachctl.Name = "refreshbt" And Not eachctl.Name = "buttonframe" Then
            'MsgBox eachctl.Name
            eachctl.Top = eachctl.Top + old - VScroll1.value
        End If
    Next
    old = VScroll1.value
End Sub

Private Sub addbt_Click()
    awsomPost.Show
End Sub

Private Sub refreshbt_Click()
    postNo = 1
    VScroll1.value = 0
    Dim counter As Integer
    counter = 1
    If magic = 0 Then
        magic = 1

        Set apiClient = New API
        If apiClient.init() Then ' TODO: we shouldn't initialize it here, i'm just testing stuff
            Dim JB
            Set JB = New JsonBag
            JB.JSON = apiClient.request("/api/v1/timelines/home")
            
            ' TODO: hack hack hack hack
            ' the below results in an overflow because PostView has a height of 2350 now
            ' can we use pixels instead of twigs? if not, lazy loading it is...
            
            Dim amount
            If JB.Count > 10 Then
                amount = 10
            Else
                amount = JB.Count
            End If
            
            newItemSet (amount)
            alreadydefined = True
            While amount >= counter
                Dim username, content, avatarPath, avatar
                username = JB.Item(counter).Item("account").Item("acct")
                
                avatar = JB.Item(counter).Item("account").Item("avatar")
                Debug.Print avatar
                If InStr(avatar, ".jpg") <> 0 Or InStr(avatar, ".jpeg") <> 0 Then
                    avatarPath = "\cache\" & JB.Item(counter).Item("account").Item("id") & ".jpg"
                    If Dir(App.Path & avatarPath) = "" Then
                        httpClient.fetch avatar, avatarPath
                        postList(counter).avatar = avatarPath
                    End If
                End If
                postList(counter).userId = JB.Item(counter).Item("account").Item("id")
                
                ' TODO: improve this somewhat
                Dim handle
                handle = FreeFile
                Open App.Path & "\cache\" & JB.Item(counter).Item("account").Item("id") & ".json" For Output As #handle
                Print #handle, JB.Item(counter).Item("account").JSON
                Close #handle

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
                        content_after = Mid$(content, test_end + 1, 33333) '??? TODO
                    End If
                    content = content_before + content_after
                    I = I - 1
                Wend
                postList(counter).Nickname = username
                postList(counter).content = content
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

Sub newItemSet(loops As Integer)
    While loops > 0
        If Not alreadydefined Then
            ReDim Preserve postList(postNo)
            Set postList(postNo) = Controls.Add("notSoAwsom.PostView", "postView" & postNo)
        End If
        
        postList(postNo).Width = awsom.Width - VScroll1.Width - 75
        postList(postNo).Top = 350 + 2300 * postNo ' TODO
        postList(postNo).Left = 0
        postList(postNo).Visible = True
        
        postNo = postNo + 1
        loops = loops - 1
    Wend
End Sub
