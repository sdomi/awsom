VERSION 5.00
Begin VB.Form awsomProfile 
   Caption         =   "awsom: <unset>"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar postListScroll 
      Height          =   2535
      Left            =   6120
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox bioObj 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1080
      Width           =   6135
   End
   Begin VB.CommandButton followBtn 
      Caption         =   "Follow"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox userImg 
      Height          =   855
      Left            =   120
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Line profileLine 
      X1              =   120
      X2              =   6120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label usernameObj 
      Caption         =   "<unset>"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label displayNameObj 
      Caption         =   "<unset>"
      BeginProperty Font 
         Size            =   17.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "awsomProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function selectUser(id As String)
    Dim handle As Integer, contents As String, user As JsonBag, _
        avatar As String
    handle = FreeFile
    Open App.Path & "\cache\" & id & ".json" For Input As #handle
    contents = StrConv(InputB(LOF(handle), handle), vbUnicode) 'https://stackoverflow.com/a/2875248 - maybe just abstract?
    Close #handle

    Set user = New JsonBag
    user.JSON = contents
    
    displayNameObj.Caption = user.Item("display_name")
    usernameObj.Caption = user.Item("fqn")
    Me.Caption = "awsom: " & user.Item("display_name") & " (" & user.Item("fqn") & ")"
    avatar = App.Path & "\cache\" & id & ".jpg"
    
    If Dir(avatar) <> "" Then
        Dim pic As Image
        userImg.Picture = LoadPicture(avatar)
        userImg.AutoRedraw = True
        userImg.PaintPicture userImg.Picture, _
        0, 0, userImg.ScaleWidth, userImg.ScaleHeight, _
        0, 0, userImg.Picture.Width / 26.46, userImg.Picture.Height / 26.46
    Else
        Debug.Print "could not load " & a
        userImg.Picture = Nothing
    End If

    ' hack hack hack
    ' todo: redo this when i'm less sick
    Dim I As Integer, test_start, test_end, content_before, content_after, bio
    bio = user.Item("note")
    I = 100
    Do While I > 0
        test_start = InStr(bio, "<br/>")
        If test_start = 0 Then
            Exit Do
        End If
        content_before = Mid$(bio, 1, test_start - 1)
        content_after = Mid$(bio, test_start + 5, 33333) '??? TODO
        bio = content_before & vbCrLf & content_after
        I = I - 1
    Loop

    I = 100
    Do While I > 0
        test_start = InStr(bio, "<")
        test_end = InStr(bio, ">")
        If test_start = 0 Or test_end = 0 Then
            Exit Do
        End If
        content_before = Mid$(bio, 1, test_start - 1)
        content_after = Mid$(bio, test_end + 1, 33333) '??? TODO
        bio = content_before & content_after
        I = I - 1
    Loop
    
    bioObj.Text = bio
End Function

Private Sub Form_Load()
    bioObj.BackColor = Me.BackColor
    bioObj.ForeColor = Me.ForeColor
End Sub

Private Sub Form_Resize()
    On Error Resume Next ' TODO: we can't guarantee that all of those values will be higher than 0

    postListScroll.Left = Me.Width - postListScroll.Width - 100
    postListScroll.Height = Me.Height - profileLine.Y1 - 420
    profileLine.X2 = Me.Width - 300
    bioObj.Width = Me.Width - 245
    displayNameObj.Width = Me.Width - userImg.Width - followBtn.Width - 165
    usernameObj.Width = Me.Width - userImg.Width - followBtn.Width - 165
    followBtn.Left = Me.Width - followBtn.Width - 165
End Sub
