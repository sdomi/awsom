VERSION 5.00
Begin VB.UserControl PostView 
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2250
   ScaleWidth      =   4800
   Begin VB.PictureBox authorImg 
      Height          =   495
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   4
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton starBtn 
      Caption         =   "x"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton boostBtn 
      Caption         =   "="
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton replyBtn 
      Caption         =   "<-"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label usernameObj 
      Caption         =   "<name>"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label messageObj 
      Caption         =   "<content>"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "PostView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public imgPath As String ' todo hack
Public userId As String

Public Property Let content(ByVal value As String)
    messageObj.Caption = value
End Property

Public Property Get content() As String
    content = messageObj.Caption
End Property

Public Property Let Nickname(ByVal value As String)
    usernameObj.Caption = value
End Property

Public Property Get Nickname() As String
    Nickname = usernameObj.Caption
End Property

Public Property Let avatar(ByVal value As String)
    imgPath = App.Path & value
    If Dir(imgPath) <> "" Then
        Dim pic As Image
        On Error Resume Next ' we cannot guarantee that all images won't error
        authorImg.Picture = LoadPicture(imgPath)
        authorImg.AutoRedraw = True
        authorImg.PaintPicture authorImg.Picture, _
        0, 0, authorImg.ScaleWidth, authorImg.ScaleHeight, _
        0, 0, authorImg.Picture.Width / 26.46, authorImg.Picture.Height / 26.46
    Else
        Debug.Print "could not load " & a
    End If
End Property

Private Sub authorImg_Click()
    awsomProfile.selectUser userId
    awsomProfile.Show ' TODO: instantiate multiple at once
End Sub

