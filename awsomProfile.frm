VERSION 5.00
Begin VB.Form awsomProfile 
   Caption         =   "<unset> - Profile"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Follow"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6120
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label bioObj 
      Caption         =   "<unset>"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6135
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
         Name            =   "MS Sans Serif"
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
Public Function selectUser(displayName As String, username As String, avatar As String)
    ' TODO: we probably want to pass an ID here and resolve everything manually :p
    displayNameObj.Caption = displayName
    usernameObj.Caption = username
    
    If Dir(avatar) <> "" Then
        Dim pic As Image
        userImg.Picture = LoadPicture(avatar)
        userImg.AutoRedraw = True
        userImg.PaintPicture userImg.Picture, _
        0, 0, userImg.ScaleWidth, userImg.ScaleHeight, _
        0, 0, userImg.Picture.Width / 26.46, userImg.Picture.Height / 26.46
    Else
        Debug.Print "could not load " & a
    End If
End Function
