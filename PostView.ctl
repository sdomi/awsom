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
      ScaleHeight     =   435
      ScaleWidth      =   435
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
Public Property Let Content(ByVal value As String)
    messageObj.Caption = value
End Property

Public Property Get Content() As String
    Content = messageObj.Caption
End Property

Public Property Let Nickname(ByVal value As String)
    usernameObj.Caption = value
End Property

Public Property Get Nickname() As String
    Nickname = usernameObj.Caption
End Property
