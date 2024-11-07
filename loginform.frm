VERSION 5.00
Begin VB.Form loginform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign in"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancelbt 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save and continue"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "API token"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
      Begin VB.TextBox token 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Instance (eg. niu.moe)"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      Begin VB.TextBox instance 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Some day, I plan to add support for logging in with your standard username, password and domain. TODO"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "It appears that this is your first time running awsom. Before continuing, you'll need to enter some details about your account in boxes below."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private file_token As Integer
Private file_instance As Integer
Private Path As String, token_path As String, instance_path As String

Private Sub cancelbt_Click()
    awsom.Show
    Unload loginform
End Sub

Private Sub Command1_Click()
    Path = App.Path
    token_path = Path & "\token.txt"
    instance_path = Path & "\instance.txt"
    Open token_path For Output As #1
    Open instance_path For Output As #2
    Print #1, token.Text
    Print #2, instance.Text
    Close #1
    Close #2
    
    awsom.Show
    Unload loginform
End Sub
