VERSION 5.00
Begin VB.Form awsomPost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Toot"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Toot!"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
Private Sub Command1_Click()
    
    Set apiClient = New API
    If apiClient.init() Then
        apiClient.request "/api/v1/statuses"
    Else
        MsgBox "Error: could not initialize API", vbCritical
        Unload Me
    End If
    
    apiClient.request "/api/v1/statuses", "status=" & Text1.Text

    Unload Me
End Sub

