VERSION 5.00
Begin VB.Form awsomPost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "awsom: New Toot"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox visibilityPickerObj 
      Height          =   315
      ItemData        =   "awsomPost.frx":0000
      Left            =   120
      List            =   "awsomPost.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton postSubmitBtn 
      Caption         =   "Toot!"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox postContentObj 
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
Public reply_to As String
Public reply_visibility As Integer

Private Sub Form_Load()
    ''TODO: reverse mapping from numbers to strings ,_, i don't want to use a loop for this
    visibilityPickerObj.ListIndex = 0
End Sub

Private Sub postSubmitBtn_Click()
    Dim payload As JsonBag
    Set payload = New JsonBag
    
    Dim visibility_api(0 To 4) As String
    visibility_api(0) = "public"
    visibility_api(1) = "local"
    visibility_api(2) = "unlisted"
    visibility_api(3) = "private"
    visibility_api(4) = "direct"
    Dim selected_visibility As String
    If visibilityPickerObj.ListIndex < 0 Then
        selected_visibility = "public"
    Else
        selected_visibility = visibility_api(visibilityPickerObj.ListIndex)
    End If
    
    With payload
        .Clear
        .Item("status") = postContentObj.Text
        .Item("visibility") = selected_visibility
        .Item("sensitive") = False
        .Item("spoiler_text") = ""
        
        If reply_to <> "" Then
            .Item("in_reply_to_id") = reply_to
            reply_to = "" ' for now, this property lives on until awsom is closed. hence, clear it
        End If
    End With
    
    Set apiClient = New API
    If apiClient.init() Then
        apiClient.request "/api/v1/statuses", payload.JSON
    Else
        MsgBox "Error: could not initialize API", vbCritical
        Unload Me
    End If
    
    Unload Me
End Sub

