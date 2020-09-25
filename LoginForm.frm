VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login erforderlich"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public okAction As Boolean
Public username As String
Public password As String

Private Sub CancelButton_Click()
    username = ""
    password = ""
    Reset (False)
    okAction = False
    Hide
End Sub

Private Sub OkButton_Click()
    username = UserTextbox.Value
    password = PasswordTextbox.Value
    okAction = True
    Hide
End Sub

Public Sub Reset(complete As Boolean)
    okAction = False
    PasswordTextbox.Value = ""
    If (complete) Then
        UrlLabel = ""
        UserTextbox = ""
    End If
End Sub

