VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Fortschritt"
   ClientHeight    =   1320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4305
   OleObjectBlob   =   "ProgressBar.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Update(complete As Integer, total As Integer)
    Dim percent As Single
    percent = 100 * complete / total
    Text.Caption = complete & " von " & total & " erledigt"
    Bar.Width = percent * 2
    DoEvents
End Sub
