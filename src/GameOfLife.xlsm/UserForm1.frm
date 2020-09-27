VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Private Sub CommandButton1_Click()
    If Selection.Interior.Color = RGB(0, 0, 0) Then
        Selection.Interior.ColorIndex = 0
    Else
        Selection.Interior.Color = RGB(0, 0, 0)
    End If
End Sub

Private Sub CommandButton2_Click()
    GameOfLife Selection
End Sub
