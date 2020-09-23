VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Form DropShadow"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cShadow As DropShadow

Private Sub Form_Load()
  Set cShadow = New DropShadow
  cShadow.CreateShadow hWnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set cShadow = Nothing
End Sub
