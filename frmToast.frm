VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmToast 
   Caption         =   "Form caption"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmToast.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmToast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub SetContent(ByVal title As String, ByVal message As String)
  Me.Caption = title
  Me.lblMessage.Caption = message
End Sub

