VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmToastNotification 
   Caption         =   "Form caption"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmToastNotification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmToastNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub SetContent(ByVal title As String, ByVal message As String)
  Me.Caption = title
  Me.lblMessage.Caption = message
End Sub


'Dependency: modToastDesigns
'The form is a single entry form for desing
Public Sub ApplyStyle(ByVal tType As enmToastType)

    Dim style As typToastStyle
    style = GetToastStyle(tType)
    
    ' Form
    Me.BackColor = style.BgColor
    Me.BorderStyle = style.BorderStyle
    
    ' Accent
    Me.lblAccentLine.BackColor = style.AccentLineColor
    
    ' Icon
    With Me.lblIcon
        .Caption = style.IconChar
        .ForeColor = style.IconColor
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' Text
    With Me.lblMessage
      .ForeColor = style.FontColor
      .FontName = style.FontStyle
      .FontSize = style.FontSize
      .BackStyle = fmBackStyleTransparent
    End With
End Sub

