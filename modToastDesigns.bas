Attribute VB_Name = "modToastDesigns"
Option Explicit
'A module used to apply base and toast notification specific styles
'Using Atlassian design system


'Enum determines how the notificaiton will look
Public Enum enmToastType
    ttInfo = 1
    ttSuccess = 2
    ttWarning = 3
    ttError = 4
    ttQuestion = 5
    ttNeutral = 6
End Enum


'Type variable captures all required properties in one variable
Public Type typToastStyle
  'Generic styles
  BgColor As Long
  BorderStyle As Long
  FontColor As Long
  FontSize As Long
  FontStyle As String
  
  'Type specific styles
  AccentLineColor As Long
  IconColor As Long
  IconChar As String
End Type


'Gets the default toast notification style as a base
Private Function GetBaseStyle() As typToastStyle

    Dim style As typToastStyle
    
    With style
      .BgColor = vbWhite
      .BorderStyle = 0 ' fmBorderStyleNone
      .FontColor = HexToVbaColorCode("#292a2e")
      .FontSize = 11
      .FontStyle = "Aptos Narrow"
      .AccentLineColor = HexToVbaColorCode("#292a2e")
      .IconChar = "i"
      .IconColor = HexToVbaColorCode("#292a2e")
    End With
    GetBaseStyle = style

End Function


'Returns the style type to apply to the form
Public Function GetToastStyle(ByVal ToastType As enmToastType) As typToastStyle
    
  Dim style As typToastStyle
  Dim color As Long
  
  'Validate toast style
  ValidateToastType ToastType
  
  'Apply base style
  style = GetBaseStyle()
  
  Select Case ToastType
    Case ttInfo
      color = HexToVbaColorCode("#1d7afc")
      style.AccentLineColor = color
      style.IconColor = color
      style.IconChar = ChrW(9432) 'Alternative to 105, 8505, 9432
    Case ttSuccess
      color = HexToVbaColorCode("#22a06b")
      style.AccentLineColor = color
      style.IconColor = color
      style.IconChar = ChrW(10003) '9989, 10003, 10004
    Case ttWarning
      color = HexToVbaColorCode("#ffc021") '#d97008
      style.AccentLineColor = color
      style.IconColor = color
      style.IconChar = ChrW(9888) '9888, 10071
    Case ttError
      color = HexToVbaColorCode("#e34935")
      style.AccentLineColor = color
      style.IconColor = color
      style.IconChar = ChrW(11199) '10005, 10006, 10060, 11198, 11199
    Case ttQuestion
      color = HexToVbaColorCode("#8270db")
      style.AccentLineColor = color
      style.IconColor = color
      style.IconChar = ChrW(63)  '63, 10067, 10068
    Case ttNeutral
      color = HexToVbaColorCode("#6b6e76")  '#505258
      style.AccentLineColor = color
      style.IconColor = color
      style.IconChar = ChrW(9432)
  End Select
  
  GetToastStyle = style

End Function


'A utility function converting
Private Function HexToVbaColorCode(ByVal HexCode As String) As Long
    Dim Red As Long, Green As Long, Blue As Long
    ' Convert hex to RGB
    Red = CLng("&H" & Mid(HexCode, 2, 2))
    Green = CLng("&H" & Mid(HexCode, 4, 2))
    Blue = CLng("&H" & Mid(HexCode, 6, 2))
    
    'Same as: HexToVbaColorCode = RGB(Red, Green, Blue)
    HexToVbaColorCode = Blue * 65536 + Green * 256 + Red
    
End Function


'Validates if a valid type has been passed in. Throws an error otherwise
Private Sub ValidateToastType(ByVal t As Long)

    Select Case t
        Case ttInfo, ttSuccess, ttWarning, ttError, ttQuestion, ttNeutral
            ' valid. do nothing
            
        Case Else
          On Error GoTo ErrorHandler
            Err.Raise vbObjectError + 1001, _
                "modToastDesigns.ValidateToastType", _
                "Invalid ToastType value passed in: " & t
    End Select

Exit Sub

ErrorHandler:
  MsgBox "An error occured in: " & Chr(13) _
          & Err.Source & Chr(13) & Chr(13) _
          & "Error message: " & Err.Description, vbExclamation
End Sub


