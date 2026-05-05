Attribute VB_Name = "modMain"
Option Explicit
'Architecture:
'frmToastNotification - View layer. Gets its data driven design from GetToastStyle(tType)
'modToastService - Controller layer. Does all the logic how the application needs to work.
'modToastDesign - Design layer. Notification Enum and Type settings along with other settings how the form needs to look.
'modToastWindowEffects - Windows API working out the screen size and positioning of the toasts at the bottom right corner.


Private gNotificationCount As Long

Public Sub TestToast()
Attribute TestToast.VB_ProcData.VB_Invoke_Func = "T\n14"
  
  Dim title As String, message As String, duration As Long
  Dim sh As Worksheet
  
  Set sh = ThisWorkbook.Sheets("Sheet1")
    
  'Assign notification values
  title = sh.Cells(1, 2).Value
  message = sh.Cells(2, 2).Value
  duration = sh.Cells(3, 2).Value
  
  'Add a counter and set notification type before calling ShowToast
  gNotificationCount = gNotificationCount + 1
  Dim tType As enmToastType
  tType = gNotificationCount Mod 6 + 1
  
  ShowToast title, gNotificationCount & " - " & message, tType, duration

End Sub


Public Sub TestToastInfo()

  Dim title As String, message As String, duration As Long
  
  title = "This is a simple title"
  message = "This is a simple message but it does not have to be simpler. Just something that's not too long perhaps"
  duration = 5
  
  'Use Ctrl+Space shortcut to see available notification types
  ' Available options ttInfo, ttSuccess, ttWarning, ttError, ttQuestion, ttNeutral
    ShowToast title, message, ttInfo, 5
  
  ' The same can be called as using number 1 instead of ttInfo. Messages are enumarated from 1 to 6
  'ShowToast title, message, 1, 5
  
End Sub


Public Sub TestToastWarning()

  Dim title As String, message As String, duration As Long
  
  title = "This is a simple title"
  message = "This is a simple message but it does not have to be simpler. Just something that's not too long perhaps"
  duration = 5
  
  'Use Ctrl+Space shortcut to see available notification types
  ' Available options ttInfo, ttSuccess, ttWarning, ttError, ttQuestion, ttNeutral
    ShowToast title, message, ttWarning, 5
  
  ' The same can be called as using number 1 instead of ttInfo. Messages are enumarated from 1 to 6
  'ShowToast title, message, 3, 5
  
End Sub


'Just a macro to show how a basic notification compares to toast notifications
Public Sub DisplayBasicNotification()
  
  Dim title As String, message As String, duration As Long
  Dim sh As Worksheet
  
  Set sh = ThisWorkbook.Sheets("Sheet1")
    
  'Assign notification values
  title = sh.Cells(1, 2).Value
  message = sh.Cells(2, 2).Value
  
  MsgBox message, vbInformation, title
  
End Sub
