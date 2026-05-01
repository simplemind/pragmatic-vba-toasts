Attribute VB_Name = "modMain"
Option Explicit

Private gNotificationCount As Long

Public Sub TestToast()
Attribute TestToast.VB_ProcData.VB_Invoke_Func = "T\n14"
  
  Dim title As String, message As String
  title = ThisWorkbook.Sheets("Sheet1").Cells(1, 2).Value
  message = ThisWorkbook.Sheets("Sheet1").Cells(2, 2).Value
  
  gNotificationCount = gNotificationCount + 1
  ShowToast title, gNotificationCount & " - " & message, 5

End Sub
