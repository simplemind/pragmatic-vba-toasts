Attribute VB_Name = "modMain"
Option Explicit
'Architecture:
'


Private gNotificationCount As Long

Public Sub TestToast()
Attribute TestToast.VB_ProcData.VB_Invoke_Func = "T\n14"
  
  Dim title As String, message As String
  title = ThisWorkbook.Sheets("Sheet1").Cells(1, 2).Value
  message = ThisWorkbook.Sheets("Sheet1").Cells(2, 2).Value
  
  gNotificationCount = gNotificationCount + 1
  Dim tType As enmToastType
  tType = gNotificationCount Mod 6 + 1
  ShowToast title, gNotificationCount & " - " & message, tType, 7

End Sub


'Temp sub
Public Sub DisplayBasicNotification()
  
  Dim title As String, message As String
  title = ThisWorkbook.Sheets("Sheet1").Cells(1, 2).Value
  message = ThisWorkbook.Sheets("Sheet1").Cells(2, 2).Value

    
  MsgBox message, vbInformation, title
  
End Sub


Public Sub DisplayIcons()

  Dim frm As frmIcons
  Set frm = frmIcons
  
  Dim arrIcons() As Variant
  arrIcons = Array("9888", "9989", "9940", "10003", "10004" _
                    , "10005", "10006", "10060", "10062", "10067" _
                    , "10069", "10071", "11197", "11198", "11199" _
                    , "33", "105", "9432", "8505" _
                    , "128512", "128500", "128501", "128502", "128503" _
                    , "128504", "128505")
  
  Dim i As Long, j As Long
  Dim lbl As MSForms.Label
  Dim StartPos As Long
  StartPos = 12
  Dim IconSize As Long
  IconSize = 35
  Dim ColSize As Long
  ColSize = 160
  Dim PosX As Long, PosY As Long
  
  
  frm.BackColor = vbWhite
  
  For i = 0 To UBound(arrIcons, 1)
    '10 per column
    PosX = 12 + ColSize * Application.WorksheetFunction.Quotient(i, 10)
    PosY = StartPos + IconSize * (i Mod 10)
    Set lbl = frm.Controls.Add("Forms.Label.1", "lblIcon_" & i, True)
    
      On Error Resume Next
      With lbl
        .Caption = ChrW(CLng(arrIcons(i))) & " - " & arrIcons(i)
      On Error GoTo 0
        .Top = PosY
        .Left = PosX
        .Width = ColSize
        .Height = 35
        .AutoSize = False
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 24
        .ForeColor = vbBlack
      End With
      
      'Debug.Print ChrW(arrIcons(i))
  Next i
  
  frm.Show

End Sub
