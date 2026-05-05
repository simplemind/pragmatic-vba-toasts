Attribute VB_Name = "modToastService"
Option Explicit
'Dependencies: modWindowEffects, modToastDesigns

'OnTime cannot access local variables so we create a global
Private gToasts As Collection
Private gToast As frmToastNotification


'Main entry point for the
Public Sub ShowToast(ByVal tTitle As String, ByVal tMessage As String, _
                      Optional ByVal tType As enmToastType, Optional ByVal tDuration As Long = 3)
  
  EnsureCollection
  
  Dim frm As frmToastNotification
  Set frm = New frmToastNotification

  'Populate the form with content
  frm.SetContent tTitle, tMessage
  
  'Style the form
  frm.ApplyStyle tType
  
  'Show the form
  frm.Show vbModeless

  DoEvents ' Ensure dimensions are cortypRect
  'Without DoEvents frm.Width and Height may still be 0 or incortypRect
  
  gToasts.Add frm 'Add to collection
    
  OrganizeAllToasts
    
  ScheduleClose frm, tDuration

End Sub


Private Sub StyleTheToast(ByRef frm As frmToastNotification, tType As enmToastType)
  
  'Get the style
  Dim tStyle As typToastStyle
  tStyle = GetToastStyle(tType)
  
  frm.lblAccentLine.BackColor = tStyle.AccentLineColor
  frm.lblIcon.Caption = tStyle.IconChar
  frm.lblIcon.ForeColor = tStyle.IconColor

End Sub


Private Sub CloseToast()

  If gToast Is Nothing Then Exit Sub
  
  Unload gToast
  Set gToast = Nothing

End Sub


'Positions the toast notification stack
' At the bottom right of the OS window
' Stacks the toasts oldest over newest
' Removes oldest notifications when new ones cannot fit into the available height
Private Sub OrganizeAllToasts()

If gToasts Is Nothing Then Exit Sub
    If gToasts.Count = 0 Then Exit Sub

    CleanInvalidToasts
    RemoveOverflowToasts
    LayoutToasts

End Sub


Private Sub CleanInvalidToasts()

    Dim i As Long
    
    For i = gToasts.Count To 1 Step -1
        If Not IsToastValid(gToasts(i)) Then
            gToasts.Remove i
        End If
    Next i

End Sub


Private Sub RemoveOverflowToasts()

    Const spacing As Long = 0
    
    Dim availableHeight As Double
    availableHeight = GetAvailableHeight()
    
    Dim totalHeight As Double
    totalHeight = GetTotalToastHeight(spacing)
    
    Do While totalHeight > availableHeight
        
        If gToasts.Count = 0 Then Exit Do
        
        Unload gToasts(1)
        gToasts.Remove 1
        
        totalHeight = GetTotalToastHeight(spacing)
        
    Loop

End Sub


Private Sub LayoutToasts()

    Const MARGIN As Long = 5
    Const spacing As Long = 0
    
    Dim rc As typRect
    rc = GetWorkArea()
    
    Dim screenRight As Double
    Dim screenBottom As Double
    
    screenRight = PixelsToPoints(rc.Right)
    screenBottom = PixelsToPoints(rc.Bottom)
    
    Dim currentBottom As Double
    currentBottom = screenBottom - MARGIN
    
    Dim i As Long
    Dim frm As frmToastNotification
    
    For i = gToasts.Count To 1 Step -1
        
        Set frm = gToasts(i)
        
        frm.Left = screenRight - frm.Width - MARGIN
        frm.Top = currentBottom - frm.Height
        
        currentBottom = frm.Top - spacing
        
    Next i

End Sub


'Safely initialise
Private Sub EnsureCollection()
    If gToasts Is Nothing Then
        Set gToasts = New Collection
    End If
End Sub


'Check if toast notification is available.
'May be closed by user unloading it from memory
Private Function IsToastValid(frm As Object) As Boolean
    On Error Resume Next
    IsToastValid = (frm.Visible = True)
End Function


'Get available height to display notifications
Private Function GetAvailableHeight() As Double

    Dim rc As typRect
    rc = GetWorkArea()
    
    GetAvailableHeight = PixelsToPoints(rc.Bottom - rc.Top)

End Function


Private Function GetTotalToastHeight(ByVal spacing As Long) As Double

    Dim total As Double
    Dim i As Long
    
    For i = 1 To gToasts.Count
        total = total + gToasts(i).Height
    Next i
    
    If gToasts.Count > 1 Then
        total = total + (gToasts.Count - 1) * spacing
    End If
    
    GetTotalToastHeight = total

End Function


'Schedules when to close a toast notification
Private Sub ScheduleClose(frm As frmToastNotification, Optional duration As Long = 3)

    'Store reference via Tag (simple trick)
    frm.Tag = Timer
    
    Application.OnTime Now + TimeSerial(0, 0, duration), "Toast_CloseHandler"

End Sub


Private Sub Toast_CloseHandler()

    If gToasts Is Nothing Then Exit Sub
    If gToasts.Count = 0 Then Exit Sub
    
    Dim frm As frmToastNotification
    
    'Close the oldest (top) if not closed manually
    If IsToastValid(gToasts(1)) Then
      Set frm = gToasts(1)
      RemoveToast frm
    End If

End Sub


Private Sub RemoveToast(frm As frmToastNotification)

    Dim i As Long
    
    ' Find and remove
    For i = gToasts.Count To 1 Step -1
      If gToasts(i) Is frm Then
        gToasts.Remove i
          Exit For
      End If
    Next i
    
    Unload frm
    
    ' Re-stack remaining
    OrganizeAllToasts

End Sub
