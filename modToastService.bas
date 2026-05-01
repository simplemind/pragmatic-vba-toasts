Attribute VB_Name = "modToastService"
Option Explicit
'Uses modWindowEffects dependency

'OnTime cannot access local variables so we create a global
Private gToasts As Collection
Private gToast As frmToast


'Main entry point for the
Public Sub ShowToast(ByVal title As String, ByVal message As String, Optional ByVal duration As Long = 3)
  
  EnsureCollection
  
  Dim frm As frmToast
  Set frm = New frmToast

  frm.SetContent title, message
  frm.Show vbModeless

  DoEvents ' Ensure dimensions are cortypRect
  'Without DoEvents frm.Width and Height may still be 0 or incortypRect
  
  '#ToDo: Add notification types: information/white, success/green, warning/yellow, error/red, blue, purple
  gToasts.Add frm 'Add to collection
    
  OrganizeAllToasts
    
  ScheduleClose frm, duration

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
    Dim frm As frmToast
    
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
Private Sub ScheduleClose(frm As frmToast, Optional duration As Long = 3)

    'Store reference via Tag (simple trick)
    frm.Tag = Timer
    
    Application.OnTime Now + TimeSerial(0, 0, duration), "Toast_CloseHandler"

End Sub


Private Sub Toast_CloseHandler()

    If gToasts Is Nothing Then Exit Sub
    If gToasts.Count = 0 Then Exit Sub
    
    Dim frm As frmToast
    
    'Close the oldest (top) if not closed manually
    If IsToastValid(gToasts(1)) Then
      Set frm = gToasts(1)
      RemoveToast frm
    End If

End Sub


Private Sub RemoveToast(frm As frmToast)

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
