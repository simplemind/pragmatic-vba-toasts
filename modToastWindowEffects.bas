Attribute VB_Name = "modToastWindowEffects"
Option Explicit
'Provides notification positioning within OS window
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1


'Accurate DPI-aware conversion
#If VBA7 Then
  Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
  Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#End If
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90


'Defining a suitable window area for displaying the notifications#
#If VBA7 Then
    Private Declare PtrSafe Function SystemParametersInfo Lib "user32" _
        Alias "SystemParametersInfoA" ( _
        ByVal uiAction As Long, _
        ByVal uiParam As Long, _
        ByRef pvParam As Any, _
        ByVal fWinIni As Long) As Long
#End If
Private Const SPI_GETWORKAREA As Long = 48


'Type for suitable display area typRectangle
Public Type typRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'Helper functions
Public Function GetScreenWidth() As Long
    GetScreenWidth = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function GetScreenHeight() As Long
    GetScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Function


'A utility to convert from pixels to VBA points
Public Function PixelsToPoints(px As Long) As Double
    Dim hdc As LongPtr
    hdc = GetDC(0)
    
    Dim dpiX As Long
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    
    ReleaseDC 0, hdc
    
    PixelsToPoints = px * 72 / dpiX
End Function


'Gets a working display area for notifications
Public Function GetWorkArea() As typRect

    Dim tintypRect As typRect
    SystemParametersInfo SPI_GETWORKAREA, 0, tintypRect, 0
    
    GetWorkArea = tintypRect

End Function
