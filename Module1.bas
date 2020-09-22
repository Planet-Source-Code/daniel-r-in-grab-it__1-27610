Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Public Sub ScreenShot(DestinationDC As Long, Top As Long, Left As Long, Width As Long, Height As Long)
Dim ScreenDC As Long, retval As Long
ScreenDC = GetDC(GetDesktopWindow)
retval = BitBlt(DestinationDC, 0, 0, Width, Height, ScreenDC, Top, Left, vbSrcCopy)
End Sub



