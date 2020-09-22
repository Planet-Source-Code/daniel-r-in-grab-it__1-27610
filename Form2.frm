VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   7
      Y2              =   102
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   0
      X2              =   99
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   130
      X2              =   130
      Y1              =   43
      Y2              =   127
   End
   Begin VB.Line Line4 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   28
      X2              =   115
      Y1              =   128
      Y2              =   128
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oldx As Long, oldy As Long
Dim upx As Long, upy As Long
Dim mousedown As Boolean

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = True

oldx = X
oldy = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If mousedown = True Then

Dim ty As Integer
Dim tx As Integer
ty = Y / Screen.TwipsPerPixelY
ty = ty - oldy / Screen.TwipsPerPixelY
tx = X / Screen.TwipsPerPixelX
tx = tx - oldx / Screen.TwipsPerPixelX
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Line1.X1 = oldx
Line1.X2 = oldx
Line1.Y1 = oldy
Line1.Y2 = Y
Line2.X1 = oldx
Line2.X2 = X
Line2.Y1 = oldy
Line2.Y2 = oldy
Line3.Y1 = oldy
Line3.Y2 = Y
Line3.X1 = X
Line3.X2 = X
Line4.X1 = oldx
Line4.X2 = X
Line4.Y1 = Y
Line4.Y2 = Y

End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mousedown = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
upx = X
upy = Y

If upx < oldx Then
    tmp = oldx
    oldx = upx
    upx = tmp
End If
If upy < oldy Then
    tmp = oldy
    oldy = upy
    upy = tmp
End If
Form3.Width = (upx * Screen.TwipsPerPixelX) - (oldx * Screen.TwipsPerPixelX) + 8 * Screen.TwipsPerPixelX
Form3.Height = (upy * Screen.TwipsPerPixelY) - (oldy * Screen.TwipsPerPixelY) + 54 * Screen.TwipsPerPixelY
Call BitBlt(Form3.hDC, 0, 0, upx - oldx, upy - oldy, Form2.hDC, oldx, oldy, vbSrcCopy)
Form2.Hide
Form3.Show

End Sub
