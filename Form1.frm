VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   705
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3360
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":57E2
   ScaleHeight     =   47
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   2520
   End
   Begin VB.Menu end 
      Caption         =   "end"
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub exit_Click()
RemoveFromTray
Unload Form2
Unload Form3
Unload Me
End
End Sub

Private Sub Form_Load()
AddToTray Form3.Icon, "Grab-IT", Me
Form2.Width = Screen.Width
Form2.Height = Screen.Height
Form2.Top = 0
Form2.Left = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RespondToTray(X) = 2 Then
Me.PopupMenu Me.end
End If
End Sub

Private Sub Timer1_Timer()
If GetKeyState(17) = -128 Or GetKeyState(17) = -127 Then
    If GetKeyState(120) = -128 Or GetKeyState(120) = -127 Then
    Call ScreenShot(Form2.hDC, 0, 0, Me.Width, Me.Height)
    Form2.Show 1, Form1
    End If
End If
'Debug.Print GetKeyState(17), GetKeyState(120)
End Sub
