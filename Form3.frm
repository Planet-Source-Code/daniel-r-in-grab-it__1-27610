VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grab-IT"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form3.frx":0A02
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu Print 
         Caption         =   "Print"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Print_Click()
Form3.PrintForm
End Sub

Private Sub save_Click()
On Error GoTo errorcapture:
CommonDialog1.ShowSave
SavePicture Form3.Image, CommonDialog1.FileName
errorcapture:
End Sub
