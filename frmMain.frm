VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GeoStar"
   ClientHeight    =   7455
   ClientLeft      =   7065
   ClientTop       =   2340
   ClientWidth     =   9015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   9015
   WindowState     =   2  'Maximized
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7425
      ScaleWidth      =   8985
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()

    P1.Height = frmMain.Height - 400
    P1.Width = frmMain.Width - 120

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Unload frmControl
    End

End Sub
