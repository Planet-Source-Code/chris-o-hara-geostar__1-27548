VERSION 5.00
Begin VB.Form frmControl 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control"
   ClientHeight    =   4410
   ClientLeft      =   420
   ClientTop       =   10305
   ClientWidth     =   3390
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   3600
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   3375
      Begin VB.TextBox aspins 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Text            =   "5"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox aslength 
         Height          =   285
         Left            =   1200
         TabIndex        =   20
         Text            =   "540"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox aangle 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Text            =   "45"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox alength 
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Text            =   "1200"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Density:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Wideness:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Shape angle:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Line length:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2640
         Picture         =   "frmControl.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2640
         Picture         =   "frmControl.frx":0CCA
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2640
         Picture         =   "frmControl.frx":1994
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2640
         Picture         =   "frmControl.frx":265E
         Top             =   2400
         Width           =   480
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Create"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   0
      TabIndex        =   10
      Top             =   3000
      Width           =   3375
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Text            =   "r"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Text            =   "r"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "r"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "B"
         Height          =   255
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "G"
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "R"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton ani 
         Caption         =   "Animate"
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "3000"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "30"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "8"
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   3120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   2640
         Picture         =   "frmControl.frx":3328
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2640
         Picture         =   "frmControl.frx":3FF2
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2640
         Picture         =   "frmControl.frx":4CBC
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "  Points:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Lines on line:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Line Length:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   0
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu newa 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu createa 
         Caption         =   "&Create"
         Shortcut        =   {F5}
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu overide 
         Caption         =   "&Over-ride"
         Shortcut        =   {F3}
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu star 
      Caption         =   "&Star Type"
      Begin VB.Menu grid 
         Caption         =   "&Grid Star"
         Shortcut        =   {F11}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu geo 
         Caption         =   "Geo &Star"
         Checked         =   -1  'True
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    'Declare PI
    Const PI = 3.141592654
    
    Dim gstar As Boolean
    Dim override As Boolean
    Dim intet As Integer

Private Sub Command10_Click()

    frmMain.P1.Cls

End Sub

Private Sub Command11_Click()

    If gstar = False Then Gridstar
    If gstar = True Then GeoStar

End Sub

Private Sub Command9_Click()

    Unload Me
    Unload frmMain
    End

End Sub

Private Sub createa_Click()

    If gstar = False Then Gridstar
    If gstar = True Then GeoStar
    
End Sub

Private Sub exit_Click()

    Unload Me
    Unload frmMain
    End

End Sub


Private Sub Form_Load()

    Me.Top = frmMain.Height - frmControl.Height + 3000
    Me.Left = 360
    frmMain.Show
    gstar = True
    sreset

End Sub

Private Sub ani_Click()

    'Declare variables
    Dim hello

    'Check if start is greator than screen
    If (CDbl(Text2.Text) * 2) > frmMain.Height - 120 Then
        
        'Provide Warning
        hello = MsgBox("Warning: The size of the star is greater than the screen. Continue?", vbYesNo + vbExclamation, "Warning")
        
        If hello = vbYes Then
            'Start Counter
            Timer1.Enabled = True
            Timer1.Interval = 500
        Else
            'Exit Sub
            Exit Sub
        End If
        
    End If
    
    'Start Counter
    Timer1.Enabled = True
    Timer1.Interval = 500
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    Unload frmMain
    End

End Sub

Private Sub geo_Click()

    Frame2.Visible = True
    Frame1.Visible = False
    geo.Checked = True
    grid.Checked = False
    gstar = True

End Sub

Private Sub grid_Click()

    Frame1.Visible = True
    Frame2.Visible = False
    geo.Checked = False
    grid.Checked = True
    gstar = False

End Sub

Private Sub newa_Click()

    sreset

End Sub

Private Sub overide_Click()

    'Declare Variable
    Dim msg As String

    If overide.Checked = True Then
    
        'Turn off override
        override = False
        overide.Checked = False
        
    Else
                
        'Provide warning msg
        msg = MsgBox("Warning: Override has been activated!" & Chr(13) & Chr(13) & "Override gives you complete freedom over star properties. " & Chr(13) & "Using high numbers could cause your computer to freeze, " & Chr(13) & "So please use at your own risk!" & Chr(13) & Chr(13) & "Do you wish to Continue?", vbExclamation + vbYesNo, "Override")
        
        'Check results
        If msg = vbYes Then
        
            'Turn On Override
            override = True
            overide.Checked = True
            
        Else
            
            'Turn Off Override
            override = False
            overide.Checked = False
            
        End If
        
    End If
    

End Sub

Private Sub Timer1_Timer()

    'Declare Variables
    Dim x1
    Dim x2
    Dim y1
    Dim y2
    Dim Midx
    Dim Midy
    Dim inte
    Dim length
    Dim points
    Dim r
    Dim g
    Dim b
    Dim temp
    Dim angle
    Dim X
    Dim Y
      
    'Declare points
    points = CDbl(Text7.Text)
    angle = 360 / points
        
    'Declare density and size of star
    length = CDbl(Text2.Text)
    inte = CDbl(Text1.Text)
                        
    'Get middle of form
    Midx = frmMain.Width / 2
    Midy = frmMain.Height / 2


    'Is animation complete?
    If intet = inte Then
        Timer1.Enabled = False
        intet = 0
        Exit Sub
    End If
    

    'Animation
    intet = intet + 1
    
    'Clear screen
    frmMain.P1.Cls

    'Create Basic outline shape of star
    For X = 1 To points Step 1
    
        'Determine line co-ordinates
        x1 = Midx
        y1 = Midy
        x2 = length * Cos(PI / 180 * (angle * X - 90)) + Midx '(Thanks to K. O. Thaha Hussain for this code snippet!)
        y2 = length * Sin(PI / 180 * (angle * X - 90)) + Midy '(Thanks to K. O. Thaha Hussain for this code snippet!)

            'Detect if colour red should be random or not
            If Text3.Text = "r" Then
                Randomize
                r = 255 * Rnd
            Else
                r = CDbl(Text3.Text)
            End If
            
            'Detect if colour green should be random or not
            If Text4.Text = "r" Then
                Randomize
                g = 255 * Rnd
            Else
                g = CDbl(Text4.Text)
            End If
            
            'Detect if colour blue should be random or not
            If Text5.Text = "r" Then
                Randomize
                b = 255 * Rnd
            Else
                b = CDbl(Text5.Text)
            End If
        
        'Draw line
        frmMain.P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)

    Next X


    'Determine which pie section of the star to draw
    For X = 1 To points Step 1
    
        'Determine which lines to draw
        For Y = 0 To length Step (length / intet)
        
            'Determine line co-ordinates
            x1 = (length - Y) * Cos(PI / 180 * (angle * X - 90)) + Midx
            y1 = (length - Y) * Sin(PI / 180 * (angle * X - 90)) + Midy
            x2 = (length - (length - Y)) * Cos(PI / 180 * (angle * (X + 1) - 90)) + Midx
            y2 = (length - (length - Y)) * Sin(PI / 180 * (angle * (X + 1) - 90)) + Midy

                'Detect if colour red should be random or not
                If Text3.Text = "r" Then
                    Randomize
                    r = 255 * Rnd
                Else
                    r = CDbl(Text3.Text)
                End If
                
                'Detect if colour green should be random or not
                If Text4.Text = "r" Then
                    Randomize
                    g = 255 * Rnd
                Else
                    g = CDbl(Text4.Text)
                End If
                
                'Detect if colour blue should be random or not
                If Text5.Text = "r" Then
                    Randomize
                    b = 255 * Rnd
                Else
                    b = CDbl(Text5.Text)
                End If
                
            'Draw line
            frmMain.P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)

        Next Y
    Next X

End Sub

Public Function sreset()

    Text3.Text = "r"
    Text4.Text = "r"
    Text5.Text = "r"
    alength.Text = 2400
    aangle.Text = 120
    aslength.Text = 540
    aspins.Text = 50
    Text7.Text = 8
    Text1.Text = 30
    Text2.Text = 3000
    frmMain.P1.Cls
    
End Function

Private Sub Text1_Gotfocus()

    'Select all of the text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text1.Text) > 100 Then Text1.Text = "100"
    If CDbl(Text1.Text) < 1 Then Text1.Text = "1"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Text2_Gotfocus()

    'Select all of the text
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text2.Text) < 500 Then Text2.Text = "500"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Text3_Gotfocus()

    'Select all of the text
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_LostFocus()
    
    'If override is active, exit
    If override = True Then Exit Sub
    If Text3.Text = "r" Then Exit Sub
        
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text3.Text) > 255 Then Text3.Text = "255"
    If CDbl(Text3.Text) < 0 Then Text3.Text = "0"
    
    Exit Sub
    
ErrorHandler:

    
End Sub

Private Sub Text4_Gotfocus()

    'Select all of the text
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)

End Sub



Private Sub Text4_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    If Text4.Text = "r" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text4.Text) > 255 Then Text4.Text = "255"
    If CDbl(Text4.Text) < 0 Then Text4.Text = "0"
    
    Exit Sub
    
ErrorHandler:
    

End Sub

Private Sub Text5_GotFocus()

    'Select all of the text
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)

End Sub


Private Sub Text5_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    If Text5.Text = "r" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text5.Text) > 255 Then Text5.Text = "255"
    If CDbl(Text5.Text) < 0 Then Text5.Text = "0"
    
    Exit Sub
    
ErrorHandler:


End Sub

Private Sub Text7_Gotfocus()

    'If override is active, exit
    If override = True Then Exit Sub
    
    'Select all of the text
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7.Text)
    
End Sub


Private Sub Text7_LostFocus()

    'If override is active, exit
    If override = True Then Exit Sub
    On Error GoTo ErrorHandler
    
    'Check to see if value is greator than limit
    If CDbl(Text7.Text) > 1080 Then Text7.Text = "1080"
    If CDbl(Text7.Text) < 3 Then Text7.Text = "3"
    Exit Sub
    
ErrorHandler:
    

End Sub


