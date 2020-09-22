Attribute VB_Name = "GeoStars"
Option Explicit

    
    'Declare PI
    Const PI = 3.141592654

Public Function Gridstar()

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
    Dim msg
    
    'Check if start is greator than screen
    If (CDbl(frmControl.Text2.Text) * 2) > frmMain.Height - 120 Then
        
        'Provide Warning
        msg = MsgBox("Warning: The size of the star is greater than the screen. Continue?", vbYesNo + vbExclamation, "Warning")
        
            If msg = vbYes Then
                'Continue
            Else
                'Exit Sub
                Exit Function
            End If
        
    End If
        
    'Clear screen
    frmMain.P1.Cls
     
    'Declare points
    points = CDbl(frmControl.Text7.Text)
    angle = 360 / points
        
    'Declare density and size of star
    length = CDbl(frmControl.Text2.Text)
    inte = CDbl(frmControl.Text1.Text)
                        
    'Get middle of form
    Midx = frmMain.Width / 2
    Midy = frmMain.Height / 2


    'Create Basic outline shape of star
    For X = 1 To points Step 1
    
        'Determine line co-ordinates
        x1 = Midx
        y1 = Midy
        x2 = length * Cos(PI / 180 * (angle * X - 90)) + Midx
        y2 = length * Sin(PI / 180 * (angle * X - 90)) + Midy

            'Detect if colour red should be random or not
            If frmControl.Text3.Text = "r" Then
                Randomize
                r = 255 * Rnd
            Else
                r = CDbl(frmControl.Text3.Text)
            End If
            
            'Detect if colour green should be random or not
            If frmControl.Text4.Text = "r" Then
                Randomize
                g = 255 * Rnd
            Else
                g = CDbl(frmControl.Text4.Text)
            End If
            
            'Detect if colour blue should be random or not
            If frmControl.Text5.Text = "r" Then
                Randomize
                b = 255 * Rnd
            Else
                b = CDbl(frmControl.Text5.Text)
            End If
        
        'Draw line
        frmMain.P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)

    Next X


    'Determine which pie section of the star to draw
    For X = 1 To points Step 1
    
        'Determine which lines to draw
        For Y = 0 To length Step (length / inte)
        
            'Determine line co-ordinates
            x1 = (length - Y) * Cos(PI / 180 * (angle * X - 90)) + Midx
            y1 = (length - Y) * Sin(PI / 180 * (angle * X - 90)) + Midy
            x2 = (length - (length - Y)) * Cos(PI / 180 * (angle * (X + 1) - 90)) + Midx
            y2 = (length - (length - Y)) * Sin(PI / 180 * (angle * (X + 1) - 90)) + Midy

                'Detect if colour red should be random or not
                If frmControl.Text3.Text = "r" Then
                    Randomize
                    r = 255 * Rnd
                Else
                    r = CDbl(frmControl.Text3.Text)
                End If
                
                'Detect if colour green should be random or not
                If frmControl.Text4.Text = "r" Then
                    Randomize
                    g = 255 * Rnd
                Else
                    g = CDbl(frmControl.Text4.Text)
                End If
                
                'Detect if colour blue should be random or not
                If frmControl.Text5.Text = "r" Then
                    Randomize
                    b = 255 * Rnd
                Else
                    b = CDbl(frmControl.Text5.Text)
                End If
                
            'Draw line
            frmMain.P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)

        Next Y
    Next X

End Function


Public Function GeoStar()

    frmMain.P1.Cls

    Dim angle, X, r, g, b, x1, x2, y1, y2, length, b1, b2, Y, spins, slength
    
    length = frmControl.alength.Text
    angle = frmControl.aangle.Text
    spins = frmControl.aspins.Text
    slength = frmControl.aslength.Text
    X = 2
    r = 0
    b = 0
    g = 0
    
    angle = CInt(angle)
    length = CInt(length)
    spins = CInt(spins)
    slength = CInt(slength)
    
    'Determine line co-ordinates


For Y = 0 To 360 Step (360 / spins)

        b1 = slength * Cos(PI / 180 * (Y)) + frmMain.Height / 2
        b2 = slength * Sin(PI / 180 * (Y)) + frmMain.Width / 2
        x2 = b2
        y2 = b1

    For X = 0 To 360 Step 1
                
        x1 = x2
        y1 = y2
        
        x2 = length * Cos(PI / 180 * ((angle * X) - Y)) + b2
        y2 = length * Sin(PI / 180 * ((angle * X) - Y)) + b1
        
            If frmControl.Text3.Text = "r" Then
                Randomize
                r = 255 * Rnd
            Else
                r = CDbl(frmControl.Text3.Text)
            End If
            
            'Detect if colour green should be random or not
            If frmControl.Text4.Text = "r" Then
                Randomize
                g = 255 * Rnd
            Else
                g = CDbl(frmControl.Text4.Text)
            End If
            
            'Detect if colour blue should be random or not
            If frmControl.Text5.Text = "r" Then
                Randomize
                b = 255 * Rnd
            Else
                b = CDbl(frmControl.Text5.Text)
            End If
        
        If X <> 0 Then frmMain.P1.Line (x1, y1)-(x2, y2), RGB(r, g, b)
        
    Next X
        
Next Y

End Function
