Attribute VB_Name = "mdlLight"

'APIs:
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Variables (these are needed for GetRGBs):
Dim Red As Long
Dim Green As Long
Dim Blue As Long

Public Sub DrawMaskedLight(Target As Long, Mask As Long, MaskWidth As Long, MaskHeight As Long, X As Integer, Y As Integer, RedB As Byte, GreenB As Byte, BlueB As Byte, Radius As Integer)
    Dim cX As Integer 'X Counter
    Dim cY As Integer 'Y Counter
    Dim TempColor As Long
    
    'This boolean array holds the state of
    'each pixel (if it was drawn or not).
    'It's useful to not draw a pixel twice.
    Dim Done() As Boolean
    ReDim Done(-Radius - 1 To MaskWidth + Radius + 1, -Radius - 1 To MaskWidth + Radius + 1)
    
    'Initialize the TempMask array with the
    'Mask. It will be enlarged in each "i"
    'loop.
    Dim TempMask() As Boolean
    ReDim TempMask(-Radius - 1 To MaskWidth + Radius + 1, -Radius - 1 To MaskWidth + Radius + 1)
    
    'This is the first time we draw in the image, so
    'we need to read from the mask. It initializes
    'the "done" array and the first drawing step (the
    'inside of the image).
    For cX = -Radius To MaskWidth + Radius
        For cY = -Radius To MaskHeight + Radius
            TempColor = GetPixel(Mask, cX, cY)
            If TempColor = vbWhite Then
                'Get the pixel and extract RGBs
                TempColor = GetPixel(Target, cX + X, cY + Y)
                GetRGBs TempColor
                
                'Increase RGB values with the given brightnesses
                Red = Red + RedB * Radius
                If Red > 255 Then Red = 255
                
                Green = Green + GreenB * Radius
                If Green > 255 Then Green = 255
                
                Blue = Blue + BlueB * Radius
                If Blue > 255 Then Blue = 255
                
                'Draw pixel
                SetPixel Target, cX + X, cY + Y, RGB(Red, Green, Blue)
                
                'This pixel has been drawn, so write it in the
                'array (so it won't be drawn the next time)
                Done(cX, cY) = True
            End If
        Next cY
    Next cX
    
    For i = 1 To Radius
        'Update the TempMask (with the Done array; this
        'is because there would be some conflicts if only
        'the Done array was used)
        For cX = -Radius To MaskWidth + Radius
            For cY = -Radius To MaskHeight + Radius
                If Done(cX, cY) Then
                    TempMask(cX, cY) = True
                End If
            Next cY
        Next cX
        
        'Draw the effect
        For cX = -Radius To MaskWidth + Radius
            For cY = -Radius To MaskHeight + Radius
                If Not Done(cX, cY) Then 'The pixel hasn't been drawn yet
                    'Check if there are "true" pixels nearby in the TempMask:
                    If TempMask(cX - 1, cY) Or TempMask(cX + 1, cY) Or _
                      TempMask(cX, cY - 1) Or TempMask(cX, cY + 1) Then
                        'Get the pixel and extract RGBs
                        TempColor = GetPixel(Target, cX + X, cY + Y)
                        GetRGBs TempColor
                        
                        'Increase RGB values with the given brightnesses
                        Red = Red + RedB * (Radius - i)
                        If Red > 255 Then Red = 255
                        
                        Green = Green + GreenB * (Radius - i)
                        If Green > 255 Then Green = 255
                        
                        Blue = Blue + BlueB * (Radius - i)
                        If Blue > 255 Then Blue = 255
                        
                        'Draw pixel
                        SetPixel Target, cX + X, cY + Y, RGB(Red, Green, Blue)
                        
                        'This pixel has been drawn, so write it in the
                        'array (so it won't be drawn the next time)
                        Done(cX, cY) = True
                    End If
                End If
            Next cY
        Next cX
    Next i
End Sub

    ' This function extracts the Red, Green and Blue
    'values from a color to 3 variables with their
    'names. You may keep it as well, but please give
    'me some credit for it too ;)
    'By the way, this function isn't very exact, so
    'you may end up with the green and blue values
    'not as they should (plus or minus 1 or 2).

Public Sub GetRGBs(RGBVal As Long)

    If RGBVal = 16777215 Then
        Red = 255
        Green = 255
        Blue = 255
        Exit Sub
    End If

    Red = RGBVal And 255
    Green = RGBVal / 256 And 255
    Blue = RGBVal / 65536 And 255
    
End Sub

