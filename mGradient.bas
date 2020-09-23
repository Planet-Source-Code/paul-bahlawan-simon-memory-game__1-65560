Attribute VB_Name = "mGradient"
'source of this module: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=60580&lngWId=1

'================================================
' Module:        mGradient.bas (Circular)
' Author:        Carles P.V. - 2005
' Dependencies:  None
' Last revision: Dec 13th, 2005
'                Improved algorithm:
'                SQR completely avoided
'================================================
' = ========================================== =
'================================================
' Module:        mGradient.bas (Linear, any angle)
' Author:        Carles P.V. - 2005
' Dependencies:  None
' Last revision: 2005.05.21
' ==============================================
'
' 2005.05.21: Minor speed improvements
'             Thanks to Robert Rayment.
'
'================================================

Option Explicit

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'//

Private Const PI      As Single = 3.14159265358979
Private Const TO_DEG  As Single = 180 / PI
Private Const TO_RAD  As Single = PI / 180
Private Const INT_ROT As Long = 100 ' Increase this value for more precision


'Circular
Public Sub PaintGradientC(ByVal hdc As Long, _
                         ByVal x As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long _
                         )

  Dim uBIH      As BITMAPINFOHEADER
  Dim lBits()   As Long
  Dim lGrad()   As Long
  Dim g         As Long
  
  Dim R1        As Long
  Dim G1        As Long
  Dim B1        As Long
  Dim R2        As Long
  Dim G2        As Long
  Dim B2        As Long
  Dim dR        As Long
  Dim dG        As Long
  Dim dB        As Long
  
  Dim Scan      As Long
  Dim Offset1   As Long
  Dim Offset2   As Long
  Dim iEnd      As Long
  Dim jEnd      As Long
  Dim iPad      As Long
  Dim jPad      As Long
  
  Dim i         As Long ' squares axis sum accumulators (-> x^2+y^2)
  Dim ia        As Long
  Dim iaa       As Long
  Dim j         As Long
  Dim ja        As Long
  Dim jaa       As Long
  
  Dim s()       As Long ' squares sequence
  Dim sc        As Long ' squares sequence counter (sequence index -> root)
  
    '-- Minor check
    If (Width > 0 And Height > 0) Then
        
        '-- Calc. gradient length ('diagonal')
        g = Sqr(Width * Width + Height * Height) \ 2
        
        '-- Decompose colors
        R1 = (Color1 And &HFF&)
        G1 = (Color1 And &HFF00&) \ 256
        B1 = (Color1 And &HFF0000) \ 65536
        R2 = (Color2 And &HFF&)
        G2 = (Color2 And &HFF00&) \ 256
        B2 = (Color2 And &HFF0000) \ 65536
        
        '-- Get color distances
        dR = R2 - R1
        dG = G2 - G1
        dB = B2 - B1
        
        '-- Size gradient-colors array
        ReDim lGrad(0 To g)
        
        '-- Build squares sequence LUT
        ReDim s(0 To g)
        For i = 1 To g
            s(i) = s(i - 1) + i + i - 1
        Next i
        
        '-- Calculate gradient-colors
        If (g = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
          Else
            For i = 0 To g
                lGrad(i) = B1 + (dB * i) \ g + 256 * (G1 + (dG * i) \ g) + 65536 * (R1 + (dR * i) \ g)
            Next i
        End If
        
        '-- Size DIB array
        ReDim lBits(Width * Height - 1) As Long
        
        '== Render gradient DIB
        
        '-- First "quadrant"...
        
        Scan = Width
        iPad = Width Mod 2
        jPad = Height Mod 2
        
        iEnd = Scan \ 2 + iPad - 1
        jEnd = Height \ 2 + jPad - 1
        Offset1 = jEnd * Scan + Scan \ 2
        
        ja = 1
        jaa = -1
        For j = 0 To jEnd
            sc = j
            ja = ja + jaa
            jaa = jaa + 2
            ia = ja + 1
            iaa = -1
            For i = Offset1 To Offset1 + iEnd
                ia = ia + iaa
                iaa = iaa + 2
                lBits(i) = lGrad(sc)
                If (ia >= s(sc) - sc) Then
                    sc = sc + 1
                End If
            Next i
            Offset1 = Offset1 - Scan
        Next j
        
        '-- Mirror first "quadrant"
        
        iEnd = iEnd - iPad
        Offset1 = 0
        Offset2 = Scan - 1

        For j = 0 To jEnd
            For i = 0 To iEnd
                lBits(Offset1 + i) = lBits(Offset2 - i)
            Next i
            Offset1 = Offset1 + Scan
            Offset2 = Offset2 + Scan
        Next j
        
        '-- Mirror first "half"

        iEnd = Scan - 1
        jEnd = jEnd - jPad
        Offset1 = (Height - 1) * Scan
        Offset2 = 0

        For j = 0 To jEnd
            For i = 0 To iEnd
                lBits(Offset1 + i) = lBits(Offset2 + i)
            Next i
            Offset1 = Offset1 - Scan
            Offset2 = Offset2 + Scan
        Next j
        
        '-- Define DIB header
        With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height
        End With
        
        '-- Paint it!
        Call StretchDIBits(hdc, x, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
    End If
End Sub


'Linear, any angle
Public Sub PaintGradientL(ByVal hdc As Long, _
                         ByVal x As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal Angle As Single _
                         )

  Dim uBIH      As BITMAPINFOHEADER
  Dim lBits()   As Long
  Dim lGrad()   As Long
  
  Dim lClr      As Long
  Dim R1        As Long
  Dim G1        As Long
  Dim B1        As Long
  Dim R2        As Long
  Dim G2        As Long
  Dim B2        As Long
  Dim dR        As Long
  Dim dG        As Long
  Dim dB        As Long
  
  Dim Scan      As Long
  Dim i         As Long
  Dim j         As Long
  Dim iIn       As Long
  Dim jIn       As Long
  Dim iEnd      As Long
  Dim jEnd      As Long
  Dim Offset    As Long
  
  Dim lQuad     As Long
  Dim AngleDiag As Single
  Dim AngleComp As Single
  
  Dim g         As Long
  Dim luSin     As Long
  Dim luCos     As Long
 
    '-- Minor check
    If (Width > 0 And Height > 0) Then
        
        '-- Right-hand [+] (ox=0º)
        Angle = -Angle + 90
        
        '-- Normalize to [0º;360º]
        Angle = Angle Mod 360
        If (Angle < 0) Then Angle = 360 + Angle
        
        '-- Get quadrant (0 - 3)
        lQuad = Angle \ 90
        
        '-- Normalize to [0º;90º]
        Angle = Angle Mod 90
        
        '-- Calc. gradient length ('distance')
        If (lQuad Mod 2 = 0) Then
            AngleDiag = Atn(Width / Height) * TO_DEG
          Else
            AngleDiag = Atn(Height / Width) * TO_DEG
        End If
        AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
        Angle = Angle * TO_RAD
        g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem
        
        '-- Decompose colors
        If (lQuad > 1) Then
            lClr = Color1
            Color1 = Color2
            Color2 = lClr
        End If
        R1 = (Color1 And &HFF&)
        G1 = (Color1 And &HFF00&) \ 256
        B1 = (Color1 And &HFF0000) \ 65536
        R2 = (Color2 And &HFF&)
        G2 = (Color2 And &HFF00&) \ 256
        B2 = (Color2 And &HFF0000) \ 65536
        
        '-- Get color distances
        dR = R2 - R1
        dG = G2 - G1
        dB = B2 - B1
        
        '-- Size gradient-colors array
        ReDim lGrad(0 To g - 1)
        
         '-- Calculate gradient-colors
        iEnd = g - 1
        If (iEnd = 0) Then
            '-- Special case (1-pixel wide gradient)
            lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
          Else
            For i = 0 To iEnd
                lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
            Next i
        End If
        
        '-- Size DIB array
        ReDim lBits(Width * Height - 1) As Long
        
        '-- Render gradient DIB
        
        iEnd = Width - 1
        jEnd = Height - 1
        
        Select Case lQuad
        
            Case 0, 2
            
                luSin = Sin(Angle) * INT_ROT
                luCos = Cos(Angle) * INT_ROT
                Offset = 0
                Scan = Width
                
            Case 1, 3
            
                luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
                luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
                Offset = jEnd * Width
                Scan = -Width
        End Select
        
        jIn = 0
        iIn = 0
        For j = 0 To jEnd
            iIn = jIn
            For i = 0 To iEnd
                lBits(i + Offset) = lGrad(iIn \ INT_ROT)
                iIn = iIn + luSin
            Next i
            jIn = jIn + luCos
            Offset = Offset + Scan
        Next j
                
        '-- Define DIB header
        With uBIH
            .biSize = 40
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = Height
        End With
        
        '-- Paint it!
        Call StretchDIBits(hdc, x, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
    End If
End Sub
