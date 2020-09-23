Attribute VB_Name = "GlyphOutline"
Option Explicit

' Notes:
'   -GetOutline returns the outline upside down (as it is returned from the
'    api function)
'   -Polygon type TT_PRIM_CSPLINE is untested as I have no fonts that return
'    this type. (Perhaps this is a Windows 2000 or better type only?)

' These constants determine how many points are generated when a spline is
' encountered in a font.
Private Const QSPLINE_COUNT As Long = 10
Private Const CSPLINE_COUNT As Long = 18

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type TEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
End Type

Private Type PANOSE
   ulculture As Long
   bFamilyType As Byte
   bSerifStyle As Byte
   bWeight As Byte
   bProportion As Byte
   bContrast As Byte
   bStrokeVariation As Byte
   bArmStyle As Byte
   bLetterform As Byte
   bMidline As Byte
   bXHeight As Byte
End Type


Private Type OUTLINETEXTMETRIC
   otmSize As Long
   otmTextMetrics As TEXTMETRIC
   otmFiller As Byte
   otmPanoseNumber As PANOSE
   otmfsSelection As Long
   otmfsType As Long
   otmsCharSlopeRise As Long
   otmsCharSlopeRun As Long
   otmItalicAngle As Long
   otmEMSquare As Long
   otmAscent As Long
   otmDescent As Long
   otmLineGap As Long
   otmsCapEmHeight As Long
   otmsXHeight As Long
   otmrcFontBox As RECT
   otmMacAscent As Long
   otmMacDescent As Long
   otmMacLineGap As Long
   otmusMinimumPPEM As Long
   otmptSubscriptSize As POINTAPI
   otmptSubscriptOffset As POINTAPI
   otmptSuperscriptSize As POINTAPI
   otmptSuperscriptOffset As POINTAPI
   otmsStrikeoutSize As Long
   otmsStrikeoutPosition As Long
   otmsUnderscorePosition As Long
   otmsUnderscoreSize As Long
   otmpFamilyName As Long 'String pointer
   otmpFaceName As Long 'String pointer
   otmpStyleName As Long 'String pointer
   otmpFullName As Long 'String pointer
   buffer As String * 256
End Type

Private Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As Long, ByVal cbData As Long, lpotm As OUTLINETEXTMETRIC) As Long


Private Type MAT2
   eM11 As FIXED
   eM12 As FIXED
   eM21 As FIXED
   eM22 As FIXED
End Type


Private Type GLYPHMETRICS
   gmBlackBoxX As Long
   gmBlackBoxY As Long
   gmptGlyphOrigin As POINTAPI
   gmCellIncX As Integer
   gmCellIncY As Integer
End Type



Private Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As Long, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long
Private Const GGO_NATIVE = 2
Private Const GGO_METRICS = 0


' These structures are not used directly, but are here
' for reference.

'Private Type TTPOLYCURVE
'   wType As Integer 'curve type.
'   cpfx As Integer  'number of pointfx structures
'   apfx As POINTFX
'End Type

'Private Type TTPOLYGONHEADER
'   cb As Long
'   dwType As Long
'   pfxStart As POINTFX
'End Type

'Private Type POINTFX
'   x As FIXED
'   y As FIXED
'End Type

'Private Type FIXED
'   fract As Integer
'   Value As Integer
'End Type


Private Const TT_PRIM_CSPLINE As Long = &H3
Private Const TT_PRIM_QSPLINE As Long = &H2
Private Const TT_PRIM_LINE As Long = &H1
Private Const TT_POLYGON_TYPE As Long = 24


'/****************************************************************************
' *  FUNCTION   : IdentityMat
' *  PURPOSE    : Fill in matrix to be the identity matrix.
' *  RETURNS    : none.
' ****************************************************************************/
Private Function IdentityMat() As MAT2
   Static lpMat As MAT2
   
   If lpMat.eM11.Value = 0 Then
      lpMat.eM11 = FixedFromDouble(1)
      lpMat.eM12 = FixedFromDouble(0)
      lpMat.eM21 = FixedFromDouble(0)
      lpMat.eM22 = FixedFromDouble(1)
   End If
   
   IdentityMat = lpMat
End Function

'
' Returns unit size the font is defined in.  If 0 is returned
' then the font is not a truetype font and won't work with
' GetOutline.
'
Public Function GetEMUnit(ByVal hdc As Long) As Long
   Dim tm As OUTLINETEXTMETRIC
   
   Call GetOutlineTextMetrics(hdc, Len(tm), tm)
   GetEMUnit = Abs(tm.otmEMSquare)
End Function

Public Sub GetOutline(ByVal hdc As Long, ByVal Letter As Long, x() As Double, y() As Double, pCount As Long, p() As Long, polyCount As Long)
   Dim size As Long
   Dim gm As GLYPHMETRICS
   Dim buffer() As Long
   
   size = GetGlyphOutline(hdc, Letter, GGO_NATIVE, gm, 0, ByVal 0&, IdentityMat)
   
   If size > 0 Then
      size = size \ 4&
      ReDim buffer(0 To size)
      
      If GetGlyphOutline(hdc, Letter, GGO_NATIVE, gm, size * 4&, buffer(0), IdentityMat) > 0 Then
         Call GetBufferOutline(buffer, x, y, pCount, p, polyCount)
      End If
   End If
End Sub

' This function gets the glyph outline data and calculates the number of points
' and polygons that will be generated when the outline polyline is generated.
' buffer() contains the glyph data.
Public Sub GetOutlineCount(ByVal hdc As Long, ByVal Letter As Long, buffer() As Long, ByRef pCount As Long, ByRef polyCount As Long)
   Dim size As Long
   Dim gm As GLYPHMETRICS
   Dim i As Long
   Dim cb As Long
   Dim cp As Long
   
   size = GetGlyphOutline(hdc, Letter, GGO_NATIVE, gm, 0, ByVal 0&, IdentityMat)
   
   If size > 0 Then
      size = size \ 4&
      ReDim buffer(0 To size)
      
      If GetGlyphOutline(hdc, Letter, GGO_NATIVE, gm, size * 4&, buffer(0), IdentityMat) > 0 Then
         i = 0
         Do
            cb = i + buffer(i) \ 4 ' get the end of the polyline
            i = i + 1
            
            If buffer(i) = TT_POLYGON_TYPE Then
               ' this is a valid polygon
               pCount = pCount + 1
               i = i + 3
               
               ' count the elements of this polyline
               Do While i < cb
                  Select Case buffer(i) And &HF  'And &HFFFF0000
                     Case TT_PRIM_CSPLINE:
                        cp = buffer(i) \ &H10000
                        i = i + 1 + cp * 2
                        pCount = pCount + (((cp - 1) / 2) * (CSPLINE_COUNT - 1))
                        
                     Case TT_PRIM_QSPLINE:
                        cp = buffer(i) \ &H10000
                        i = i + 1 + cp * 2
                        pCount = pCount + ((cp - 1) * (QSPLINE_COUNT - 1))
                        
                     Case TT_PRIM_LINE:
                        ' Find the number of points in the line
                        cp = buffer(i) \ &H10000
                        i = i + 1 + cp * 2
                        pCount = pCount + cp
                        
                        
                     Case Else
                        'Skip this record, we don't know
                        'what to do with it.
                        i = i + 1 + (buffer(i) And &HFFFF&)
                  End Select
               Loop
           
               ' get the last point (same as the first)
               pCount = pCount + 1
               
               polyCount = polyCount + 1
            Else
               Erase buffer
               Exit Sub
            End If
            
         Loop While i < size
      End If
   End If
End Sub


' Given a buffer as returned from GetGlyphOutline, this calculates the
' polyline that represents the letter.
Public Sub GetBufferOutline(buffer() As Long, x() As Double, y() As Double, pCount As Long, p() As Long, polyCount As Long)
   Dim size As Long
   Dim i As Long
   Dim cb As Long
   Dim cp As Long
   Dim x2 As Double
   Dim y2 As Double
   Dim x3 As Double
   Dim y3 As Double
   Dim X4 As Double
   Dim Y4 As Double
   Dim pStart As Long
   Dim xStart As Double
   Dim yStart As Double
   
   size = UBound(buffer)
   
   i = 0
   Do
      cb = i + buffer(i) \ 4 ' get the end of the polyline
      i = i + 1
      
      If buffer(i) = TT_POLYGON_TYPE Then
         ' this is a valid polygon
         pStart = pCount
         
         
         ' get the first point
         i = i + 1
         xStart = DoubleFromFixedAsLong(buffer(i))
         x(pCount) = xStart
         i = i + 1
         yStart = DoubleFromFixedAsLong(buffer(i))
         y(pCount) = yStart
         i = i + 1
         pCount = pCount + 1
         
         ' get the elements of this polyline
         Do While i < cb
            Select Case buffer(i) And &HF  'And &HFFFF0000
               Case TT_PRIM_CSPLINE:
                  cp = buffer(i) \ &H10000
                  
                  i = i + 1
                  Do While cp > 0
                     ' CubicBezier starts with the last good point
                     ' so decrease the pCount by 1
                     pCount = pCount - 1
                     x2 = DoubleFromFixedAsLong(buffer(i))
                     y2 = DoubleFromFixedAsLong(buffer(i + 1))
                     
                     x3 = DoubleFromFixedAsLong(buffer(i + 2))
                     y3 = DoubleFromFixedAsLong(buffer(i + 3))
                     
                     X4 = DoubleFromFixedAsLong(buffer(i + 4))
                     Y4 = DoubleFromFixedAsLong(buffer(i + 5))
                     
                     If cp > 3 Then
                        X4 = (X4 + x3) / 2#
                        Y4 = (Y4 + y3) / 2#
                        i = i + 4
                        cp = cp - 2
                     Else
                        i = i + 6
                        cp = cp - 3
                     End If
                     
                     Call CubicBezier(x, y, pCount, QSPLINE_COUNT, _
                        x(pCount), y(pCount), x2, y2, x3, y3, X4, Y4)
                     
                  Loop
                  

                  
               Case TT_PRIM_QSPLINE:
                  cp = buffer(i) \ &H10000
                  i = i + 1
                  
                  Do While cp > 0
                     ' QuadraticBezier starts with the last good point
                     ' so decrease the pCount by 1
                     pCount = pCount - 1
                     x2 = DoubleFromFixedAsLong(buffer(i))
                     y2 = DoubleFromFixedAsLong(buffer(i + 1))
                     
                     x3 = DoubleFromFixedAsLong(buffer(i + 2))
                     y3 = DoubleFromFixedAsLong(buffer(i + 3))
                     
                     If cp > 2 Then
                        x3 = (x3 + x2) / 2#
                        y3 = (y3 + y2) / 2#
                        i = i + 2
                        cp = cp - 1
                     Else
                        i = i + 4
                        cp = cp - 2
                     End If
                     
                     Call QuadraticBezier(x, y, pCount, QSPLINE_COUNT, _
                        x(pCount), y(pCount), x2, y2, x3, y3)
                     
                  Loop
                  
                  
               Case TT_PRIM_LINE:
                  ' Find the number of points in the line
                  cp = buffer(i) \ &H10000
                  i = i + 1
                  
                  Do While cp > 0
                     x(pCount) = DoubleFromFixedAsLong(buffer(i))
                     i = i + 1
                     y(pCount) = DoubleFromFixedAsLong(buffer(i))
                     i = i + 1
                     pCount = pCount + 1
                     cp = cp - 1
                  Loop
                  
               Case Else
                  'Skip this record, we don't know
                  'what to do with it.
                  i = i + 1 + (buffer(i) And &HFFFF&)
            End Select
         Loop
     
         ' get the last point (same as the first)
         x(pCount) = xStart
         y(pCount) = yStart
         pCount = pCount + 1
         
         p(polyCount) = pCount - pStart
         polyCount = polyCount + 1
      Else
         Exit Sub
      End If
      
   Loop While i < size

End Sub

