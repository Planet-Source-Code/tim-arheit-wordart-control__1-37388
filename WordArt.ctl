VERSION 5.00
Begin VB.UserControl WordArt 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   PropertyPages   =   "WordArt.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "WordArt.ctx":0013
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1320
      ScaleHeight     =   1215
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "WordArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit


'
' Notes:
'   - This control will only work with TrueType fonts.
'

'**************************
'** API declarations
'**************************
Private Const WM_PAINT = &HF
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
   (ByVal hdc As Long, ByVal hObject As Long) As Long
   
Private Declare Function DeleteObject Lib "gdi32" _
   (ByVal hObject As Long) As Long
   
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long  'lpPoint = POINTAPI

Private Const WEIGHT_NORMAL = 400
Private Const WEIGHT_BOLD = 700

Private Const LF_FACESIZE = 32

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName As String * LF_FACESIZE
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
   "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Type APISize
   cx As Long
   cy As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias _
   "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, _
   ByVal cbString As Long, lpSize As APISize) As Long

Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, _
   lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long

Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Const MM_TEXT = 1
Private Const MM_TWIPS = 6

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
   ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID = 0

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Type LOGBRUSH
   lbStyle As Long
   lbColor As Long
   lbHatch As Long
End Type

Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Const BS_HOLLOW = 1&

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, _
   lpBits As Any) As Long

Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
   ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
   ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Private Const WM_KEYDOWN = &H100
Private Const KEY_PAINT = 500


'**************************
'** Exposed events
'**************************
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'******************************
'** Internal variables that are
'** exposed though properties.
'******************************
Private mAutosize As Boolean
Private mCaption As String

Private mFontName As String
Private mFontItalic As Boolean
Private mFontSize As Single
Private mFontColor As Long
Private mFontWeight As Long
Private mFontWidth As Single

Private mBackColor As Long
Private mBackStyle As Long

Private mFillColor As Long
Private mFillStyle As Long
Private mGradientStart As Long
Private mGradientEnd As Long
Private mFillPicture As IPictureDisp

Private mLineWidth As Long

Private mFontOffset As Single
Private mOffsetAngle As Single
Private mOffsetColor As Long

Private mRotation As Single

Private mBottomAlignment As Long
Private mTopAlignment As Long

' Coordinates used for the top alignment.
Private mTopX1 As Single
Private mTopX2 As Single
Private mTopX3 As Single
Private mTopY1 As Single
Private mTopY2 As Single
Private mTopY3 As Single


' Coordinates used for the bottom alignment.
Private mBtmX1 As Single
Private mBtmX2 As Single
Private mBtmX3 As Single
Private mBtmY1 As Single
Private mBtmY2 As Single
Private mBtmY3 As Single



'***********************************
'** Internal variables (not exposed)
'***********************************

Private mChangeFont As Boolean           ' True if the font polygons need recalculated.
Private mChangeTransformation As Boolean ' True if the transformation function needs recalculated.
Private mChangeTransFont As Boolean      ' True if the font needs transformed
Private mChangeFinal As Boolean          ' True if the final font representation has changed.
Private mPaintCalled As Boolean          ' True if paint has already been triggered.
Private mInPaint As Boolean              ' True if we are currently in paint.
Private mChangeOffset As Boolean

'*********************************************
'**  Internal variables used for intermediate
'**  calculations.
'*********************************************

'*********************************************
'**  These values hold temporary values
'**  calculated by CalculateFont.  They
'**  are all for the nominal font size
'**  and are used in later calculations
'*********************************************
Private mFontX() As Double
Private mFontY() As Double
Private mFontPoly() As Long
Private mFontCWidth As Double
Private mFontCHeight As Double

'**********************************************
'** The following variable hold transformation
'** variables used to map the font to it's
'** final form.
'**********************************************
Private Const TRANSFORM_COUNT As Long = 100  ' Number of points in the
                                             ' transformation array.
Private mTransTop() As Double
Private mTransBottom() As Double
Private mTransWidth As Double       ' Width of each transformation step

'**********************************************
'** Transformed font variables
'**********************************************
Private mTransX() As Double
Private mTransY() As Double
Private mFontMinX As Double
Private mFontMinY As Double
Private mFontMaxX As Double
Private mFontMaxY As Double


'**********************************************
'** The following is the final transformed
'** polygon information representing the
'** the entire caption.
'**********************************************

Private mFFontX() As Long
Private mFFontY() As Long

'*********************************************
'** Outout polygon for API use with the
'** shift variables
'*********************************************
Private mFFontPoly() As POINTAPI
Private mSFontPoly() As POINTAPI
Private mFFontShiftX As Double
Private mFFontShiftY As Double


'**************************************
'** Public enums used for various
'** public properties.
'**************************************

Public Enum waFillStyle
   waFillNone = 0
   waFillSolid = 1
   waFillTransparent = 2
   waFillGradient = 3
   waFillUserDefined = 4
End Enum

Public Enum waAlignment
   waAlignmentNone = 0
   waAlignmentParallel = 1
   waAlignmentLinear = 2
   waAlignmentArc = 3
   waAlignmentQuadratic = 4
   waAlignmentCubic = 5
   waAlignmentFit3 = 6
   waAlignmentFit4 = 7
   waAlignmentFit5 = 8
End Enum

Public Enum waBackStyle
   waBackTransparent = 0
   waBackOpaque = 1
End Enum


Private Const DEG_TO_RAD = 1.74532925199433E-02

Private Sub UserControl_Initialize()
   ' set default value for all the internal variables.
   mAutosize = False
   mCaption = "WordArt"
   mFontName = "Arial"
   mFontItalic = False
   mFontSize = 12
   mFontWeight = 400
   mFontWidth = 12
   mBackColor = vbWhite
   mBackStyle = 1
   mFillColor = vbBlack
   mFillStyle = waFillSolid
   mGradientStart = vbBlue
   mGradientEnd = vbWhite
   Set mFillPicture = Nothing
   mLineWidth = 1
   mFontOffset = 0
   mOffsetAngle = 0
   mRotation = 0
   mBottomAlignment = waAlignmentNone
   mTopAlignment = waAlignmentNone
   mTopX1 = 0
   mTopX2 = 0
   mTopX3 = 0
   mTopY1 = 0
   mTopY2 = 0
   mTopY3 = 0
   mBtmX1 = 0
   mBtmX2 = 0
   mBtmX3 = 0
   mBtmY1 = 0
   mBtmY2 = 0
   mBtmY3 = 0
   
   UserControl.Font.Weight = mFontWeight
   UserControl.Font.size = mFontSize
   UserControl.Font.Italic = mFontItalic
   UserControl.Font.Name = mFontName
   UserControl.Font.Weight = mFontWeight
   
   mChangeFont = True
   mChangeTransformation = True
   mChangeTransFont = True
   mChangeFinal = True
   mPaintCalled = False
   mChangeOffset = False
   mInPaint = False
End Sub


Private Sub UserControl_InitProperties()
   Call SetPaint
End Sub

Private Sub UserControl_Terminate()
   ' Nothing to clean up yet.
End Sub

' Sets the control to redraw when it can.
Public Sub SetPaint()
   If Ambient.UserMode = False Then
      Call UserControl_Paint
   Else
      If Not mPaintCalled Then
         mPaintCalled = True
         'Call PostMessage(UserControl.hwnd, WM_PAINT, 0, 0)
         Call PostMessage(UserControl.hwnd, WM_KEYDOWN, KEY_PAINT, 0)
      End If
   End If
End Sub

'********************************************
'** Typical property get/let functions
'** No real calculation is done here directly
'** All these do the following:
'**    - Set the new property if it has changed.
'**    - Set the flag to recalculate only
'**      the necessary parts required.
'**    - Call the paint method indirectly.
'********************************************

Public Property Get Font() As IFontDisp
   Set Font = UserControl.Font
End Property

Public Property Set Font(newfont As IFontDisp)
   UserControl.Font.Bold = newfont.Bold
   UserControl.Font.Italic = newfont.Italic
   UserControl.Font.size = newfont.size
   UserControl.Font.Name = newfont.Name
   UserControl.Font.Weight = newfont.Weight
   
   ' Set the other properties from this
   'FINISH
   mFontItalic = newfont.Italic
   mFontSize = newfont.size
   mFontWidth = newfont.size
   mFontName = newfont.Name
   mFontWeight = newfont.Weight
   
   PropertyChanged "FontItalic"
   PropertyChanged "FontSize"
   PropertyChanged "FontWidth"
   PropertyChanged "FontWeight"
   PropertyChanged "FontBold"
   PropertyChanged "FontName"
   PropertyChanged "Font"
   
   mChangeFont = True
   Call SetPaint
End Property

Public Property Get AutoSize() As Boolean
   AutoSize = mAutosize
End Property

Public Property Let AutoSize(ByVal b As Boolean)
   If mAutosize <> b Then
      mAutosize = b
      Call SetPaint
      PropertyChanged "AutoSize"
   End If
End Property

Public Property Get Caption() As String
   Caption = mCaption
End Property

Public Property Let Caption(NewCaption As String)
   If NewCaption <> mCaption Then
      mCaption = NewCaption
      mChangeFont = True
      Call SetPaint
      PropertyChanged "Caption"
   End If
End Property

Public Property Get FontName() As String
   FontName = mFontName
End Property

Public Property Let FontName(newfont As String)
   If newfont <> mFontName Then
      mFontName = newfont
      mChangeFont = True
      Call SetPaint
      UserControl.Font.Name = mFontName
      PropertyChanged "FontName"
      PropertyChanged "Font"
   End If
End Property

Public Property Get FontItalic() As Boolean
   FontItalic = mFontItalic
End Property

Public Property Let FontItalic(ByVal b As Boolean)
   If b <> mFontItalic Then
      mFontItalic = b
      mChangeFont = True
      Call SetPaint
      UserControl.Font.Italic = mFontItalic
      PropertyChanged "FontItalic"
      PropertyChanged "Font"
   End If
End Property

Public Property Get FontBold() As Boolean
   If mFontWeight > (WEIGHT_BOLD + WEIGHT_NORMAL) / 2 Then
      FontBold = True
   Else
      FontBold = False
   End If
End Property

Public Property Let FontBold(ByVal b As Boolean)
   Dim i As Long
   
   If b Then
      i = WEIGHT_BOLD
   Else
      i = WEIGHT_NORMAL
   End If
   
   If i <> mFontWeight Then
      mFontWeight = i
      mChangeFont = True
      Call SetPaint
      UserControl.Font.Weight = mFontWeight
      PropertyChanged "FontBold"
      PropertyChanged "FontWeight"
      PropertyChanged "Font"
   End If
End Property

Public Property Get FontWidth() As Single
   FontWidth = mFontWidth
End Property

Public Property Let FontWidth(ByVal NewWidth As Single)
   If NewWidth < 1 Then NewWidth = 1
   
   If mFontWidth <> NewWidth Then
      mFontWidth = NewWidth
      mChangeFont = True
      Call SetPaint
      PropertyChanged "FontWidth"
   End If
End Property

Public Property Get FontSize() As Single
   FontSize = mFontSize
End Property

Public Property Let FontSize(ByVal NewSize As Single)
   If NewSize < 1 Then NewSize = 1
   
   If mFontSize <> NewSize Then
      mFontSize = NewSize
      mFontWidth = NewSize
      mChangeFont = True
      Call SetPaint
      UserControl.Font.size = mFontSize
      PropertyChanged "FontSize"
      PropertyChanged "FontWidth"
      PropertyChanged "Font"
   End If
End Property

Public Property Get FontWeight() As Long
   FontWeight = mFontWeight
End Property

Public Property Let FontWeight(ByVal NewWeight As Long)
   ' Make sure the weight is in the valid range
   If NewWeight < 0 Or NewWeight > 1000 Then
      NewWeight = WEIGHT_NORMAL ' Normal weight
   End If
   
   If NewWeight <> mFontWeight Then
      mFontWeight = NewWeight
      mChangeFont = True
      Call SetPaint
      UserControl.Font.Weight = mFontWeight
      PropertyChanged "FontWeight"
      PropertyChanged "FontBold"
      PropertyChanged "Font"
   End If
End Property

Public Property Get FontColor() As OLE_COLOR
   FontColor = mFontColor
End Property

Public Property Let FontColor(ByVal NewColor As OLE_COLOR)
   If mFontColor <> NewColor Then
      mFontColor = NewColor
      Call SetPaint
      PropertyChanged "FontColor"
   End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
   BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
   If mBackColor <> NewColor Then
      mBackColor = NewColor
      Call SetPaint
      PropertyChanged "BackColor"
   End If
End Property

Public Property Get BackStyle() As waBackStyle
   BackStyle = mBackStyle
End Property

Public Property Let BackStyle(ByVal NewStyle As waBackStyle)
   ' Make sure the style is 0=transparent or 1=opaque
   If NewStyle < 0 Or NewStyle > 1 Then NewStyle = 1
   
   If mBackStyle <> NewStyle Then
      mBackStyle = NewStyle
      Call SetPaint
      PropertyChanged "BackStyle"
   End If
End Property

Public Property Get FillColor() As OLE_COLOR
   FillColor = mFillColor
End Property

Public Property Let FillColor(ByVal NewColor As OLE_COLOR)
   If mFillColor <> NewColor Then
      mFillColor = NewColor
      If mFillStyle = waFillSolid Then
         Call SetPaint
      End If
      PropertyChanged "FillColor"
   End If
End Property

Public Property Get FillStyle() As waFillStyle
   FillStyle = mFillStyle
End Property

Public Property Let FillStyle(ByVal NewStyle As waFillStyle)
   If mFillStyle <> NewStyle Then
      mFillStyle = NewStyle
      Call SetPaint
      PropertyChanged "FillStyle"
   End If
End Property

Public Property Get GradientStart() As OLE_COLOR
   GradientStart = mGradientStart
End Property

Public Property Let GradientStart(ByVal NewColor As OLE_COLOR)
   If mGradientStart <> NewColor Then
      mGradientStart = NewColor
      If mFillStyle = waFillGradient Then
         Call SetPaint
      End If
      PropertyChanged "GradientStart"
   End If
End Property

Public Property Get GradientEnd() As OLE_COLOR
   GradientEnd = mGradientEnd
End Property

Public Property Let GradientEnd(ByVal NewColor As OLE_COLOR)
   If mGradientEnd <> NewColor Then
      mGradientEnd = NewColor
      If mFillStyle = waFillGradient Then
         Call SetPaint
      End If
      PropertyChanged "GradientEnd"
   End If
End Property

Public Property Get FillPicture() As IPictureDisp
   Set FillPicture = mFillPicture
End Property

Public Property Set FillPicture(NewPicture As IPictureDisp)
   Set mFillPicture = NewPicture
   If mFillStyle = waFillUserDefined Then
      Call SetPaint
   End If
   PropertyChanged "FillPicture"
End Property

Public Property Get LineWidth() As Long
   LineWidth = mLineWidth
End Property

Public Property Let LineWidth(ByVal NewWidth As Long)
   If NewWidth < 1 Then NewWidth = 1
   
   If mLineWidth <> NewWidth Then
      mLineWidth = NewWidth
      Call SetPaint
      PropertyChanged "LineWidth"
   End If
End Property

Public Property Get FontOffset() As Single
   FontOffset = mFontOffset
End Property

Public Property Let FontOffset(ByVal NewOffset As Single)
   Dim d As Double
   
   d = UserControl.ScaleX(NewOffset, vbPixels, vbTwips)
   
   If mFontOffset <> d Then
      mFontOffset = d
      mChangeOffset = True
      Call SetPaint
      PropertyChanged "FontOffset"
   End If
End Property

Public Property Get OffsetAngle() As Single
   OffsetAngle = mOffsetAngle
End Property

Public Property Let OffsetAngle(ByVal NewAngle As Single)
   ' Limit this to some halfway reasonable value.
   If OffsetAngle > 3600 Or OffsetAngle < -3600 Then OffsetAngle = 0
   
   If mOffsetAngle <> NewAngle Then
      mOffsetAngle = NewAngle
      mChangeOffset = True
      Call SetPaint
      PropertyChanged "OffsetAngle"
   End If
End Property

Public Property Get OffsetColor() As OLE_COLOR
   OffsetColor = mOffsetColor
End Property

Public Property Let OffsetColor(ByVal NewColor As OLE_COLOR)
   If mOffsetColor <> NewColor Then
      mOffsetColor = NewColor
      Call SetPaint
      PropertyChanged "OffsetColor"
   End If
End Property

Public Property Get Rotation() As Single
   Rotation = mRotation
End Property

Public Property Let Rotation(ByVal NewRotation As Single)
   ' Make sure the rotation is valid
   If NewRotation < 0 Then NewRotation = 0
   If NewRotation > 360 Then NewRotation = 0
   
   If mRotation <> NewRotation Then
      mRotation = NewRotation
      mChangeTransFont = True
      Call SetPaint
      PropertyChanged "Rotation"
   End If
End Property

Public Property Get BottomAlignment() As waAlignment
   BottomAlignment = mBottomAlignment
End Property

Public Property Let BottomAlignment(ByVal NewAlignment As waAlignment)
   If mBottomAlignment <> NewAlignment Then
      mBottomAlignment = NewAlignment
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomAlignment"
   End If
End Property

Public Property Get TopAlignment() As waAlignment
   TopAlignment = mTopAlignment
End Property

Public Property Let TopAlignment(ByVal NewAlignment As waAlignment)
   If mTopAlignment <> NewAlignment Then
      mTopAlignment = NewAlignment
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopAlignment"
   End If
End Property

'*****************************************
'** Methods used to set and get the
'** transformation coordinates.
'*****************************************

Public Sub SetBottomAlignment(ByVal X1 As Single, ByVal Y1 As Single, _
                              Optional ByVal x2 As Single = 1, Optional ByVal y2 As Single = 0, _
                              Optional ByVal x3 As Single = 1, Optional ByVal y3 As Single = 0)
                        
   If (X1 < 0) Or (X1 > 1) Then X1 = 1
   If (x2 < 0) Or (x2 > 1) Then x2 = 1
   If (x3 < 0) Or (x3 > 1) Then x3 = 1
   If (Y1 < -1) Or (Y1 > 1) Then Y1 = 0
   If (y2 < -1) Or (y2 > 1) Then y2 = 0
   If (y3 < -1) Or (y3 > 1) Then y3 = 0
   
   If (mBtmX1 <> X1) Or (mBtmY1 <> Y1) Or _
      (mBtmX2 <> x2) Or (mBtmY2 <> y2) Or _
      (mBtmX3 <> x3) Or (mBtmY3 <> y3) Then
      
      mBtmX1 = X1
      mBtmX2 = x2
      mBtmX3 = x3
      mBtmY1 = Y1
      mBtmY2 = y2
      mBtmY3 = y3
      mChangeTransformation = True
      Call SetPaint
   End If
End Sub

Public Sub GetBottomAlignment(X1 As Single, Y1 As Single, _
                              Optional x2 As Single, Optional y2 As Single, _
                              Optional x3 As Single, Optional y3 As Single)

   X1 = mBtmX1
   x2 = mBtmX2
   x3 = mBtmX3
   Y1 = mBtmY1
   y2 = mBtmY2
   y3 = mBtmY3
End Sub
                         
Public Sub SetTopAlignment(ByVal X1 As Single, ByVal Y1 As Single, _
                           Optional ByVal x2 As Single = 1, Optional ByVal y2 As Single = 0, _
                           Optional ByVal x3 As Single = 1, Optional ByVal y3 As Single = 0)
                        
   If (X1 < 0) Or (X1 > 1) Then X1 = 1
   If (x2 < 0) Or (x2 > 1) Then x2 = 1
   If (x3 < 0) Or (x3 > 1) Then x3 = 1
   If (Y1 < -1) Or (Y1 > 1) Then Y1 = 0
   If (y2 < -1) Or (y2 > 1) Then y2 = 0
   If (y3 < -1) Or (y3 > 1) Then y3 = 0
   
   If (mTopX1 <> X1) Or (mTopY1 <> Y1) Or _
      (mTopX2 <> x2) Or (mTopY2 <> y2) Or _
      (mTopX3 <> x3) Or (mTopY3 <> y3) Then
      
      mTopX1 = X1
      mTopX2 = x2
      mTopX3 = x3
      mTopY1 = Y1
      mTopY2 = y2
      mTopY3 = y3
      mChangeTransformation = True
      Call SetPaint
   End If
End Sub

Public Sub GetTopAlignment(X1 As Single, Y1 As Single, _
                           Optional x2 As Single, Optional y2 As Single, _
                           Optional x3 As Single, Optional y3 As Single)

   X1 = mTopX1
   x2 = mTopX2
   x3 = mTopX3
   Y1 = mTopY1
   y2 = mTopY2
   y3 = mTopY3
End Sub

Public Property Get TopX1() As Single
   TopX1 = mTopX1
End Property

Public Property Let TopX1(ByVal x As Single)
   If x > 1 Or x < 0 Then x = 1
   If mTopX1 <> x Then
      mTopX1 = x
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopX1"
   End If
End Property

Public Property Get TopX2() As Single
   TopX2 = mTopX2
End Property

Public Property Let TopX2(ByVal x As Single)
   If x > 1 Or x < 0 Then x = 1
   If mTopX2 <> x Then
      mTopX1 = x
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopX2"
   End If
End Property

Public Property Get TopX3() As Single
   TopX3 = mTopX3
End Property

Public Property Let TopX3(ByVal x As Single)
   If x > 1 Or x < 0 Then x = 1
   If mTopX3 <> x Then
      mTopX3 = x
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopX3"
   End If
End Property

Public Property Get TopY1() As Single
   TopY1 = mTopY1
End Property

Public Property Let TopY1(ByVal y As Single)
   If y > 1 Or y < -1 Then y = 0
   If mTopY1 <> y Then
      mTopY1 = y
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopY1"
   End If
End Property

Public Property Get TopY2() As Single
   TopY2 = mTopY2
End Property

Public Property Let TopY2(ByVal y As Single)
   If y > 1 Or y < -1 Then y = 0
   If mTopY2 <> y Then
      mTopY2 = y
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopY2"
   End If
End Property

Public Property Get TopY3() As Single
   TopY3 = mTopY3
End Property

Public Property Let TopY3(ByVal y As Single)
   If y > 1 Or y < -1 Then y = 0
   If mTopY3 <> y Then
      mTopY3 = y
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "TopY3"
   End If
End Property

Public Property Get BottomX1() As Single
   BottomX1 = mBtmX1
End Property

Public Property Let BottomX1(ByVal x As Single)
   If x > 1 Or x < 0 Then x = 1
   If mBtmX1 <> x Then
      mBtmX1 = x
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomX1"
   End If
End Property

Public Property Get BottomX2() As Single
   BottomX2 = mBtmX2
End Property

Public Property Let BottomX2(ByVal x As Single)
   If x > 1 Or x < 0 Then x = 1
   If mBtmX2 <> x Then
      mBtmX2 = x
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomX2"
   End If
End Property

Public Property Get BottomX3() As Single
   BottomX3 = mBtmX3
End Property

Public Property Let BottomX3(ByVal x As Single)
   If x > 1 Or x < 0 Then x = 1
   If mBtmX3 <> x Then
      mBtmX3 = x
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomX3"
   End If
End Property

Public Property Get BottomY1() As Single
   BottomY1 = mBtmY1
End Property

Public Property Let BottomY1(ByVal y As Single)
   If y > 1 Or y < -1 Then y = 0
   If mBtmY1 <> y Then
      mBtmY1 = y
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomY1"
   End If
End Property

Public Property Get BottomY2() As Single
   BottomY2 = mBtmY2
End Property

Public Property Let BottomY2(ByVal y As Single)
   If y > 1 Or y < -1 Then y = 0
   If mBtmY2 <> y Then
      mBtmY2 = y
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomY2"
   End If
End Property

Public Property Get BottomY3() As Single
   BottomY3 = mBtmY3
End Property

Public Property Let BottomY3(ByVal y As Single)
   If y > 1 Or y < -1 Then y = 0
   If mBtmY3 <> y Then
      mBtmY3 = y
      mChangeTransformation = True
      Call SetPaint
      PropertyChanged "BottomY3"
   End If
End Property

'
' ************************************************
' **   Event passthought functions.
' **   These don't do anything other than pass the
' **   user events though
' ************************************************

Private Sub UserControl_Click()
   RaiseEvent Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = KEY_PAINT Then
      Call UserControl_Paint
   Else
      RaiseEvent KeyDown(KeyCode, Shift)
   End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
   RaiseEvent DblClick
End Sub



'*********************************
'**  The resize event triggers
'**  a redraw if it was not caused
'**  from withing the Paint event.
'*********************************

Private Sub UserControl_Resize()
   mChangeFinal = True
   
   If Not mInPaint Then
      Call SetPaint
   End If
End Sub


'*************************************
'**  The Paint event and the following
'**  CalculateX functions are the
'**  workhorse of this control.
'*************************************

Private Sub UserControl_Paint()
   Dim d As Double
   Dim mm As Long
   Dim hBrush As Long
   Dim hPen As Long
   Dim lBrush As LOGBRUSH
   Dim hBitmap As Long
   Dim hTempDC As Long
   Dim i As Long
   Dim x As Long
   Dim y As Long
   Dim w As Long
   Dim h As Long
   Dim w2 As Long
   Dim h2 As Long
   
   Dim bmWidth As Long
   Dim bmHeight As Long
   Dim bmOffsetX As Long
   Dim bmOffsetY As Long
   
   Dim OffsetShadowX As Double
   Dim OffsetShadowY As Double
   
   
   On Error GoTo ErrorHandle:
   
   If Not mInPaint Then
      mInPaint = True
      
      If Len(mCaption) = 0 Then
         ' Draw the background only.
         UserControl.Cls
         
         ' Draw the control
         mm = SetMapMode(UserControl.hdc, MM_TWIPS)
   
         
         ' Draw the background
         hPen = CreatePen(PS_SOLID, 1, mBackColor)
         hBrush = CreateSolidBrush(mBackColor)
         hPen = SelectObject(UserControl.hdc, hPen)
         hBrush = SelectObject(UserControl.hdc, hBrush)
         
         Call Rectangle(UserControl.hdc, 0, 0, UserControl.width, -UserControl.height)
         
         hPen = SelectObject(UserControl.hdc, hPen)
         hBrush = SelectObject(UserControl.hdc, hBrush)
         Call DeleteObject(hPen)
         hPen = 0
         Call DeleteObject(hBrush)
         hBrush = 0
         
         ' Clean up
         Call SetMapMode(UserControl.hdc, mm)
         mm = 0
         
         
         ' Handle the transparent background.
         If mBackStyle = 0 Then
            UserControl.BackStyle = vbTransparent
            UserControl.MaskPicture = UserControl.Image
            UserControl.MaskColor = mBackColor
         Else
            Set UserControl.MaskPicture = Nothing
            UserControl.BackStyle = 1
         End If
   
      Else
         
         ' Calculate only the portions required
         If mChangeFont Then
            Call CalculateFont
            mChangeTransformation = True
            mChangeTransFont = True
         End If
         
         If mChangeTransformation Then
            Call CalculateTransformation
            mChangeTransFont = True
         End If
         
         If mChangeTransFont Then
            Call CalculateTransFont
         End If
         
         ' Calculate the adjustment for the shadow
         If mFontOffset = 0 Then
            OffsetShadowX = 0
            OffsetShadowY = 0
         Else
            OffsetShadowX = Cos(DEG_TO_RAD * (mOffsetAngle + mRotation)) * mFontOffset
            OffsetShadowY = Sin(DEG_TO_RAD * (mOffsetAngle + mRotation)) * mFontOffset
         End If
         
         ' Check the size of the box
         If mAutosize Then
            UserControl.width = mFontMaxX - mFontMinX + Abs(OffsetShadowX) + UserControl.ScaleX(2, vbPixels, vbTwips)
            UserControl.height = mFontMaxY - mFontMinY + Abs(OffsetShadowY) + UserControl.ScaleY(2, vbPixels, vbTwips)
         End If
         
         
         ' Calculate the shift necessary to fit the font
         ' in the box.
         d = -mFontMinX + (UserControl.width - (mFontMaxX - mFontMinX + OffsetShadowX)) / 2
         If d <> mFFontShiftX Then
            mFFontShiftX = d
            mChangeFinal = True
         End If
         
         bmOffsetX = mFontMinX + d
         
         d = -mFontMinY + (UserControl.height - (mFontMaxY - mFontMinY + OffsetShadowY)) / 2
         If d <> mFFontShiftY Then
           mFFontShiftY = d
           mChangeFinal = True
         End If
         
         bmOffsetY = mFontMinY + d
         bmWidth = mFontMaxX - mFontMinX
         bmHeight = mFontMaxY - mFontMinY
         
         
         
         ' Calculate the API polyline
         If mChangeFinal Then
            Call CalculateAPIPoly
            If mFontOffset <> 0 Then
               mChangeOffset = True
            End If
         End If
         
         ' Calculate the offset/shadow text.
         If mChangeOffset And (mFontOffset <> 0) Then
            Call CalculateAPIShadow(CLng(OffsetShadowX), CLng(OffsetShadowY))
         End If
         
         
         UserControl.Cls
         
         ' Draw the control
         mm = SetMapMode(UserControl.hdc, MM_TWIPS)
         
              
         ' Draw the background
         hPen = CreatePen(PS_SOLID, 1, mBackColor)
         hBrush = CreateSolidBrush(mBackColor)
         hPen = SelectObject(UserControl.hdc, hPen)
         hBrush = SelectObject(UserControl.hdc, hBrush)
         
         Call Rectangle(UserControl.hdc, 0, 0, UserControl.width, -UserControl.height)
         
         hPen = SelectObject(UserControl.hdc, hPen)
         hBrush = SelectObject(UserControl.hdc, hBrush)
         Call DeleteObject(hPen)
         hPen = 0
         Call DeleteObject(hBrush)
         hBrush = 0
         
         ' Draw the shadow/Offset
         If mFontOffset <> 0 Then
            hBrush = CreateSolidBrush(mOffsetColor)
            hPen = CreatePen(PS_SOLID, 1, mOffsetColor)
            hBrush = SelectObject(UserControl.hdc, hBrush)
            hPen = SelectObject(UserControl.hdc, hPen)
            
            Call PolyPolygon(UserControl.hdc, mSFontPoly(0), mFontPoly(0), UBound(mFontPoly) - LBound(mFontPoly) + 1)
            
            hBrush = SelectObject(UserControl.hdc, hBrush)
            hPen = SelectObject(UserControl.hdc, hPen)
            Call DeleteObject(hBrush)
            hBrush = 0
            Call DeleteObject(hPen)
            hPen = 0
         End If
         
         ' Draw the interior if it is not a solid fill.
         ' (bitmap pens won't work with all systems.)
         If ((mFillStyle = waFillUserDefined) And (Not mFillPicture Is Nothing)) Or (mFillStyle = waFillGradient) Then
            ' Mask out the area
            hBrush = CreateSolidBrush(vbBlack)
            hPen = CreatePen(PS_SOLID, 1, vbBlack)
            hBrush = SelectObject(UserControl.hdc, hBrush)
            hPen = SelectObject(UserControl.hdc, hPen)
            
            Call PolyPolygon(UserControl.hdc, mFFontPoly(0), mFontPoly(0), UBound(mFontPoly) - LBound(mFontPoly) + 1)
            
            hBrush = SelectObject(UserControl.hdc, hBrush)
            hPen = SelectObject(UserControl.hdc, hPen)
            
            ' Create a working bitmap
            hTempDC = CreateCompatibleDC(UserControl.hdc)
            hBitmap = CreateCompatibleBitmap(UserControl.hdc, _
               UserControl.ScaleX(UserControl.width, vbTwips, vbPixels), _
               UserControl.ScaleY(UserControl.height, vbTwips, vbPixels))
            Call SetMapMode(hTempDC, MM_TWIPS)
            
            ' Draw the background
            hBitmap = SelectObject(hTempDC, hBitmap)
            hBrush = SelectObject(hTempDC, hBrush)
            hPen = SelectObject(hTempDC, hPen)
            Call Rectangle(hTempDC, 0, 0, UserControl.width, -UserControl.height)
            hBrush = SelectObject(hTempDC, hBrush)
            hPen = SelectObject(hTempDC, hPen)
            Call DeleteObject(hBrush)
            hBrush = 0
            Call DeleteObject(hPen)
            hPen = 0
            
            ' Draw the font mask
            hBrush = CreateSolidBrush(vbWhite)
            hPen = CreatePen(PS_SOLID, 1, vbWhite)
            hBrush = SelectObject(hTempDC, hBrush)
            hPen = SelectObject(hTempDC, hPen)
            
            Call PolyPolygon(hTempDC, mFFontPoly(0), mFontPoly(0), UBound(mFontPoly) - LBound(mFontPoly) + 1)
            
            hBrush = SelectObject(hTempDC, hBrush)
            hPen = SelectObject(hTempDC, hPen)
            Call DeleteObject(hBrush)
            hBrush = 0
            Call DeleteObject(hPen)
            hPen = 0
            
            ' paint the font
            Select Case mFillStyle
               Case waFillUserDefined:
                  picTemp.AutoSize = True
                  picTemp.Picture = mFillPicture
                  'Set picTemp.Image = picTemp.Picture
                  h = picTemp.ScaleY(mFillPicture.height, vbHimetric, vbPixels)
                  w = picTemp.ScaleX(mFillPicture.width, vbHimetric, vbPixels)
                  h2 = picTemp.ScaleX(UserControl.height, vbTwips, vbPixels)
                  w2 = picTemp.ScaleY(UserControl.width, vbTwips, vbPixels)
                  
                  ' Not sure why this doesn't work in MM_TWIPS
                  i = SetMapMode(hTempDC, MM_TEXT)
                  
                  For x = 0 To w2 Step w
                     For y = 0 To h2 Step h
                        Call BitBlt(hTempDC, x, y, w, h, _
                           picTemp.hdc, 0, 0, SRCAND)
                     Next y
                  Next x
                  
                  Call SetMapMode(hTempDC, i)
               
               Case waFillGradient:
                  ' Create the picture
                  Set picTemp.Picture = Nothing
                  picTemp.AutoSize = False
                  picTemp.width = bmWidth
                  picTemp.height = bmHeight
                  
                  Call GradientFill(picTemp.hdc, 0, 0, bmWidth, bmHeight, _
                     mGradientStart, mGradientEnd)
                  
                  ' Create the masked text.
                  i = SetMapMode(hTempDC, MM_TEXT)
                  Call BitBlt(hTempDC, _
                     UserControl.ScaleX(bmOffsetX, vbTwips, vbPixels), _
                     UserControl.ScaleY(bmOffsetY, vbTwips, vbPixels), _
                     UserControl.ScaleX(bmWidth, vbTwips, vbPixels), _
                     UserControl.ScaleY(bmHeight, vbTwips, vbPixels), picTemp.hdc, 0, 0, SRCAND)
                  Call SetMapMode(hTempDC, i)
                  
            End Select
            
            ' BitBlt the font to the control  (OR operation)
            Call BitBlt(UserControl.hdc, 0, 0, UserControl.width, -UserControl.height, _
               hTempDC, 0, 0, SRCPAINT)
            
            ' Cleanup
            hBitmap = SelectObject(hTempDC, hBitmap)
            Call DeleteObject(hBitmap)
            hBitmap = 0
            Call ReleaseDC(0, hTempDC)
            hTempDC = 0
         End If
         
         
         ' Draw the font
         hPen = CreatePen(PS_SOLID, mLineWidth, mFontColor)
         
         ' Create the brush for the given FillStyle
         Select Case mFillStyle
            Case waFillNone, waFillUserDefined, waFillGradient:
               lBrush.lbStyle = BS_HOLLOW
               lBrush.lbColor = 0
               lBrush.lbHatch = 0
               hBrush = CreateBrushIndirect(lBrush)
               
            Case waFillSolid:
               hBrush = CreateSolidBrush(mFillColor)
               
            Case waFillTransparent:
               hBrush = CreateSolidBrush(mBackColor)
            
         End Select
         
         
         hPen = SelectObject(UserControl.hdc, hPen)
         hBrush = SelectObject(UserControl.hdc, hBrush)
         
         Call PolyPolygon(UserControl.hdc, mFFontPoly(0), mFontPoly(0), UBound(mFontPoly) - LBound(mFontPoly) + 1)
         
         hBrush = SelectObject(UserControl.hdc, hBrush)
         Call DeleteObject(hBrush)
         hBrush = 0
         hPen = SelectObject(UserControl.hdc, hPen)
         Call DeleteObject(hPen)
         hPen = 0
         
         ' Clean up
         Call SetMapMode(UserControl.hdc, mm)
         mm = 0
         
         ' Handle the transparent background.
         If mBackStyle = 0 Then
            UserControl.BackStyle = vbTransparent
            UserControl.MaskPicture = UserControl.Image
            UserControl.MaskColor = mBackColor
         Else
            Set UserControl.MaskPicture = Nothing
            UserControl.BackStyle = 1
         End If
      End If
      
      ' Reset all the flags.
      mChangeOffset = False
      mChangeFinal = False
      mChangeFont = False
      mChangeTransformation = False
      mChangeTransFont = False
      mPaintCalled = False
      mInPaint = False
   End If
   
ErrorHandle:
   ' Cleanup objects the best we can if failure
   If hPen <> 0 Then Call DeleteObject(hPen)
   If hBrush <> 0 Then Call DeleteObject(hPen)
   If mm <> 0 Then Call SetMapMode(UserControl.hdc, mm)
   If hBitmap <> 0 Then Call DeleteObject(hBitmap)
   If hTempDC <> 0 Then Call ReleaseDC(0, hTempDC)
   
   ' We can't be sure any of these have been
   ' calculated properly.
   mChangeOffset = True
   mChangeFinal = True
   mChangeFont = True
   mChangeTransformation = True
   mChangeTransFont = True
   
   mPaintCalled = False
   mInPaint = False
End Sub

' This calculates the polylines that outline
' the font for the normal font size.
Private Sub CalculateFont()
   Dim i As Long
   Dim pCount As Long
   Dim polyCount As Long
   Dim buffer() As Long
   Dim GlyphData() As Variant
   Dim w As Long
   Dim FLog As LOGFONT
   Dim hFont As Long
   Dim fSize As APISize
   Dim j As Long
   Dim jLast As Long
   Dim WidthScale As Double
   

   'Create the font.
   FLog.lfHeight = -mFontSize * 20#
   FLog.lfWidth = 0
   FLog.lfEscapement = 0
   FLog.lfOrientation = 0
   FLog.lfWeight = mFontWeight
   If mFontItalic Then
      FLog.lfItalic = 255
   Else
      FLog.lfItalic = 0
   End If
   FLog.lfUnderline = 0
   FLog.lfStrikeOut = 0
   FLog.lfCharSet = 0
   FLog.lfOutPrecision = 0
   FLog.lfClipPrecision = 0
   FLog.lfQuality = 0
   FLog.lfPitchAndFamily = 0
   FLog.lfFaceName = mFontName & vbNullChar
   
   hFont = CreateFontIndirect(FLog)
   hFont = SelectObject(hdc, hFont)
   
   '' Get the emUnit of the font
   'FLog.lfHeight = -GetEMUnit(hdc)
   '
   'hFont = SelectObject(hdc, hFont)
   'Call DeleteObject(hFont)
   
   '' Create the font of the correct size.
   'hFont = CreateFontIndirect(FLog)
   'hFont = SelectObject(hdc, hFont)
   
   
   ReDim GlyphData(1 To Len(mCaption))
   
   ' Precalculate the space required.
   For i = 1 To Len(mCaption)
       Call GetOutlineCount(hdc, Asc(Mid$(mCaption, i, 1)), buffer(), pCount, polyCount)
       ' Save the buffer for later use
       GlyphData(i) = buffer
   Next i
   
   
   ' Allocate the space for the font data
   ReDim mFontX(0 To pCount - 1)
   ReDim mFontY(0 To pCount - 1)
   ReDim mFontPoly(0 To polyCount - 1)
   pCount = 0
   polyCount = 0
   w = 0
   jLast = 0
   
   ' Get the glyph outlines
   For i = 1 To Len(mCaption)
      buffer = GlyphData(i)
      Call GetBufferOutline(buffer, mFontX, mFontY, pCount, mFontPoly, polyCount)
      
      ' Shift each character to it's appropriate position
      For j = jLast To pCount - 1
         mFontX(j) = mFontX(j) + w
         mFontY(j) = -mFontY(j)
      Next j
      jLast = pCount
      
      ' Get this character's width
      Call GetTextExtentPoint32(hdc, Mid$(mCaption, i, 1), 1, fSize)
      w = w + fSize.cx
   Next i
   
   ' Store the normalized font height and width.
   Call GetTextExtentPoint32(hdc, "X", 1, fSize)
   mFontCHeight = fSize.cy
   mFontCWidth = w
   
   
   'Release the font
   hFont = SelectObject(hdc, hFont)
   Call DeleteObject(hFont)
   
   ' Adjust the width
   ' if required.
   ScaleWidth = mFontWidth / mFontSize
   
   If ScaleWidth <> 1 Then
      For i = 0 To pCount - 1
         mFontX(i) = ScaleWidth * mFontX(i)
      Next i
   End If
End Sub

' This calculates the transformation curves/functions
' that will be used to warp the font to the proposed
' alignment.
Private Sub CalculateTransformation()
   Dim al As Long
   Dim OffsetY As Double
   Dim i As Long
   Dim bCopyBottom As Boolean
   Dim bCopyTop As Boolean
   
   
   bCopyBottom = False
   bCopyTop = False
   
   ' Make sure the top alignment is valid
   If mTopAlignment = waAlignmentParallel Then
      If mBottomAlignment = waAlignmentParallel Then
         al = waAlignmentNone
      Else
         al = mBottomAlignment
         bCopyBottom = True
      End If
   Else
      al = mTopAlignment
   End If
   
   mTransWidth = mFontCWidth / CDbl(TRANSFORM_COUNT + 1)
   ReDim mTransTop(0 To TRANSFORM_COUNT)
   OffsetY = mFontCHeight '+ mFontMinY
   
   If Not bCopyBottom Then
      Call CalculateAlignment(al, mTransTop, OffsetY, _
         mTopX1, mTopY1, mTopX2, mTopX2, mTopX3, mTopY3)
   End If
   
   
   ' Make sure the bottom alignment is valid
   If mBottomAlignment = waAlignmentParallel Then
      If mTopAlignment = waAlignmentParallel Then
         al = waAlignmentNone
      Else
         al = mTopAlignment
         bCopyTop = True
      End If
   Else
      al = mBottomAlignment
   End If
   
   ReDim mTransBottom(0 To TRANSFORM_COUNT)
   OffsetY = 0
   
   If Not bCopyTop Then
      Call CalculateAlignment(al, mTransBottom, OffsetY, _
         mBtmX1, mBtmY1, mBtmX2, mBtmY2, mBtmX3, mBtmY3)
   End If
   
   If bCopyTop Then
      For i = 0 To TRANSFORM_COUNT
         mTransBottom(i) = mTransTop(i)
      Next i
   ElseIf bCopyBottom Then
      For i = 0 To TRANSFORM_COUNT
         mTransTop(i) = mTransBottom(i)
      Next i
   End If
   
      
   ' Move the top alignment to adjust for the font height
   OffsetY = mTransBottom(0) - mTransTop(0) + mFontSize * 20#
   For i = 0 To TRANSFORM_COUNT
      mTransTop(i) = mTransTop(i) + OffsetY
   Next i
End Sub

' Calculates an alignment.
Private Sub CalculateAlignment(ByVal alType As Long, alArray() As Double, ByVal alOffsetY As Double, _
   ByVal X1 As Double, ByVal Y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
   ByVal x3 As Double, ByVal y3 As Double)
   
   Dim i As Long
   Dim dt As Double
   Dim x() As Double
   Dim y() As Double
   Dim c() As Double
   Dim cf As CurveFit
   Dim dx As Double
   Dim dmin As Double
   Dim dmax As Double
   
   ' Calculate the alignment transformation array
   Select Case alType
      Case waAlignmentNone:
         ' Alignment is a linear line (no transformation really)
         For i = 0 To TRANSFORM_COUNT
            alArray(i) = 0
         Next i
         
      Case waAlignmentLinear:
         ' Alignment is a line from 0,0 to 1, Y1
         dt = (Y1 * mFontCWidth) / CDbl(TRANSFORM_COUNT)
         alArray(i) = 0
         
         For i = 1 To TRANSFORM_COUNT
            alArray(i) = alArray(i - 1) + dt
         Next i
      
      Case waAlignmentArc:
         Set cf = New CurveFit
         
         ReDim x(0 To 2)
         ReDim y(0 To 2)
         
         x(0) = 0#
         y(0) = 0#
         x(1) = X1
         y(1) = Y1
         x(2) = x2
         y(2) = 1#
         
         Call cf.PolynomialCurveFit(x, y, 2, c)
         
         dt = 1# / TRANSFORM_COUNT
         dx = 0
         
         For i = 0 To TRANSFORM_COUNT
            alArray(i) = c(1) + dx * c(2) + dx * dx * c(3)
            alArray(i) = alArray(i) * mFontCWidth
            dx = dx + dt
         Next i
          
         dmin = alArray(0)
         dmax = dmin
         
         For i = 1 To TRANSFORM_COUNT
            If alArray(i) < dmin Then
               dmin = alArray(i)
            ElseIf alArray(i) > dmax Then
               dmax = alArray(i)
            End If
         Next i
         
         
         For i = 0 To TRANSFORM_COUNT
            alArray(i) = alArray(i) * mFontCHeight / (dmax - dmin)
         Next i
         
         
      Case waAlignmentFit3, waAlignmentFit4, waAlignmentFit5:
         Set cf = New CurveFit
         
         ReDim x(0 To 3)
         ReDim y(0 To 3)
         
         x(0) = 0#
         y(0) = 0#
         x(1) = X1
         y(1) = Y1
         x(2) = x2
         y(2) = y2
         x(3) = x3
         y(3) = y3
      
         dt = 1# / TRANSFORM_COUNT
         dx = 0
         
         Select Case alType
            Case waAlignmentFit3:
               Call cf.PolynomialCurveFit(x, y, 3, c)
               For i = 0 To TRANSFORM_COUNT
                  alArray(i) = c(1) + dx * c(2) + dx * dx * c(3) + (dx ^ 3) * c(4)
                  dx = dx + dt
               Next i
         
            Case waAlignmentFit4:
               Call cf.PolynomialCurveFit(x, y, 4, c)
               For i = 0 To TRANSFORM_COUNT
                  alArray(i) = c(1) + dx * c(2) + dx * dx * c(3) + (dx ^ 3) * c(4) + (dx ^ 4) * c(5)
                  dx = dx + dt
               Next i
               
            Case waAlignmentFit5:
               Call cf.PolynomialCurveFit(x, y, 5, c)
               For i = 0 To TRANSFORM_COUNT
                  alArray(i) = c(1) + dx * c(2) + dx * dx * c(3) + (dx ^ 3) * c(4) + (dx ^ 4) * c(5) + (dx ^ 5) * c(6)
                  dx = dx + dt
               Next i
         End Select
         
         dmin = alArray(0)
         dmax = dmin
         
         For i = 1 To TRANSFORM_COUNT
            If alArray(i) < dmin Then
               dmin = alArray(i)
            ElseIf alArray(i) > dmax Then
               dmax = alArray(i)
            End If
         Next i
         
         
         For i = 0 To TRANSFORM_COUNT
            alArray(i) = alArray(i) * mFontCHeight / (dmax - dmin)
         Next i
         
      Case waAlignmentQuadratic:
         ' Alignment is a bezier curve though points:
         '   0,0 - X1, Y1 - 1, Y2
         ReDim x(0 To TRANSFORM_COUNT)
         Call QuadraticBezier(x, alArray, 0, TRANSFORM_COUNT + 1, 0, 0, _
            X1 * mFontCWidth, Y1 * mFontCWidth, mFontCWidth, y2 * mFontCWidth)
      
         Call NormalizeCurve(x, alArray, mTransWidth)
         
      Case waAlignmentCubic:
         ' Alignment is a bezier curve though points:
         '   0,0 - X1,Y1 - X2,Y2 - 1,Y3
         ReDim x(0 To TRANSFORM_COUNT)
         Call CubicBezier(x, alArray, 0, TRANSFORM_COUNT + 1, 0, 0, _
            X1 * mFontCWidth, Y1 * mFontCWidth, x2 * mFontCWidth, y2 * mFontCWidth, _
            mFontCWidth, y3 * mFontCWidth)
         
         Call NormalizeCurve(x, alArray, mTransWidth)
      
      
   End Select
   
   ' Adjust the offset of the alignment if necessary
   ' This is not needed.
   'If OffsetY <> 0 Then
   '   For i = 0 To TRANSFORM_COUNT - 1
   '      alArray(i) = alArray(i) + OffsetY
   '   Next i
   'End If
End Sub

' Takes the curve defined by the polyline x(),y() and normalizes it.
' That is, calculate the points along the curve at even intervals of x.
Public Sub NormalizeCurve(x() As Double, y() As Double, xStep As Double)
   Dim yout() As Double
   Dim dx As Double
   Dim i As Long
   Dim j As Long
   Dim lastX As Double
   Dim lastY As Double
   Dim currX As Double
   Dim currY As Double
   
   lastX = x(0)
   lastY = y(0)
   currX = x(1)
   currY = y(1)
   dx = 0
   j = 1
   ReDim yout(0 To TRANSFORM_COUNT)
   
   For i = 1 To TRANSFORM_COUNT - 1
      dx = dx + xStep
      
      ' (lastx,lastY)-(currX,currY) defines the line segment
      ' point dx falls on.
      Do While Not ((dx > lastX) And (dx <= currX))
         Do While x(j) < dx
            j = j + 1
         Loop
         
         lastX = currX
         lastY = currY
         currX = x(j)
         currY = y(j)
      Loop
      
      ' prorate the current position
      yout(i) = ((dx - lastX) / (currX - lastX)) * (currY - lastY) + lastY
   Next i
   
   yout(TRANSFORM_COUNT) = y(TRANSFORM_COUNT)
   
   ' Finalize the output
   x(0) = 0
   y(0) = yout(0)
   For i = 1 To TRANSFORM_COUNT
      x(i) = x(i - 1) + xStep
      y(i) = yout(i)
   Next i
End Sub

' This function transforms and rotates the font
' to it's final form.
Private Sub CalculateTransFont()
   Dim i As Long
   Dim ang As Double
   Dim sinAng As Double
   Dim cosAng As Double
   
   Dim x As Double
   Dim y As Double
   Dim dHeight As Double
   Dim dx As Long
   Dim dxMax As Long
   Dim maxY As Double
   Dim minY As Double
   
   ' Allocate space
   ReDim mTransX(LBound(mFontX) To UBound(mFontX))
   ReDim mTransY(LBound(mFontY) To UBound(mFontY))
   
      
   ' Transform
   If (mTopAlignment <> waAlignmentNone) Or _
      (mBottomAlignment <> waAlignmentNone) Then
      
      ' Calculate the min and max Y values
      minY = mFontY(LBound(mFontY))
      maxY = minY
      
      For i = LBound(mFontX) + 1 To UBound(mTransX)
         If mFontY(i) > maxY Then
            maxY = mFontY(i)
         ElseIf mFontY(i) < minY Then
            minY = mFontY(i)
         End If
      Next i
      
      dHeight = maxY - minY
      dxMax = TRANSFORM_COUNT
      
      ' Transform the font
      For i = LBound(mFontX) To UBound(mFontX)
         mTransX(i) = mFontX(i)
         
         dx = CLng(mFontX(i) \ mTransWidth)
         'If dx > dxMax Then
         '   dx = dxMax
         'End If
         
         mTransY(i) = mTransBottom(dx) + _
            (mTransTop(dx) - mTransBottom(dx)) * ((mFontY(i) - minY) / dHeight)
      Next i
   Else
   
      ' No real transformation
      For i = LBound(mFontX) To UBound(mFontX)
         mTransX(i) = mFontX(i)
         mTransY(i) = mFontY(i)
      Next i
   End If
   
   
   
   ' Rotate
   If mRotation <> 0 Then
      ang = DEG_TO_RAD * (-mRotation)
      sinAng = Sin(ang)
      cosAng = Cos(ang)
      
      For i = LBound(mFontX) To UBound(mFontX)
         x = mTransX(i)
         y = mTransY(i)
         
         mTransX(i) = x * cosAng - y * sinAng
         mTransY(i) = x * sinAng + y * cosAng
      Next i
   End If
   
   ' Calculate the bounding box
   mFontMinX = mTransX(LBound(mTransX))
   mFontMaxX = mFontMinX
   mFontMinY = mTransY(LBound(mTransY))
   mFontMaxY = mFontMinY
    
   For i = LBound(mTransX) + 1 To UBound(mTransX)
      If mTransX(i) > mFontMaxX Then
         mFontMaxX = mTransX(i)
      ElseIf mTransX(i) < mFontMinX Then
         mFontMinX = mTransX(i)
      End If
      
      If mTransY(i) > mFontMaxY Then
         mFontMaxY = mTransY(i)
      ElseIf mTransY(i) < mFontMinY Then
         mFontMinY = mTransY(i)
      End If
   Next i
End Sub

' This function changes the TransFont form to the integer
' form required by the API (shifted to be centered in the control.)
Private Sub CalculateAPIPoly()
   Dim i As Long
   
   ReDim mFFontPoly(LBound(mTransX) To UBound(mTransX))
   
   
   For i = LBound(mTransX) To UBound(mTransX)
      mFFontPoly(i).x = CLng(mTransX(i) + mFFontShiftX)
      mFFontPoly(i).y = -CLng(mTransY(i) + mFFontShiftY)
   Next i
End Sub

Private Sub CalculateAPIShadow(ByVal OffsetX As Long, ByVal OffsetY As Long)
   Dim i As Long
   
   ReDim mSFontPoly(LBound(mFFontPoly) To UBound(mFFontPoly))
   
   For i = LBound(mFFontPoly) To UBound(mFFontPoly)
      mSFontPoly(i).x = mFFontPoly(i).x + OffsetX
      mSFontPoly(i).y = mFFontPoly(i).y + OffsetY
   Next i
End Sub


' Draws a gradient fill top to bottom.
Private Sub GradientFill(ByVal hdc As Long, ByVal Left As Single, ByVal Top As Single, ByVal width As Single, ByVal height As Single, ByVal StartColor As Long, ByVal endColor As Long)
   Dim currRed As Single
   Dim currGreen As Single
   Dim currBlue As Single
   Dim stepRed As Single
   Dim stepGreen As Single
   Dim stepBlue As Single
   Dim stepNumber As Long
   Dim stepY As Single
   Dim hBrush As Long
   Dim FillArea As RECT
   Dim i As Long
   Dim OldMode As Long
   Dim tpx As Single
   Dim tpy As Single
   Dim y As Single
   
   
   'Break the colors apart.
   currRed = CSng(StartColor And &HFF&)
   currGreen = CSng((StartColor And &HFF00&) / &H100&)
   currBlue = CSng((StartColor And &HFF0000) / &H10000)
   
   stepRed = CSng(endColor And &HFF&)
   stepGreen = CSng((endColor And &HFF00&) / &H100&)
   stepBlue = CSng((endColor And &HFF0000) / &H10000)
   
   ' Calculate the fewest number of steps to make a good gradient.
   stepNumber = Abs(currRed - stepRed)
   i = Abs(currGreen - stepGreen)
   If i > stepNumber Then stepNumber = i
   i = Abs(currBlue - stepBlue)
   If i > stepNumber Then stepNumber = i
   If stepNumber < 1 Then stepNumber = 1
   
   ' Calculate the step sizes
   stepRed = (currRed - stepRed) / CSng(stepNumber)
   stepGreen = (currGreen - stepGreen) / CSng(stepNumber)
   stepBlue = (currBlue - stepBlue) / CSng(stepNumber)
   stepY = height / CSng(stepNumber)
   
   ' Calculate twips per pixel
   tpx = GetDeviceCaps(hdc, LOGPIXELSX)
   If tpx <> 0 Then
      tpx = 1440# / tpx
   Else
      tpx = 12 '?
   End If
   
   tpy = GetDeviceCaps(hdc, LOGPIXELSY)
   If tpy <> 0 Then
      tpy = 1440# / tpy
   Else
      tpy = 12 '?
   End If
   
   y = Top / tpx
   FillArea.Left = Left / tpy
   FillArea.Right = (Left + width) / tpx
   stepY = stepY / tpy
   

   
   OldMode = SetMapMode(hdc, MM_TEXT)
   
   ' Draw the gradient
   For i = 1 To stepNumber
      FillArea.Top = y
      FillArea.Bottom = y + stepY
      
      'Create the brush
      hBrush = CreateSolidBrush(RGB(CInt(currRed), CInt(currGreen), CInt(currBlue)))
      
      Call FillRect(hdc, FillArea, hBrush)
      
      'Delete the brush
      Call DeleteObject(hBrush)
      
      'Pick the new color and new top coordinate
      currRed = currRed - stepRed
      currBlue = currBlue - stepBlue
      currGreen = currGreen - stepGreen
      y = y + stepY
   Next i
   
   Call SetMapMode(hdc, OldMode)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      Call .WriteProperty("Autosize", mAutosize, False)
      Call .WriteProperty("Caption", mCaption, "WordArt")
      Call .WriteProperty("FontName", mFontName, "Arial")
      Call .WriteProperty("FontItalic", mFontItalic, True)
      Call .WriteProperty("FontSize", mFontSize, 12)
      Call .WriteProperty("FontWidth", mFontWidth, 12)
      Call .WriteProperty("FontColor", mFontColor, vbBlack)
      Call .WriteProperty("FontWeight", mFontWeight, WEIGHT_NORMAL)
      Call .WriteProperty("BackColor", mBackColor, vbWhite)
      Call .WriteProperty("BackStyle", mBackStyle, waBackOpaque)
      Call .WriteProperty("FillColor", mFillColor, vbBlack)
      Call .WriteProperty("FillStyle", mFillStyle, waFillSolid)
      Call .WriteProperty("GradientStart", mGradientStart, vbBlue)
      Call .WriteProperty("GradientEnd", mGradientEnd, vbWhite)
      
      If mFillPicture Is Nothing Then
         Call .WriteProperty("FillPicture", "None", "None")
      Else
         Call .WriteProperty("FillPicture", GetStringFromPic(mFillPicture), "None")
      End If
      
      Call .WriteProperty("LineWidth", mLineWidth, 1)
      Call .WriteProperty("FontOffset", mFontOffset, 0)
      Call .WriteProperty("OffsetAngle", mOffsetAngle, 0)
      Call .WriteProperty("OffsetColor", mOffsetColor, RGB(200, 200, 200))
      Call .WriteProperty("Rotation", mRotation, 0)
      Call .WriteProperty("BottomAlignment", mBottomAlignment, waAlignmentNone)
      Call .WriteProperty("TopAlignment", mTopAlignment, waAlignmentNone)
      Call .WriteProperty("TopX1", mTopX1, 1)
      Call .WriteProperty("TopX2", mTopX2, 1)
      Call .WriteProperty("TopX3", mTopX3, 1)
      Call .WriteProperty("TopY1", mTopY1, 0)
      Call .WriteProperty("TopY2", mTopY2, 0)
      Call .WriteProperty("TopY3", mTopY3, 0)
      Call .WriteProperty("BtmX1", mBtmX1, 1)
      Call .WriteProperty("BtmX2", mBtmX2, 1)
      Call .WriteProperty("BtmX3", mBtmX3, 1)
      Call .WriteProperty("BtmY1", mBtmY1, 0)
      Call .WriteProperty("BtmY2", mBtmY2, 0)
      Call .WriteProperty("BtmY3", mBtmY3, 0)
   End With
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Dim s As String
   
   With PropBag
      mAutosize = .ReadProperty("Autosize", False)
      mCaption = .ReadProperty("Caption", "WordArt")
      mFontName = .ReadProperty("FontName", "Arial")
      mFontItalic = .ReadProperty("FontItalic", False)
      mFontSize = .ReadProperty("FontSize", 12)
      mFontColor = .ReadProperty("FontColor", vbBlack)
      mFontWidth = .ReadProperty("FontWidth", 12)
      mFontWeight = .ReadProperty("FontWeight", WEIGHT_NORMAL)
      mBackColor = .ReadProperty("BackColor", vbWhite)
      mBackStyle = .ReadProperty("BackStyle", waBackOpaque)
      mFillColor = .ReadProperty("FillColor", vbBlack)
      mFillStyle = .ReadProperty("FillStyle", waFillSolid)
      mGradientStart = .ReadProperty("GradientStart", vbBlue)
      mGradientEnd = .ReadProperty("GradientEnd", vbWhite)
      
      s = .ReadProperty("FillPicture", "None")
      If s = "None" Then
         Set mFillPicture = Nothing
      Else
         Set mFillPicture = GetPicFromString(s)
      End If
      
      mLineWidth = .ReadProperty("LineWidth", 1)
      mFontOffset = .ReadProperty("FontOffset", 0)
      mOffsetAngle = .ReadProperty("OffsetAngle", 0)
      mOffsetColor = .ReadProperty("OffsetColor", RGB(200, 200, 200))
      mRotation = .ReadProperty("Rotation", 0)
      mBottomAlignment = .ReadProperty("BottomAlignment", waAlignmentNone)
      mTopAlignment = .ReadProperty("TopAlignment", waAlignmentNone)
      mTopX1 = .ReadProperty("TopX1", 1)
      mTopX2 = .ReadProperty("TopX2", 1)
      mTopX3 = .ReadProperty("TopX3", 1)
      mTopY1 = .ReadProperty("TopY1", 0)
      mTopY2 = .ReadProperty("TopY2", 0)
      mTopY3 = .ReadProperty("TopY3", 0)
      mBtmX1 = .ReadProperty("BtmX1", 1)
      mBtmX2 = .ReadProperty("BtmX2", 1)
      mBtmX3 = .ReadProperty("BtmX3", 1)
      mBtmY1 = .ReadProperty("BtmY1", 0)
      mBtmY2 = .ReadProperty("BtmY2", 0)
      mBtmY3 = .ReadProperty("BtmY3", 0)
   End With
   
   UserControl.Font.Weight = mFontWeight
   UserControl.Font.size = mFontSize
   UserControl.Font.Italic = mFontItalic
   UserControl.Font.Name = mFontName
   UserControl.Font.Weight = mFontWeight
   
   mChangeFont = True
   mChangeTransformation = True
   mChangeTransFont = True
   mChangeFinal = True
   mChangeOffset = True
   
   Call UserControl_Paint
End Sub

