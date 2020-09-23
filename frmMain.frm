VERSION 5.00
Object = "{75D71916-58B1-4002-A7A4-6842DF94990D}#22.0#0"; "WArt.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "WordArtTest"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   960
   End
   Begin WArt.WordArt WordArt1 
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   3413
      FontItalic      =   0   'False
      FontSize        =   24
      FontWidth       =   24
      TopAlignment    =   1
      TopX1           =   0.5
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Double
Dim iStep As Double


Private Sub Form_Load()
   WordArt1.Rotation = 0
   Timer1.Interval = 100
   WordArt1.FontSize = 34
   
   
   WordArt1.TopAlignment = waAlignmentCubic
   WordArt1.BottomAlignment = waAlignmentQuadratic
   i = 0
   iStep = 0.1
   
   Call WordArt1.SetTopAlignment(0.3, i, 0.7, i / 2, 1, 0)
   Call WordArt1.SetBottomAlignment(0.2, -i, 1, 0)

   WordArt1.FontColor = vbRed
   WordArt1.FillColor = vbBlue
   WordArt1.BackStyle = 0 ' transparent
   WordArt1.BackColor = vbYellow
   WordArt1.FillStyle = waFillSolid
   WordArt1.FontOffset = 10
   WordArt1.OffsetAngle = -45
   WordArt1.offsetcolor = RGB(200, 200, 200)
   
   WordArt1.GradientEnd = vbRed
   WordArt1.GradientStart = vbBlue
   
   'Set WordArt1.FillPicture = LoadPicture("c:\vbasic\document\cshape.bmp")
   'WordArt1.FillStyle = waFillUserDefined
   
   'Timer1.Enabled = False
   WordArt1.FillStyle = waFillGradient
   
   'Timer1.Enabled = False
   'Call WordArt1.SetTopAlignment(0.3, 0.5, 0.5, -0.2, 0.8, 0.3)
   'WordArt1.TopAlignment = waAlignmentFit5
   'WordArt1.BottomAlignment = waAlignmentParallel
End Sub

Private Sub Form_Resize()
   WordArt1.Move 0, 0
   WordArt1.Width = ScaleWidth
   WordArt1.Height = ScaleHeight
End Sub

Private Sub Timer1_Timer()
   WordArt1.Rotation = WordArt1.Rotation + 10
   i = i + iStep
   If i > 1 Then
      i = 1
      iStep = -iStep
   ElseIf i < -1 Then
      i = -1
      iStep = -iStep
   End If
   
   Call WordArt1.SetTopAlignment(0.3, i, 0.7, i / 2, 1, 0)
   Call WordArt1.SetBottomAlignment(0.2, i, 1, 0)
   
End Sub

