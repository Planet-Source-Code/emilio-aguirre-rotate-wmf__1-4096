VERSION 5.00
Begin VB.UserControl angButton 
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   ForeColor       =   &H8000000F&
   MousePointer    =   2  'Cross
   ScaleHeight     =   83
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   93
   ToolboxBitmap   =   "angButton.ctx":0000
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   0
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   1
      Top             =   0
      Width           =   1440
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      Height          =   1260
      Left            =   1440
      Picture         =   "angButton.ctx":0312
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2460
   End
End
Attribute VB_Name = "angButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------
' AngButton (c) Copyright Emilio Aguirre 1999
'               eaguirre@comtrade.com.mx
'_----------------------------------------------------
Option Explicit
'Types
Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

'API Declares & Constants
Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, _
        ByVal nCount As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Const BS_SOLID = 0
   
'Enumarations
Enum TraceValue
  Set_Off = 0
  Set_On = 1
End Enum

'Default Property Values:
Const m_def_Angle = 0
Const m_def_Color = vbRed
Const m_def_Trace = Set_On
Const m_PI = 3.14159265358979

'Property Variables:
Dim m_Angle As Integer
Dim m_Color As OLE_COLOR
Dim m_Trace As TraceValue
Dim m_blnMouse As Boolean
'Value of angle in degrees
Public Property Get Angle() As Integer
    Angle = m_Angle
End Property

Public Property Let Angle(ByVal New_Angle As Integer)
    m_Angle = New_Angle
    PropertyChanged "Angle"
End Property
'Draw color
Public Property Get color() As OLE_COLOR
    color = m_Color
End Property

Public Property Let color(ByVal New_Color As OLE_COLOR)
    m_Color = New_Color
    PropertyChanged "Color"
    PaintControl
End Property
'Value of trace mode
Public Property Get Trace() As TraceValue
    Trace = m_Trace
End Property

Public Property Let Trace(ByVal New_Trace As TraceValue)
    m_Trace = New_Trace
    PropertyChanged "Trace"
    PaintControl
End Property

Private Sub picMask_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If (Not m_blnMouse) Then m_blnMouse = True
  CalculateNewAngle x, y
End Sub

Private Sub picMask_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If ((Button And vbLeftButton) > 0) And (m_blnMouse) Then
   CalculateNewAngle x, y
End If
End Sub

Private Sub picMask_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If (m_blnMouse) Then m_blnMouse = False
End Sub

Private Sub picMask_Paint()
PaintControl
End Sub

Private Sub UserControl_Resize()
Height = 1200: Width = 1200 'Force to keep original values in twips
picMask.Height = 1200
picMask.Width = 1200
PaintControl
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Angle = PropBag.ReadProperty("Angle", m_def_Angle)
    m_Color = PropBag.ReadProperty("Color", m_def_Color)
    m_Trace = PropBag.ReadProperty("Trace", m_def_Trace)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Angle", m_Angle, m_def_Angle)
    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
    Call PropBag.WriteProperty("Trace", m_Trace, m_def_Trace)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Angle = m_def_Angle
    m_Color = m_def_Color
    m_Trace = m_def_Trace
End Sub

Private Sub CalculateNewAngle(x As Single, y As Single)
Dim intX As Integer
Dim intY As Integer
Dim intAngle As Integer
Dim blnNoKeep As Boolean    'Flag for prevent redrawing when it is not necessary
blnNoKeep = True
If y > 40 Then
    ' Plus button. Increments the angle by one
    If (x > 3) And (x < 15) And (y > 58) And (y < 68) Then
      intAngle = Angle + 1
      If intAngle > 180 Then intAngle = 180
    ElseIf (x > 63) And (x < 74) And (y > 58) And (y < 68) Then
    ' Minus buton. Decrements the angle by one
       intAngle = Angle - 1
       If intAngle < 0 Then intAngle = 0
    Else
       blnNoKeep = False
    End If
Else
    'Calculate the position of the click button, in a standard coordinate
    'system.
    intX = x - 40
    intY = (y - 40) * -1
    If intY = 0 Then
       If intX > 0 Then
         intAngle = 0
       Else
         intAngle = 180
        End If
    Else
      If intY > 0 Then
        If intX = 0 Then
          intAngle = 90
        Else
          intAngle = (Atn(intY / intX) * (180 / m_PI))
        End If
        If (intAngle < 0) Then intAngle = 180 + intAngle
      Else
        blnNoKeep = False 'No repainting
      End If
    End If
End If
If blnNoKeep Then
   Angle = intAngle
   PaintControl
End If
End Sub

Private Sub PaintControl()
Dim sngTheta As Single          'Angle in Radians
Dim j As Integer
Dim ang As Integer
Dim col As Long
Dim m_P(3) As POINTAPI
Dim m_R(3) As POINTAPI
Dim lb As LOGBRUSH
Dim brush As Long
Dim pen As Long

ang = Angle
col = color
m_P(0).x = 36:  m_P(0).y = 0
m_P(1).x = 27: m_P(1).y = 0
m_P(2).x = 10:  m_P(2).y = 5
m_P(3).x = 10:  m_P(3).y = -5
'Drawing the background
BitBlt picMask.hdc, 0, 0, 79, 79, picImage.hdc, 80, 0, SRCAND
BitBlt picMask.hdc, 0, 0, 79, 79, picImage.hdc, 0, 0, SRCPAINT
'Drawing the angle marker
sngTheta = -ang * m_PI / 180
If Trace = Set_Off Then
    'Trace off option
     For j = 0 To 3
      If j = 0 Then
        picMask.DrawWidth = 5
      Else
        picMask.DrawWidth = 1
      End If
      m_R(j).x = (m_P(j).x * Cos(sngTheta) - m_P(j).y * Sin(sngTheta)) + 40
      m_R(j).y = (m_P(j).x * Sin(sngTheta) + m_P(j).y * Cos(sngTheta)) + 40
      picMask.PSet (m_R(j).x, m_R(j).y), col
    Next j
    picMask.DrawWidth = 2
    
    lb.lbStyle = BS_SOLID
    lb.lbColor = col
    lb.lbHatch = 0
    brush = CreateBrushIndirect(lb)
    pen = CreatePen(0, 1, col)
    SelectObject picMask.hdc, brush
    SelectObject picMask.hdc, pen
    Polygon picMask.hdc, m_R(1), 3
    DeleteObject pen
    DeleteObject brush
Else
   'Trace on option
   picMask.DrawWidth = 5
   m_R(0).x = (m_P(0).x * Cos(sngTheta) - m_P(0).y * Sin(sngTheta)) + 40
   m_R(0).y = (m_P(0).x * Sin(sngTheta) + m_P(0).y * Cos(sngTheta)) + 40
   picMask.PSet (m_R(0).x, m_R(0).y), col
   picMask.DrawWidth = 10
   If ang > 0 Then picMask.Circle (42, 40), 20, col, 0, ang * m_PI / 180
   picMask.DrawWidth = 2
End If
'Display LED numbers
LEDNumbers picMask, 28, 57, Angle, color, 1
End Sub

Private Sub LEDNumbers(objCurrent As Object, ByVal x As Single, ByVal y As Single, ByVal intNbr As Integer, ByVal olecolor As OLE_COLOR, Optional ByVal sngScale As Single)
'----------------------------------------------------
'objCurrent     - Object where the numbers will be painted
'x,y            - Position of the upper-left point where
'                 the numbers will start to appear
'intNbr         - Value (max. 3 characters)
'olecolor       - Color for painting
'sngScale       - Scale Factor
'----------------------------------------------------
Dim intDigit As Integer     ' Next digit number for paint
Dim intLine As Integer      ' Next number line for paint.
Dim i As Integer            ' Each line number is composed by 7 lines:
Dim j As Integer            '    0  __
Dim intCurPos As Integer    '   1  |  |  2
Dim X1 As Single            '   3   __
Dim Y1 As Single            '   4  |  |  5
Dim X2 As Single            '    6  __
Dim Y2 As Single
Dim strChain As String      ' String chain number (example. for constructing number 3 we need
Dim strNum As String        ' to draw lines 0, 2, 3, 5 and 6

If sngScale = 0 Then sngScale = 1       'Default Scale Value
intCurPos = 0
strNum = Format(CStr(intNbr), "000")    'Format the output
For j = 1 To Len(strNum)
intDigit = Val(Mid$(strNum, j, 1))
    strChain = Choose(intDigit + 1, "654210", "52", "64320", "65320", "5321", "65310", "654310", "520", "6543210", "653210")
    For i = 1 To Len(strChain)
       intLine = Val(Mid$(strChain, i, 1))
       Select Case intLine
         Case 0, 3, 6                   'Drawing lines 0,3 and 6
           X1 = x + 1: Y1 = y + (intLine * 2): X2 = x + 5: Y2 = y + (intLine * 2)
         Case Else
           If intLine < 3 Then
             Y1 = y + 1: Y2 = y + 5
           Else
             Y1 = y + 7: Y2 = y + 11
           End If
           If (intLine = 1) Or (intLine = 4) Then
              X1 = x: X2 = x
           Else
             X1 = x + 6: X2 = x + 6
           End If
       End Select
       objCurrent.Line ((X1 + intCurPos) * sngScale, Y1 * sngScale)-((X2 + intCurPos) * sngScale, Y2 * sngScale), olecolor
    Next i
    intCurPos = intCurPos + 9           'skip to the next character position
Next j
End Sub
