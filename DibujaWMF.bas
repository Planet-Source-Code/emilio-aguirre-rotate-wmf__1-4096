Attribute VB_Name = "DibujaWMF"
'--------------------------------------------------------------
' Author: Emilio Aguirre
' Date  : 5/01/98
' Prog  : DibujaWMF
' Desc  : Rotate a Windows Metafile
'--------------------------------------------------------------
Option Explicit
'--------------------------------------------------------------
' Constants Definition
'--------------------------------------------------------------
Public Const PI = 3.14159265358979

Public Const META_ANIMATEPALETTE = &H436
Public Const META_ARC = &H817
Public Const META_BITBLT = &H922
Public Const META_CREATEBRUSHINDIRECT = &H2FC
Public Const META_CREATEFONTINDIRECT = &H2FB
Public Const META_CREATEPALETTE = &HF7
Public Const META_CREATEPATTERNBRUSH = &H1F9
Public Const META_CREATEPENINDIRECT = &H2FA
Public Const META_CREATEREGION = &H6FF
Public Const META_CHORD = &H830
Public Const META_DELETEOBJECT = &H1F0
Public Const META_DIBBITBLT = &H940
Public Const META_DIBCREATEPATTERNBRUSH = &H142
Public Const META_DIBSTRETCHBLT = &HB41
Public Const META_ELLIPSE = &H418
Public Const META_ESCAPE = &H626
Public Const META_EXCLUDECLIPRECT = &H415
Public Const META_EXTFLOODFILL = &H548
Public Const META_EXTTEXTOUT = &HA32
Public Const META_FILLREGION = &H228
Public Const META_FLOODFILL = &H419
Public Const META_FRAMEREGION = &H429
Public Const META_INTERSECTCLIPRECT = &H416
Public Const META_INVERTREGION = &H12A
Public Const META_LINETO = &H213
Public Const META_MOVETO = &H214
Public Const META_OFFSETCLIPRGN = &H220
Public Const META_OFFSETVIEWPORTORG = &H211
Public Const META_OFFSETWINDOWORG = &H20F
Public Const META_PAINTREGION = &H12B
Public Const META_PATBLT = &H61D
Public Const META_PIE = &H81A
Public Const META_POLYGON = &H324
Public Const META_POLYLINE = &H325
Public Const META_POLYPOLYGON = &H538
Public Const META_REALIZEPALETTE = &H35
Public Const META_RECTANGLE = &H41B
Public Const META_RESIZEPALETTE = &H139
Public Const META_RESTOREDC = &H127
Public Const META_ROUNDRECT = &H61C
Public Const META_SAVEDC = &H1E
Public Const META_SCALEVIEWPORTEXT = &H412
Public Const META_SCALEWINDOWEXT = &H410
Public Const META_SELECTCLIPREGION = &H12C
Public Const META_SELECTOBJECT = &H12D
Public Const META_SELECTPALETTE = &H234
Public Const META_SETBKCOLOR = &H201
Public Const META_SETBKMODE = &H102
Public Const META_SETDIBTODEV = &HD33
Public Const META_SETMAPMODE = &H103
Public Const META_SETMAPPERFLAGS = &H231
Public Const META_SETPALENTRIES = &H37
Public Const META_SETPIXEL = &H41F
Public Const META_SETPOLYFILLMODE = &H106
Public Const META_SETRELABS = &H105
Public Const META_SETROP2 = &H104
Public Const META_SETSTRETCHBLTMODE = &H107
Public Const META_SETTEXTALIGN = &H12E
Public Const META_SETTEXTCOLOR = &H209
Public Const META_SETTEXTCHAREXTRA = &H108
Public Const META_SETTEXTJUSTIFICATION = &H20A
Public Const META_SETVIEWPORTEXT = &H20E
Public Const META_SETVIEWPORTORG = &H20D
Public Const META_SETWINDOWEXT = &H20C
Public Const META_SETWINDOWORG = &H20B
Public Const META_STRETCHBLT = &HB23
Public Const META_STRETCHDIB = &HF43
Public Const META_TEXTOUT = &H521

'--------------------------------------------------------------
' API Types
'--------------------------------------------------------------
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type HANDLETABLE
        objectHandle(1) As Long
End Type

Type METARECORD
        rdSize As Long
        rdFunction As Integer
        rdParm(1) As Integer
End Type

Type POINT
        x As Integer
        y As Integer
End Type
'--------------------------------------------------------------
' API Declare Section
'--------------------------------------------------------------
Declare Function EnumMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMetafile As Long, ByVal lpMFEnumProc As Long, ByVal lPARAM As Long) As Long
Declare Function PlayMetaFileRecord Lib "gdi32" (ByVal hdc As Long, lpHandletable As HANDLETABLE, lpMetaRecord As METARECORD, ByVal nHandles As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long

'--------------------------------------------------------------
' Variable Declaration
'--------------------------------------------------------------
Public Angle As Integer      'Rotation Angle

Function rot(P As POINT, ByVal Angle As Double) As POINT
'--------------------------------------------------------------
'This function rotates a point P (x,y)
'--------------------------------------------------------------
Dim radians As Double
radians = (Angle * PI) / 180 'Convert the angle to radians
rot.x = P.x * Cos(radians) + (P.y * -Sin(radians))
rot.y = P.x * Sin(radians) + (P.y * Cos(radians))
End Function

Public Function EnumMetaRecord(ByVal DC As Long, ByRef lphTable As HANDLETABLE, ByRef lpMFR As METARECORD, ByVal nObj As Integer, ByVal lPARAM As Long) As Integer
'-----------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim Address As Long
Dim pt(3) As POINT
Dim dx As Integer
Dim dy As Integer
'-----------------------------------------------------------
Select Case lpMFR.rdFunction
  Case META_LINETO              'Rotate a line
    pt(0).x = lpMFR.rdParm(0)
    pt(0).y = lpMFR.rdParm(1)
    pt(0) = rot(pt(0), Angle)
    lpMFR.rdParm(0) = pt(0).x
    lpMFR.rdParm(1) = pt(0).y
  Case META_SETPIXEL
    Address = VarPtr(lpMFR)
    CopyMemory pt(0).y, ByVal Address + 8, 2
    CopyMemory pt(0).x, ByVal Address + 10, 2
    pt(0) = rot(pt(0), Angle)
    CopyMemory ByVal Address + 8, pt(0).y, 2
    CopyMemory ByVal Address + 10, pt(0).x, 2
  Case META_POLYLINE, META_POLYGON    'Rotate a Polyline, Rotate a Polygon
    Address = VarPtr(lpMFR)
    For i = 0 To lpMFR.rdParm(0) - 1
      CopyMemory pt(0).y, ByVal Address + 8 + (i * 4), 2
      CopyMemory pt(0).x, ByVal Address + 10 + (i * 4), 2
      pt(0) = rot(pt(0), Angle)
      CopyMemory ByVal Address + 8 + (i * 4), pt(0).y, 2
      CopyMemory ByVal Address + 10 + (i * 4), pt(0).x, 2
    Next i
  Case META_POLYPOLYGON
    Address = VarPtr(lpMFR)
    j = lpMFR.rdParm(0)
    l = 0
    For k = 0 To j - 1
      CopyMemory i, ByVal Address + 8 + (k * 2), 2
      l = l + i
    Next k
    For i = 0 To l - 1
      CopyMemory pt(0).y, ByVal Address + 8 + (j * 2) + (i * 4), 2
      CopyMemory pt(0).x, ByVal Address + 10 + (j * 2) + (i * 4), 2
      pt(0) = rot(pt(0), Angle)
      CopyMemory ByVal Address + 8 + (j * 2) + (i * 4), pt(0).y, 2
      CopyMemory ByVal Address + 10 + (j * 2) + (i * 4), pt(0).x, 2
    Next i
  Case META_ELLIPSE, META_RECTANGLE
    Address = VarPtr(lpMFR)
    For i = 0 To 1
      CopyMemory pt(i).y, ByVal Address + 6 + (i * 4), 2
      CopyMemory pt(i).x, ByVal Address + 8 + (i * 4), 2
    Next i
    dx = Abs((pt(1).x - pt(0).x) / 2)
    dy = Abs((pt(1).y - pt(0).y) / 2)
    If (pt(1).x > pt(0).x) Then dx = dx * -1
    If (pt(1).y > pt(0).y) Then dy = dy * -1
    pt(2).x = pt(1).x + dx
    pt(2).y = pt(1).y + dy
    pt(2) = rot(pt(2), -Angle)
    pt(0).x = pt(2).x - dx
    pt(0).y = pt(2).y - dy
    pt(1).x = pt(2).x + dx
    pt(1).y = pt(2).y + dy
    For i = 0 To 1
      CopyMemory ByVal Address + 6 + (i * 4), pt(i).y, 2
      CopyMemory ByVal Address + 8 + (i * 4), pt(i).x, 2
    Next i
 End Select

'-----------------------------------------------------------
'PlayMetaFileRecord DC, lphTable, lpMFR, nObj
'The line above must be a comment
'-----------------------------------------------------------
'Continue with the next metafile record
EnumMetaRecord = 1
End Function
