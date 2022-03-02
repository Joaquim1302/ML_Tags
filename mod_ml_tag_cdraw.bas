Attribute VB_Name = "mod_ml_tag_cdraw"
Option Explicit

Type tagPosType
    x As Double
    y As Double
End Type

Dim TAG1 As tagPosType
Dim TAG2 As tagPosType
Dim TAG3 As tagPosType
Dim TAG4 As tagPosType

' Function ConvertUnits(Value As Double, FromUnit As cdrUnit, ToUnit As cdrUnit) As Double
'tagsPos(1).x = c(56.205, cdrMillimeter, cdrInch)

Public Sub openMainForm()


'1
    TAG1.x = ConvertUnits(6.59, cdrMillimeter, cdrInch)
    TAG1.y = ConvertUnits(278.55, cdrMillimeter, cdrInch)
'2
    TAG2.x = ConvertUnits(110.45, cdrMillimeter, cdrInch)
    TAG2.y = ConvertUnits(278.55, cdrMillimeter, cdrInch)
'3
    TAG3.x = ConvertUnits(6.59, cdrMillimeter, cdrInch)
    TAG3.y = ConvertUnits(136.402, cdrMillimeter, cdrInch)
'4
    TAG4.x = ConvertUnits(110.45, cdrMillimeter, cdrInch)
    TAG4.y = ConvertUnits(136.402, cdrMillimeter, cdrInch)
    frmMainDraw.Show vbModeless
End Sub

Sub cls_doc()
    ' Recorded 10/12/20
    ' remove todos os obejetos do documento
    ActiveLayer.Shapes.All.CreateSelection
    ActiveSelection.Cut
End Sub

Sub pos_tag_1()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG1.x, TAG1.y
End Sub

Sub pos_tag_2()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG2.x, TAG2.y
End Sub

Sub pos_tag_3()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG3.x, TAG3.y
End Sub

Sub pos_tag_4()
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG4.x, TAG4.y
End Sub

Private Sub DrawEllipse(s As Shape)
  ActiveDocument.ActiveLayer.CreateEllipse2 s.PositionX, s.PositionY, 0.1
End Sub

Sub Test()
  Dim s As Shape
  Set s = ActiveShape
  ActiveDocument.ReferencePoint = cdrTopLeft
  DrawEllipse s
  ActiveDocument.ReferencePoint = cdrTopRight
  DrawEllipse s
  ActiveDocument.ReferencePoint = cdrBottomLeft
  DrawEllipse s
  ActiveDocument.ReferencePoint = cdrBottomRight
  DrawEllipse s
  ActiveDocument.ReferencePoint = cdrCenter
  DrawEllipse s
End Sub
