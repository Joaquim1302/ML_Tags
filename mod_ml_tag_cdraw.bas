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
Const LV1 As Double = 6.59      ' guia vertical 1
Const LV2 As Double = 110.45    ' guia vertical 2
Const LH1 As Double = 276       ' guia horizontal 1
Const LH2 As Double = 136.402   ' guia horizontal 2

'1
    TAG1.x = ConvertUnits(LV1, cdrMillimeter, cdrInch)
    TAG1.y = ConvertUnits(LH1, cdrMillimeter, cdrInch)
'2
    TAG2.x = ConvertUnits(LV2, cdrMillimeter, cdrInch)
    TAG2.y = ConvertUnits(LH1, cdrMillimeter, cdrInch)
'3
    TAG3.x = ConvertUnits(LV1, cdrMillimeter, cdrInch)
    TAG3.y = ConvertUnits(LH2, cdrMillimeter, cdrInch)
'4
    TAG4.x = ConvertUnits(LV2, cdrMillimeter, cdrInch)
    TAG4.y = ConvertUnits(LH2, cdrMillimeter, cdrInch)
    frmMainDraw.Show vbModeless
End Sub

Sub cls_doc()
    ' Recorded 10/12/20
    ' remove todos os obejetos do documento
    ActiveLayer.Shapes.All.CreateSelection
    ActiveSelection.Cut
End Sub

Function getLastTmpCPT()
    Dim tmpPath As String
    Dim FileName As String
    Dim Directory As String
    Dim MostRecentFile As String
    Dim MostRecentDate As Date
    Dim FileSpec As String
    
    ' retorna "" se não existir
    getLastTmpCPT = ""
    
    tmpPath = Environ("Temp")
    
    'Specify the file type, if any
    FileSpec = "*.tmp"
    'specify the directory
    Directory = tmpPath & "\"
    FileName = Dir(Directory & FileSpec)
    
    If FileName <> "" Then
        MostRecentFile = FileName
        MostRecentDate = FileDateTime(Directory & FileName)
        Do While FileName <> ""
            If FileDateTime(Directory & FileName) > MostRecentDate And InStr(1, FileName, "-crl-") > 0 Then
                 MostRecentFile = FileName
                 MostRecentDate = FileDateTime(Directory & FileName)
            End If
            FileName = Dir
        Loop
    End If
    
    getLastTmpCPT = MostRecentFile

End Function


Sub clipBoardPaste()
    ' Recorded 02/03/22
    ' Recording of this command is not supported
    Dim lastTmpCPT As String
    
    lastTmpCPT = getLastTmpCPT
    
    If Len(lastTmpCPT) > 0 Then
    
        Dim impopt As StructImportOptions
        Set impopt = CreateStructImportOptions
        With impopt
            .Mode = cdrImportFull
            With .ColorConversionOptions
                .SourceColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
                .TargetColorProfileList = "sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%"
            End With
        End With
        
        Dim impflt As ImportFilter
        Set impflt = ActiveLayer.ImportEx(Environ("Temp") & "\" & lastTmpCPT, cdrCPT, impopt)
        impflt.Finish

    End If
End Sub

Sub pos_tag_1()
    
    If Not frmMainDraw.cbMover Then
        clipBoardPaste
    End If
    
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG1.x, TAG1.y
End Sub

Sub pos_tag_2()

    If Not frmMainDraw.cbMover Then
        clipBoardPaste
    End If

    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG2.x, TAG2.y
End Sub

Sub pos_tag_3()

    If Not frmMainDraw.cbMover Then
        clipBoardPaste
    End If

    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG3.x, TAG3.y
End Sub

Sub pos_tag_4()

    If Not frmMainDraw.cbMover Then
        clipBoardPaste
    End If

    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    ActiveDocument.ReferencePoint = cdrTopLeft
    OrigSelection.SetPosition TAG4.x, TAG4.y
End Sub

Sub print_pdf()
    With ActiveDocument.PrintSettings
        .SelectPrinter "Microsoft Print to PDF"
    End With
    ActiveDocument.PrintOut
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
