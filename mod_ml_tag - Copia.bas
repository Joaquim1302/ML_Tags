Attribute VB_Name = "mod_ml_tag"
Option Explicit

Type pType
    x As Integer
    y As Integer
End Type

Type tagMaskType
    W As Integer ' altura
    L As Integer ' comprimento
End Type

Type rectMaskType
    p1 As pType
    p2 As pType
End Type

Type maskMovType
    msk As rectMaskType
    newPos As pType
End Type

Type tagType
    Nome As String
    moveArr(5) As maskMovType
End Type

Dim RECTMASK As tagMaskType
Dim TAG1POS As pType
Dim TAG2POS As pType
Dim TAG3POS As pType

Dim tG1 As rectMaskType
Dim tG2 As rectMaskType
Dim tG3 As rectMaskType

Dim tag As tagType

Private Function xPoint(p As pType, x As Integer, y As Integer) As pType
    xPoint.x = p.x + x
    xPoint.y = p.y + y
End Function
    
Sub openMainForm()
    RECTMASK.L = 2343
    RECTMASK.W = 3842 '3902 -> 3842 após recorte superior
    
    TAG1POS.x = 1
    TAG1POS.y = 0
    TAG2POS.x = 2423
    TAG2POS.y = 0
    TAG3POS.x = 4845
    TAG3POS.y = 0
        
    tG1.p1 = TAG1POS
    tG1.p2 = xPoint(TAG1POS, RECTMASK.L, RECTMASK.W)
    tG2.p1 = TAG2POS
    tG2.p2 = xPoint(TAG2POS, RECTMASK.L, RECTMASK.W)
    tG3.p1 = TAG3POS
    tG3.p2 = xPoint(TAG3POS, RECTMASK.L, RECTMASK.W)
      
    frmMain.Show vbModeless
    frmMain.op_padrao = True
End Sub

Public Sub redimdoc()
    With ActiveDocument.Application.CorelScript
    ' recorte superior
        .ImageDeskewCrop 0, 6535, 3491, 3526, 2039
    End With
    With ActiveDocument.Application.CorelScript
    ' redimensiona nas medidas para o corel
        .ImageResample 7192, RECTMASK.W, 600, 600, True
    End With
End Sub

Public Sub cutPaste_tag()
    With ActiveDocument.Application.CorelScript
        .EditCopy
        .EditPasteDocument
    End With
    With ActiveDocument.Application.CorelScript
        .ObjectSelectAll
        .EditCopy
    End With
    With ActiveWindow
        .Close
    End With
    With ActiveDocument.Application.CorelScript
        .EditCutMask 5, 255, 255, 255, 0
    End With
End Sub

Private Sub classicTag(p0 As pType)
Dim m As rectMaskType
'1 ------------
    With m
        .p1.x = p0.x
        .p1.y = 2240 '2272 ESTAVA ERRADA OU FORMATO FOI ALTERADO
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 265
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -425, False
    End With
'2 -------------
    With m
        .p1.x = p0.x
        .p1.y = 2593
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 65
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -480, False
    End With
'3 -------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W - 5
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -525, False
    End With
End Sub

Private Sub linhas2Tag(p0 As pType)
Dim m As rectMaskType
'1 ------------
    With m
        .p1.x = p0.x + 97
        .p1.y = 1834
        .p2.x = .p1.x + 700
        .p2.y = .p1.y + 1000  '2134
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, 770, False
    End With
    
    classicTag p0
End Sub

Private Sub linhas2Tag2(p0 As pType)
Dim m As rectMaskType
'1 -------------
    With m
        .p1.x = p0.x
        .p1.y = 2590  '2635 ESTAVA ERRADA OU FORMATO FOI ALTERADO
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 80
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -90, False
    End With
'2 ------------
    With m
        .p1.x = p0.x + 97
        .p1.y = 1834
        .p2.x = .p1.x + 702
        .p2.y = .p1.y + 300
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, 760, False
    End With
'3 ------------
    With m
        .p1.x = p0.x
        .p1.y = 2240  '2272 ESTAVA ERRADA OU FORMATO FOI ALTERADO
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 700
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -455, False
    End With
    
'4 -------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W - 5
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -525, False
    End With
End Sub

Private Sub linhas2Tag2BAK(p0 As pType) ' ESTAVA ERRADA
Dim m As rectMaskType
'1 -------------
    With m
        .p1.x = p0.x
        .p1.y = 2595
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 80
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -50, False
    End With
'2 ------------
    With m
        .p1.x = p0.x + 97
        .p1.y = 1834
        .p2.x = .p1.x + 702
        .p2.y = .p1.y + 300
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, 800, False
    End With
'3 ------------
    With m
        .p1.x = p0.x
        .p1.y = 2272
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 700
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -425, False
    End With
    
'4 -------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -525, False
    End With
End Sub

Private Sub jadlogTag(p0 As pType)
Dim m As rectMaskType
'1 ------------
    With m
        .p1.x = p0.x
        .p1.y = 2228
        .p2.x = p0.x + RECTMASK.L
        .p2.y = 2972
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -425, False
    End With
'2 ------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -464, False
    End With
End Sub

Private Sub sequoia(p0 As pType)
Dim m As rectMaskType
'1 ------------
    With m
        .p1.x = p0.x + 97
        .p1.y = 2020
        .p2.x = .p1.x + 980
        .p2.y = .p1.y + 780
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -200, False
    End With
'2 ------------
    With m
        .p1.x = p0.x + 1290
        .p1.y = 2000
        .p2.x = .p1.x + 1050
        .p2.y = .p1.y + 880
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -300, False
    End With
'3 -------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -400, False
    End With
End Sub

Private Sub azulcargo(p0 As pType)
Dim m As rectMaskType
'1 ------------
    With m
        .p1.x = p0.x
        .p1.y = 2200
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 650
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -370, False
    End With
'2 -------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -525, False
    End With
End Sub

Private Sub totalExpress(p0 As pType)
Dim m As rectMaskType
'1 ------------
    With m
        .p1.x = p0.x + 100
        .p1.y = 1980
        .p2.x = .p1.x + 935
        .p2.y = .p1.y + 120
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 340, 0, False
    End With
'2 -------------
    With m
        .p1.x = .p1.x + 340
        .p1.y = .p1.y
        .p2.x = .p1.x + 1325
        .p2.y = .p1.y + 280
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 375, -480, False
    End With
'3 ------------
    With m
        .p1.x = p0.x
        .p1.y = 2272
        .p2.x = p0.x + RECTMASK.L
        .p2.y = .p1.y + 690
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -425, False
    End With
'4 -------------
    With m
        .p1.x = p0.x
        .p1.y = 3015
        .p2.x = p0.x + RECTMASK.L
        .p2.y = RECTMASK.W
    End With
    With ActiveDocument.Application.CorelScript
        .MaskRectangle m.p1.x, m.p1.y, m.p2.x, m.p2.y, 0
        .MaskFloaterTranslate 0, -470, False
    End With
End Sub

' ------------------------------------
Private Sub clsTag(p0 As pType)
    If frmMain.op_padrao Then
        classicTag p0
    ElseIf frmMain.op_2_linhas Then
        linhas2Tag2 p0
    ElseIf frmMain.op_jadlog Then
        jadlogTag p0
    ElseIf frmMain.op_sequoia Then
        sequoia p0
    ElseIf frmMain.op_azulcargo Then
        azulcargo p0
    ElseIf frmMain.op_total_express Then
        totalExpress p0
    End If
End Sub

Public Sub copy_tag_1()
    If Not frmMain.cbCopiar Then
        clsTag tG1.p1
    End If
    With ActiveDocument.Application.CorelScript
        .MaskRectangle tG1.p1.x, tG1.p1.y, tG1.p2.x, tG1.p2.y, 0, 0
        .EditCopy
    End With
End Sub

Public Sub copy_tag_2()
    If Not frmMain.cbCopiar Then
        clsTag tG2.p1
    End If
    With ActiveDocument.Application.CorelScript
        .MaskRectangle tG2.p1.x, tG2.p1.y, tG2.p2.x, tG2.p2.y, 0, 0
        .EditCopy
    End With
End Sub

Public Sub copy_tag_3()
    If Not frmMain.cbCopiar Then
        clsTag tG3.p1
    End If
    With ActiveDocument.Application.CorelScript
        .MaskRectangle tG3.p1.x, tG1.p1.y, tG3.p2.x, tG3.p2.y, 0, 0
        .EditCopy
    End With
End Sub

