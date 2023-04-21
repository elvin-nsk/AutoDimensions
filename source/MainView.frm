VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   OleObjectBlob   =   "MainView.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' Declarations
    
Private Enum DimensionUnits
    DimensionUnitsMillimeters = 0
    DimensionUnitsCentimeters = 1
    DimensionUnitsMeters = 2
    DimensionUnitsInches = 3
    DimensionUnitsPixels = 4
End Enum

Private Type typeThis
    DimensionDistance As TextBoxHandler
    DimensionLineThickness As TextBoxHandler
    DimensionTextSize As TextBoxHandler
    Quantity As TextBoxHandler
End Type
Private This As typeThis

Private sr As ShapeRange
Private Status As Boolean
Private StartingScale As Double

Private MsgNoObject As String, MsgOnlyNumbers As String, MsgEmpty As String
Private ScaleText As String, CurrentScaleText As String
Private QuantityPrefix As String

'===============================================================================
' Handlers

Private Sub UserForm_Initialize()

    With This
        Set .DimensionDistance = _
            TextBoxHandler.Create(DimensionDistance, TextBoxTypeDouble, 0)
        .DimensionDistance.Value = 10
        Set .DimensionLineThickness = _
            TextBoxHandler.Create(DimensionLineThickness, TextBoxTypeDouble, 0.0001)
        Set .DimensionTextSize = _
            TextBoxHandler.Create(DimensionTextSize, TextBoxTypeDouble, 1)
        Set .Quantity = _
            TextBoxHandler.Create(Quantity, TextBoxTypeDouble, 0)
    End With
    
    LabelUnitsOfDistance.Caption = "mm"
    Dim1.Visible = True
    
    ActiveDocument.ResetSettings
    
    ActiveDocument.Unit = cdrMillimeter ' important
    ActiveDocument.Rulers.HUnits = cdrMillimeter
    ActiveDocument.Rulers.VUnits = cdrMillimeter

    StartingScale = ActiveDocument.WorldScale
    
    Set sr = ActiveSelection.Shapes.All
    
    UILinguagem
    
End Sub

Private Sub UserForm_Terminate()
    
    '*****************************************************************
    Application.EventsEnabled = False
    Optimization = True 'True
    '*****************************************************************
    
    ActiveDocument.ClearSelection
    
    ActiveDocument.ResetSettings 'Restaurar configura��es
    
    If OptionSaveScale.Value = False Then
        ' r�tabli l'echelle du document de travail
        ActiveDocument.WorldScale = StartingScale
        ' r�tabli l'echelle du document de travail
        'ActiveDocument.Unit = StartingScale
    End If
    
    Application.EventsEnabled = True
    Optimization = False
    
    Application.Refresh
    ActiveWindow.Refresh
    
End Sub

Private Sub ComboEscala_Change()
    
    On Error GoTo Err
    
    ActiveDocument.WorldScale = _
        Replace( _
            Replace( _
                Replace(ScaleList.Text, CurrentScaleText, ""), "1/", "" _
            ), ".", "," _
        )
Err:

End Sub

Private Sub UnitsList_Change()
    Select Case UnitsList.ListIndex
        Case DimensionUnitsMillimeters
            This.DimensionDistance.Value = 10
            LabelUnitsOfDistance.Caption = "mm"
            LabelUnitsOfLineThickness.Caption = "mm"
            This.DimensionLineThickness = 0.5
        Case DimensionUnitsCentimeters
            This.DimensionDistance.Value = 1
            LabelUnitsOfDistance.Caption = "cm"
            LabelUnitsOfLineThickness.Caption = "cm"
            This.DimensionLineThickness = 0.05
        Case DimensionUnitsMeters
            This.DimensionDistance.Value = 0.01
            LabelUnitsOfDistance.Caption = "m"
            LabelUnitsOfLineThickness.Caption = "m"
            This.DimensionLineThickness = 0.0005
        Case DimensionUnitsInches
            This.DimensionDistance.Value = 0.4
            LabelUnitsOfDistance.Caption = "''"
            LabelUnitsOfLineThickness.Caption = "''"
            This.DimensionLineThickness = 0.02
        Case DimensionUnitsPixels
            This.DimensionDistance.Value = 118
            LabelUnitsOfDistance.Caption = "px"
            LabelUnitsOfLineThickness.Caption = "px"
            This.DimensionLineThickness = 5.6
    End Select
End Sub

Private Sub ButtonApply_Click()
    ActiveDocument.ResetSettings
    
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox MsgNoObject, vbExclamation, MainView.Caption
        Exit Sub
    End If
    
    With ActiveDocument
        Select Case UnitsList.ListIndex
            Case DimensionUnitsMillimeters
                .Unit = cdrMillimeter ' important
                .Rulers.HUnits = cdrMillimeter
                .Rulers.VUnits = cdrMillimeter
            Case DimensionUnitsCentimeters
                .Unit = cdrCentimeter ' important
                .Rulers.HUnits = cdrCentimeter
                .Rulers.VUnits = cdrCentimeter
            Case DimensionUnitsMeters
                .Unit = cdrMeter ' important
                .Rulers.HUnits = cdrMeter
                .Rulers.VUnits = cdrMeter
            Case DimensionUnitsInches
                .Unit = cdrInch ' important
                .Rulers.HUnits = cdrInch
                .Rulers.VUnits = cdrInch
            Case DimensionUnitsPixels
                .Unit = cdrPixel ' important
                .Rulers.HUnits = cdrPixel
                .Rulers.VUnits = cdrPixel
        End Select
    
        .BeginCommandGroup DIMENSIONS_STR
    
    End With
    
    Set sr = ActiveSelection.Shapes.All
    
    ComboEscala_Change
    
    Init
    
    Actualise
    
    ActiveDocument.EndCommandGroup
    
    Application.Refresh
    ActiveWindow.Refresh
    
End Sub

Private Sub ButtonReset_Click()

    Dim DimensionShapes As ShapeRange
    
    Set DimensionShapes = ActivePage.Shapes.FindShapes(DIMENSIONS_STR)
    
    If DimensionShapes.Count = 0 Then Exit Sub
    
    '*****************************************************************
    Application.EventsEnabled = False
    Optimization = True 'True
    '*****************************************************************
    
    ActiveDocument.BeginCommandGroup DIMENSIONS_STR
    
    Dim Shape As Shape
    For Each Shape In DimensionShapes
        Shape.Delete
    Next Shape
    
    ActiveDocument.EndCommandGroup
    
    '*****************************************************************
    Application.EventsEnabled = True
    Optimization = False
    Application.Refresh
    '*****************************************************************
    
End Sub
 
Private Sub OptionSimple_Change()
    If OptionSimple.Value = True Then
        Position8.Value = True
        ActivePosition False
    Else
        Position1.Value = True
        ActivePosition True
    End If
End Sub

Private Sub ScaleList_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii < vbKey0 And KeyAscii <> 47 Then
        KeyAscii = 0
        MsgBox MsgOnlyNumbers, vbExclamation, MainView.Caption
    ElseIf KeyAscii > vbKey9 And KeyAscii <> 47 Then
        KeyAscii = 0
        MsgBox MsgOnlyNumbers, vbExclamation, MainView.Caption
    End If
End Sub

Private Sub ScaleList_AfterUpDate()
    If ScaleList = "" Then
        ScaleList.ListIndex = 0
        MsgBox MsgEmpty, vbExclamation, MainView.Caption
    End If
End Sub

Private Sub Position1_Click()
    SetVisibleDimAndWriteSetting 1
End Sub
Private Sub Position2_Click()
    SetVisibleDimAndWriteSetting 2
End Sub
Private Sub Position3_Click()
    SetVisibleDimAndWriteSetting 3
End Sub
Private Sub Position4_Click()
    SetVisibleDimAndWriteSetting 4
End Sub
Private Sub Position5_Click()
    SetVisibleDimAndWriteSetting 5
End Sub
Private Sub Position6_Click()
    SetVisibleDimAndWriteSetting 6
End Sub
Private Sub Position7_Click()
    SetVisibleDimAndWriteSetting 7
End Sub
Private Sub Position8_Click()
    SetVisibleDimAndWriteSetting 8
End Sub

Private Sub UrlKafard_Click()
    OpenUrl "http://kafard62.free.fr"
End Sub

Private Sub UrlFabricePayPal_Click()
    OpenUrl "www.paypal.me/FabriceVN/2"
End Sub

Private Sub UrlCorelNaVeia_Click()
    OpenUrl "https://corelnaveia.com"
End Sub

Private Sub UrlFerreira_Click()
    OpenUrl "https://wa.link/bp8m8j"
End Sub

'===============================================================================
' Helpers

Private Function Init()
    
    ScaleList.Clear
    ScaleList.AddItem CurrentScaleText & ActiveDocument.WorldScale, 0
    
    ScaleList.AddItem "1/1"
    ScaleList.AddItem "1/2"
    ScaleList.AddItem "1/4"
    ScaleList.AddItem "1/5"
    ScaleList.AddItem "1/10"
    ScaleList.AddItem "1/15"
    ScaleList.AddItem "1/20"
    ScaleList.AddItem "1/40"
    ScaleList.AddItem "1/50"
    
    'RecupVal = GetSetting(APP_NAME, DIMENSIONS_STR, "position", "1")
    
    'Select Case RecupVal
    '    Case "1"
    '        Position1.Value = True
    '    Case "2"
    '        Position2.Value = True
    '    Case "3"
    '        Position3.Value = True
    '    Case "4"
    '        Position4.Value = True
    '    Case "5"
    '        Position5.Value = True
    '    Case "6"
    '        Position6.Value = True
    '    Case "7"
    '        Position7.Value = True
    '    Case "8"
    '        Position8.Value = True
    'End Select
    
    ScaleList.ListIndex = 0
    
End Function

Private Function DrawDimensions( _
                     Shapes As Shapes, _
                     Spacing As Long, _
                     TextHeight As Double, _
                     TechnicalDrawing As Boolean _
                 )
    
    Dim ShOrign As Shape
    Dim Shape As Shape
    
    Const BallEnd As Long = 53
    Const LineEnd As Long = 59
    
    Dim HorizontalDimension As Shape
    Dim TextHorizontal As Shape
    Dim VerticalDimension As Shape
    Dim TextVertical As Shape
    Dim InfoScaleQuantity As Shape
    Dim Text As String
    Dim Line1 As Shape
    Dim Line2 As Shape
    Dim Line3 As Shape
    Dim Line4 As Shape
    
    Dim UnitsSuffix As String
    
    Dim AllShapes As New ShapeRange
    
    ActiveDocument.PreserveSelection = True
    
    If OptionUnits.Value = True Then UnitsSuffix = LabelUnitsOfDistance.Caption
    
    If TechnicalDrawing Then
        
        For Each ShOrign In Shapes
            
            ''''''''''''''''' C�tes HORIZONTALES
            If Position1.Value = True _
            Or Position7.Value = True _
            Or Position8.Value = True Then
                
                ' C�tes Horizontales
                Set HorizontalDimension = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX, _
                        ShOrign.BottomY - This.DimensionDistance, _
                        ShOrign.RightX, _
                        ShOrign.BottomY - This.DimensionDistance _
                    )
                HorizontalDimension.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(8), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(BallEnd), _
                    ArrowHeads(BallEnd), _
                    cdrFalse, cdrTrue, _
                    cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Set TextHorizontal = _
                    ActiveLayer.CreateArtisticText( _
                        HorizontalDimension.CenterX, _
                        HorizontalDimension.CenterY, _
                        VBA.Format( _
                            ShOrign.SizeWidth * ActiveDocument.WorldScale, _
                            "##0.0##" _
                        ) & " " & UnitsSuffix & vbCrLf, _
                        , , DIMENSIONS_FONT, TextHeight, cdrTrue, cdrFalse, , _
                        cdrCenterAlignment _
                    )
                TextHorizontal.SetPosition _
                    HorizontalDimension.CenterX _
                  - (TextHorizontal.SizeWidth / 2), _
                    HorizontalDimension.CenterY _
                  - (TextHorizontal.SizeHeight * 0.5)
                    
                Set Line3 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX, _
                        ShOrign.BottomY - This.DimensionLineThickness, _
                        ShOrign.LeftX, _
                        ShOrign.BottomY - This.DimensionDistance _
                      - This.DimensionLineThickness _
                    )
                Set Line4 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.RightX, _
                        ShOrign.BottomY - This.DimensionLineThickness, _
                        ShOrign.RightX, _
                        ShOrign.BottomY - This.DimensionDistance _
                      - This.DimensionLineThickness _
                    )
                Line3.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                Line4.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                
                HorizontalDimension.Name = DIMENSIONS_STR
                TextHorizontal.Name = DIMENSIONS_STR
                
                Line3.Name = DIMENSIONS_STR
                Line4.Name = DIMENSIONS_STR
                    
                ' Ajuste le texte si besoin
                If OptionAjustText.Value = True Then
                    If TextHorizontal.SizeWidth >= ShOrign.SizeWidth Then
                        TextHorizontal.SizeWidth = ShOrign.SizeWidth - 1
                    End If
                    TextHorizontal.AlignToShape cdrAlignHCenter, ShOrign
                End If
                
            End If
            
            If Position3.Value = True _
            Or Position4.Value = True _
            Or Position5.Value = True Then
                
                ' C�tes Horizontales
                Set HorizontalDimension = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX, _
                        ShOrign.TopY + This.DimensionDistance, _
                        ShOrign.RightX, _
                        ShOrign.TopY + This.DimensionDistance _
                    )
                    HorizontalDimension.Outline.SetProperties _
                        This.DimensionLineThickness, _
                        OutlineStyles(8), _
                        CreateColor(DIMENSIONS_COLOR), _
                        ArrowHeads(BallEnd), _
                        ArrowHeads(BallEnd), _
                        cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                        cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Set TextHorizontal = _
                    ActiveLayer.CreateArtisticText( _
                        HorizontalDimension.CenterX, _
                        HorizontalDimension.CenterY, _
                        VBA.Format( _
                            ShOrign.SizeWidth * ActiveDocument.WorldScale, _
                            "##0.0##" _
                        ) & " " & UnitsSuffix & vbCrLf, _
                        , , DIMENSIONS_FONT, TextHeight, cdrTrue, cdrFalse, , _
                        cdrCenterAlignment _
                    )
                    TextHorizontal.SetPosition _
                        HorizontalDimension.CenterX _
                      - (TextHorizontal.SizeWidth / 2), _
                        HorizontalDimension.CenterY _
                      + (TextHorizontal.SizeHeight * 1.5)
                    
                Set Line3 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX, _
                        ShOrign.TopY - This.DimensionLineThickness, _
                        ShOrign.LeftX, _
                        ShOrign.TopY + This.DimensionDistance _
                      + This.DimensionLineThickness _
                    )
                Set Line4 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.RightX, _
                        ShOrign.TopY - This.DimensionLineThickness, _
                        ShOrign.RightX, _
                        ShOrign.TopY + This.DimensionDistance _
                      + This.DimensionLineThickness _
                    )
                Line3.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                Line4.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Line3.Name = DIMENSIONS_STR
                Line4.Name = DIMENSIONS_STR
                
                HorizontalDimension.Name = DIMENSIONS_STR
                TextHorizontal.Name = DIMENSIONS_STR
                    
                ' Ajuste le texte si besoin
                If OptionAjustText.Value = True Then
                    If TextHorizontal.SizeWidth >= ShOrign.SizeWidth Then
                        TextHorizontal.SizeWidth = ShOrign.SizeWidth - 1
                    End If
                    TextHorizontal.AlignToShape cdrAlignHCenter, ShOrign
                End If
                
            End If
            
            ''''''''''''''''' C�tes VERTICALES
            If Position1.Value = True _
            Or Position2.Value = True _
            Or Position3.Value = True Then
                
                ' C�tes verticales
                Set VerticalDimension = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX - This.DimensionDistance, _
                        ShOrign.TopY, _
                        ShOrign.LeftX - This.DimensionDistance, _
                        ShOrign.BottomY _
                    )
                VerticalDimension.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(8), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(BallEnd), _
                    ArrowHeads(BallEnd), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Set TextVertical = _
                    ActiveLayer.CreateArtisticText( _
                        VerticalDimension.CenterX, _
                        VerticalDimension.CenterY, _
                        VBA.Format( _
                            ShOrign.SizeHeight * ActiveDocument.WorldScale, _
                            "##0.0##" _
                        ) & " " & UnitsSuffix, _
                        , , DIMENSIONS_FONT, TextHeight, cdrTrue, cdrFalse, , _
                        cdrCenterAlignment _
                    )
                TextVertical.Rotate 90#
                TextVertical.PositionX = _
                TextVertical.PositionX - TextVertical.SizeWidth
                TextVertical.CenterY = VerticalDimension.CenterY
                    
                Set Line1 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX - This.DimensionLineThickness, _
                        ShOrign.TopY, _
                        ShOrign.LeftX - This.DimensionDistance _
                      - This.DimensionLineThickness, _
                        ShOrign.TopY _
                    )
                Set Line2 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.LeftX - This.DimensionLineThickness, _
                        ShOrign.BottomY, _
                        ShOrign.LeftX - This.DimensionDistance _
                      - This.DimensionLineThickness, _
                        ShOrign.BottomY _
                    )
                Line1.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                Line2.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Line1.Name = DIMENSIONS_STR
                Line2.Name = DIMENSIONS_STR
                
                VerticalDimension.Name = DIMENSIONS_STR
                TextVertical.Name = DIMENSIONS_STR
                    
                ' Ajuste le texte si besoin
                If OptionAjustText.Value = True Then
                    If TextVertical.SizeHeight >= ShOrign.SizeHeight Then
                        TextVertical.SizeHeight = ShOrign.SizeHeight - 1
                    End If
                        TextVertical.AlignToShape cdrAlignVCenter, ShOrign
                End If
                
            End If
            
            If Position5.Value = True _
            Or Position7.Value = True _
            Or Position6.Value = True Then
                
                ' C�tes verticales
                Set VerticalDimension = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.RightX + This.DimensionDistance, _
                        ShOrign.TopY, _
                        ShOrign.RightX + This.DimensionDistance, _
                        ShOrign.BottomY _
                    )
                VerticalDimension.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(8), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(BallEnd), _
                    ArrowHeads(BallEnd), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Set TextVertical = _
                    ActiveLayer.CreateArtisticText( _
                        VerticalDimension.CenterX, _
                        VerticalDimension.CenterY, _
                        VBA.Format( _
                            ShOrign.SizeHeight * ActiveDocument.WorldScale, _
                            "##0.0##" _
                        ) & " " & UnitsSuffix, _
                        , , DIMENSIONS_FONT, TextHeight, cdrTrue, cdrFalse, , _
                        cdrCenterAlignment _
                    )
                TextVertical.Rotate -90#
                TextVertical.PositionX = _
                    TextVertical.PositionX + TextVertical.SizeWidth
                TextVertical.CenterY = VerticalDimension.CenterY
                    
                Set Line1 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.RightX + This.DimensionLineThickness, _
                        ShOrign.TopY, _
                        ShOrign.RightX + This.DimensionDistance _
                      + This.DimensionLineThickness, ShOrign.TopY _
                    )
                Set Line2 = _
                    ActiveLayer.CreateLineSegment( _
                        ShOrign.RightX + This.DimensionLineThickness, _
                        ShOrign.BottomY, _
                        ShOrign.RightX + This.DimensionDistance _
                      + This.DimensionLineThickness, ShOrign.BottomY _
                    )
                Line1.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                Line2.Outline.SetProperties _
                    This.DimensionLineThickness, _
                    OutlineStyles(0), _
                    CreateColor(DIMENSIONS_COLOR), _
                    ArrowHeads(0), ArrowHeads(0), _
                    cdrFalse, cdrTrue, cdrOutlineButtLineCaps, _
                    cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#
                    
                Line1.Name = DIMENSIONS_STR
                Line2.Name = DIMENSIONS_STR
                
                VerticalDimension.Name = DIMENSIONS_STR
                TextVertical.Name = DIMENSIONS_STR
                    
                ' Ajuste le texte si besoin
                If OptionAjustText.Value = True Then
                    If TextVertical.SizeHeight >= ShOrign.SizeHeight Then
                        TextVertical.SizeHeight = ShOrign.SizeHeight - 1
                    End If
                    TextVertical.AlignToShape cdrAlignVCenter, ShOrign
                End If
                
            End If
            
            ' Pr�cision echelle et quantit�
            If OptionScale.Value = True Or This.Quantity.Value <> "" Then
                Dim d�calage As Double
                
                If OptionScale.Value = True Then
                    Text = ScaleText & ActiveDocument.WorldScale
                Else
                    Text = ""
                End If
                If Text <> "" And This.Quantity.Value > 0 Then _
                    Text = Text & " - "
                If This.Quantity.Value > 0 Then _
                    Text = Text & QuantityPrefix & This.Quantity.Value
                
                If Position1.Value = True _
                Or Position7.Value = True _
                Or Position8.Value = True Then
                    d�calage = HorizontalDimension.BottomY
                Else
                    If Position2.Value = True Or Position6.Value = True Then
                        d�calage = ShOrign.BottomY '+ 4
                    Else
                        d�calage = ShOrign.BottomY '+ 4
                    End If
                End If
                
                Set InfoScaleQuantity = _
                    ActiveLayer.CreateArtisticText( _
                        ShOrign.CenterX, _
                        d�calage, _
                        Text, , , DIMENSIONS_FONT, TextHeight, _
                        cdrTrue, cdrFalse, , cdrCenterAlignment _
                    )
                InfoScaleQuantity.PositionY = _
                    InfoScaleQuantity.PositionY _
                  - 2 * (InfoScaleQuantity.SizeHeight * 1.5)
                InfoScaleQuantity.Name = DIMENSIONS_STR
                    
            End If
            
            If Position1.Value = True _
            Or Position3.Value = True _
            Or Position5.Value = True _
            Or Position7.Value = True Then
                If OptionGroupDimensions.Value = True Then
                    If OptionScale.Value = True Or This.Quantity.Value <> "" Then
                        Set Shape = _
                            ActiveDocument.CreateShapeRangeFromArray( _
                                VerticalDimension, HorizontalDimension, _
                                TextHorizontal, TextVertical, _
                                InfoScaleQuantity, Text, _
                                Line1, Line2, Line3, Line4 _
                            ).Group
                    Else
                        Set Shape = _
                            ActiveDocument.CreateShapeRangeFromArray( _
                                VerticalDimension, HorizontalDimension, _
                                TextHorizontal, TextVertical, _
                                Line1, Line2, Line3, Line4 _
                            ).Group
                    End If
                End If
            End If
            
            If Position2.Value = True Or Position6.Value = True Then
                If OptionGroupDimensions.Value = True Then
                    If OptionScale.Value = True Or This.Quantity.Value <> "" Then
                        Set Shape = _
                            ActiveDocument.CreateShapeRangeFromArray( _
                                VerticalDimension, TextVertical, _
                                InfoScaleQuantity, Text, _
                                Line1, Line2 _
                            ).Group
                    Else
                        Set Shape = _
                            ActiveDocument.CreateShapeRangeFromArray( _
                                VerticalDimension, TextVertical, _
                                Line1, Line2 _
                            ).Group
                    End If
                End If
            End If
            
            If Position4.Value = True Or Position8.Value = True Then
                If OptionGroupDimensions.Value = True Then
                    If OptionScale.Value = True Or This.Quantity.Value <> "" Then
                        Set Shape = _
                            ActiveDocument.CreateShapeRangeFromArray( _
                                HorizontalDimension, TextHorizontal, _
                                InfoScaleQuantity, Text, _
                                Line3, Line4 _
                            ).Group
                    Else
                        Set Shape = _
                            ActiveDocument.CreateShapeRangeFromArray( _
                                HorizontalDimension, TextHorizontal, _
                                Line3, Line4 _
                            ).Group
                    End If
                End If
            End If
            
        Next ShOrign
    Else
        
        For Each ShOrign In Shapes
            
            ' C�tes Horizontales
            Set TextHorizontal = _
                ActiveLayer.CreateArtisticText( _
                    ShOrign.CenterX, _
                    ShOrign.BottomY - This.DimensionDistance, _
                    VBA.Format( _
                        ShOrign.SizeWidth * ActiveDocument.WorldScale, _
                        "##0.0##" _
                    ) & "x" _
                  & VBA.Format( _
                        ShOrign.SizeHeight * ActiveDocument.WorldScale, _
                        "##0.0##" _
                    ) & " " & UnitsSuffix, _
                    , , DIMENSIONS_FONT, TextHeight, cdrTrue, cdrFalse, , _
                    cdrCenterAlignment _
                )
            'TextHorizontal.PositionY = TextHorizontal.PositionY - (TextHorizontal.PositionY / 1.5) / TextHorizontal.PositionY
            TextHorizontal.SetPosition _
                TextHorizontal.CenterX - (TextHorizontal.SizeWidth / 2), _
                TextHorizontal.BottomY
            
            ' Renomme tout pour mieux retrouver
            TextHorizontal.Name = DIMENSIONS_STR
                
            ' Pr�cision echelle et quantit�
            If OptionScale.Value = True Or This.Quantity.Value <> "" Then
                If OptionScale.Value = True Then
                    Text = ScaleText & ActiveDocument.WorldScale
                Else
                    Text = ""
                End If
                If Text <> "" And This.Quantity.Value > 0 Then _
                    Text = Text & " - "
                If This.Quantity.Value > 0 Then _
                    Text = Text & QuantityPrefix & This.Quantity.Value
                
                Set InfoScaleQuantity = _
                    ActiveLayer.CreateArtisticText( _
                        TextHorizontal.CenterX, _
                        TextHorizontal.CenterY, _
                        Text, , , DIMENSIONS_FONT, TextHeight, cdrTrue, cdrFalse, , _
                        cdrCenterAlignment _
                    )
                InfoScaleQuantity.Outline.SetNoOutline
                InfoScaleQuantity.PositionY = _
                    InfoScaleQuantity.PositionY _
                  - (InfoScaleQuantity.SizeHeight * 2)
                InfoScaleQuantity.Name = DIMENSIONS_STR
            End If
            
            If OptionGroupDimensions.Value = True Then
                On Error Resume Next
                Set Shape = _
                    ActiveDocument.CreateShapeRangeFromArray( _
                        TextHorizontal, InfoScaleQuantity _
                    ).Group
            End If
            
        Next ShOrign
        
    End If
    
    ActiveDocument.ClearSelection
    sr.CreateSelection
    
End Function

Private Function Actualise()
    
    '*****************************************************************
    Application.EventsEnabled = False
    Optimization = True 'True
    '*****************************************************************
    
    WriteSettings
    
    'Atualizar cotas =======================
    'ActiveDocument.ActivePage.ActiveLayer.Shapes.FindShapes(DIMENSIONS_STR).Delete
    
    sr.AddToSelection
    
    DrawDimensions _
        sr.Shapes, _
        This.DimensionDistance.Value, _
        This.DimensionTextSize.Value, _
        OptionTechnic.Value
    
    '*****************************************************************
    Application.EventsEnabled = True
    Optimization = False
    Application.Refresh
    '*****************************************************************
    
End Function

Private Sub ActivePosition(Valeur As Boolean)
    Position1.Enabled = Valeur
    Position2.Enabled = Valeur
    Position3.Enabled = Valeur
    Position4.Enabled = Valeur
    Position5.Enabled = Valeur
    Position6.Enabled = Valeur
    Position7.Enabled = Valeur
    Position8.Enabled = Valeur
End Sub

Private Sub SetVisibleDimAndWriteSetting(ByVal DimIndex As Long)
    WriteSetting "position", DimIndex
    Dim1.Visible = False
    Dim2.Visible = False
    Dim3.Visible = False
    Dim4.Visible = False
    Dim5.Visible = False
    Dim6.Visible = False
    Dim7.Visible = False
    Dim8.Visible = False
    Select Case DimIndex
        Case 1
            Dim1.Visible = True
        Case 2
            Dim2.Visible = True
        Case 3
            Dim3.Visible = True
        Case 4
            Dim4.Visible = True
        Case 5
            Dim5.Visible = True
        Case 6
            Dim6.Visible = True
        Case 7
            Dim7.Visible = True
        Case 8
            Dim8.Visible = True
    End Select
End Sub

Private Sub OpenUrl(ByVal Url As String)
    With VBA.CreateObject("WScript.Shell")
        .Run Url
    End With
End Sub

Private Function WriteSettings()
    With This
        WriteSetting "espacement", .DimensionDistance
        WriteSetting "hauteurtypo", .DimensionTextSize
        WriteSetting "linhacota", .DimensionLineThickness
        WriteSetting "technique", OptionTechnic
    End With
End Function

Private Function ReadSetting( _
                     ByVal Key As String, _
                     Optional ByVal DefaultSetting As String _
                 ) As String
    ReadSetting = GetSetting(APP_NAME, DIMENSIONS_STR, Key, DefaultSetting)
End Function

Private Sub WriteSetting( _
                ByVal Key As String, _
                ByVal Setting As String _
            )
    SaveSetting APP_NAME, DIMENSIONS_STR, Key, Setting
End Sub

Private Sub UILinguagem()
    On Error Resume Next
    
    DoEvents
    
    Select Case Application.UILanguage
        Case cdrEnglishUS 'cdrEnglishUS (1033)
            English_US
        Case cdrSpanish 'cdrSpanish (1034)
            Espanol_ES
        Case cdrFrench 'cdrFrench (1036)
            Francais_FR
        Case cdrRussian 'cdrRussian (1049)
            Russian_RU
        Case Else 'Idioma padr�o
            Portugues_BR
    End Select
    
End Sub

'===============================================================================
' Localizations

Private Sub Portugues_BR()
    On Error Resume Next
    
    ' Fonction ajout� le 24/10/16 pour r�cup�rer la version automatiquement depuis le GMS
    Me.Caption = "Dimens�o Autom�tica | BR" & APP_VERSION
        
    FrameParameters.Caption = "Par�metros"
    FrameTextType.Caption = "Tipo de Texto"
    LabelUnits.Caption = "Unidades"
    
    UrlCorelNaVeia.ControlTipText = "Acesse Outras Macros no CorelNaVeia.com..."
    UrlKafard.ControlTipText = "Acesse o Site do Fabrice para outras Macros..."
    UrlFerreira.ControlTipText = "Contate Ferreira Felipe pelo WhatsApp para outras Macros..."
    UrlFabricePayPal.ControlTipText = "Pague um Cafezinho ao Fabrice..."
    UnitsList.ControlTipText = "Selecionar Unidades de Medida"
    DimensionDistance.ControlTipText = "Dist�ncia da Cota"
    DimensionTextSize.ControlTipText = "Fonte da Cota"
    Quantity.ControlTipText = "Adicionar Quantitidades e Detalhes"
    QuantityPrefix = "Quantitat: "
    DimensionLineThickness.ControlTipText = "Linha da Cota"
    ScaleList.ControlTipText = "Selecionar/Incluir Escala"
    
    LabelScale.Caption = "Escalas"
    OptionScale.Caption = "Incluir Escala"
    OptionAjustText.Caption = "Ajustar Legenda"
    OptionUnits.Caption = "Incluir Legenda"
    LabelQuantity.Caption = "Quantitativo"
    OptionTechnic.Caption = "T�cnico"
    OptionSimple.Caption = "Simples"
    OptionGroupDimensions.Caption = "Agrupar Cotas"
    OptionSaveScale.Caption = "Salvar Escala Atual!"
    ButtonApply.Caption = "APLICAR"
    ButtonReset.Caption = "LIMPAR"
    
    UnitsList.AddItem "Mil�metros"
    UnitsList.AddItem "Cent�metros"
    UnitsList.AddItem "Metros"
    UnitsList.AddItem "Polegadas"
    UnitsList.AddItem "Pixels"
    UnitsList.ListIndex = 0
    
    ScaleText = "Escala 1/"
    CurrentScaleText = "(Atual) 1/"
    
    ScaleList.AddItem CurrentScaleText & ActiveDocument.WorldScale, 0
    ScaleList.AddItem "1/1"
    ScaleList.AddItem "1/2"
    ScaleList.AddItem "1/4"
    ScaleList.AddItem "1/5"
    ScaleList.AddItem "1/10"
    ScaleList.AddItem "1/15"
    ScaleList.AddItem "1/20"
    ScaleList.AddItem "1/40"
    ScaleList.AddItem "1/50"
    ScaleList.ListIndex = 0
    
    'Notifications
    MsgNoObject = "Selecione uma ou mais formas previamente!"
    MsgOnlyNumbers = "Insira apenas n�meros!"
    MsgEmpty = "Esse campo n�o pode ficar vazio!"
    
End Sub

Private Sub Francais_FR()
    On Error Resume Next
    
    ' Fonction ajout� le 24/10/16 pour r�cup�rer la version automatiquement depuis le GMS
    Me.Caption = "Dimension Automatique | FR & APP_VERSION"
        
    FrameParameters.Caption = "Param�tres"
    FrameTextType.Caption = "Type de texte"
    LabelUnits.Caption = "Unit�s"
    
    UrlCorelNaVeia.ControlTipText = "Aller � Autres macros sur CorelNaVeia.com..."
    UrlKafard.ControlTipText = "Visitez le site de Fabrice pour d�autres macros..."
    UrlFerreira.ControlTipText = "Contactez Ferreira Felipe sur WhatsApp pour d�autres macros..."
    UrlFabricePayPal.ControlTipText = "Acheter une tasse de caf� � Fabrice..."
    UnitsList.ControlTipText = "S�lectionner les Unit�s de mesure"
    DimensionDistance.ControlTipText = "Distance de Quota"
    DimensionTextSize.ControlTipText = "Police de Quota"
    Quantity.ControlTipText = "Ajouter des Quantit�s et des D�tails"
    QuantityPrefix = "Quantit�: "
    DimensionLineThickness.ControlTipText = "Ligne de Quota"
    ScaleList.ControlTipText = "S�lectionner/Inclure l��chelle"
    
    LabelScale.Caption = "�chelles"
    OptionScale.Caption = "Inclure l��chelle"
    OptionAjustText.Caption = "Ajuster la L�gende"
    OptionUnits.Caption = "Inclure la L�gende"
    LabelQuantity.Caption = "Quantitatif"
    OptionTechnic.Caption = "Technique"
    OptionSimple.Caption = "Simple"
    OptionGroupDimensions.Caption = "Groupe Quotas"
    OptionSaveScale.Caption = "Enregistrer l'�chelle actuelle!"
    ButtonApply.Caption = "APPLIQUER"
    ButtonReset.Caption = "NETTOYER"
    
    UnitsList.AddItem "Millim�tres"
    UnitsList.AddItem "Centim�tres"
    UnitsList.AddItem "M�tres"
    UnitsList.AddItem "Pouces"
    UnitsList.AddItem "Pixels"
    UnitsList.ListIndex = 0
    
    ScaleText = "�chelle 1/"
    CurrentScaleText = "(Actuel) 1/"
    
    ScaleList.AddItem CurrentScaleText & ActiveDocument.WorldScale, 0
    ScaleList.AddItem "1/1"
    ScaleList.AddItem "1/2"
    ScaleList.AddItem "1/4"
    ScaleList.AddItem "1/5"
    ScaleList.AddItem "1/10"
    ScaleList.AddItem "1/15"
    ScaleList.AddItem "1/20"
    ScaleList.AddItem "1/40"
    ScaleList.AddItem "1/50"
    ScaleList.ListIndex = 0
    
    'Notifications
    MsgNoObject = "S�lectionnez une ou plusieurs formes � l�avance!"
    MsgOnlyNumbers = "N�entrez que des chiffres!"
    MsgEmpty = "Ce champ ne peut pas �tre vide!"
    
End Sub

Private Sub English_US()
    On Error Resume Next
    
    ' Fonction ajout� le 24/10/16 pour r�cup�rer la version automatiquement depuis le GMS
    Me.Caption = "Automatic Dimensions | US " & APP_VERSION
        
    FrameParameters.Caption = "Parameters"
    FrameTextType.Caption = "Text Type"
    LabelUnits.Caption = "Units"
    
    UrlCorelNaVeia.ControlTipText = "Access other Macros on CorelNaVeia.com ..."
    UrlKafard.ControlTipText = "Visit the Fabrice Website for other Macros..."
    UrlFerreira.ControlTipText = "Contact Ferreira Felipe on WhatsApp for other macros..."
    UrlFabricePayPal.ControlTipText = "Pay one coffe for Fabrice..."
    UnitsList.ControlTipText = "Choose Measurement Units"
    DimensionDistance.ControlTipText = "Distance Quote"
    DimensionTextSize.ControlTipText = "Font Quote"
    Quantity.ControlTipText = "Add Quantities and Details"
    QuantityPrefix = "Quantity: "
    DimensionLineThickness.ControlTipText = "Line Quote"
    ScaleList.ControlTipText = "Select/Add Scale"
    
    LabelScale.Caption = "Scales"
    OptionScale.Caption = "Add Scale"
    OptionAjustText.Caption = "Ajust Legend"
    OptionUnits.Caption = "Add Legend"
    LabelQuantity.Caption = "Quantitative"
    OptionTechnic.Caption = "Technician"
    OptionSimple.Caption = "Simple"
    OptionGroupDimensions.Caption = "Group Quotes"
    OptionSaveScale.Caption = "Save Current Scale!"
    ButtonApply.Caption = "APPLY"
    ButtonReset.Caption = "RESET ALL"
    
    UnitsList.AddItem "Millimeters"
    UnitsList.AddItem "Centimeters"
    UnitsList.AddItem "Meters"
    UnitsList.AddItem "Inches"
    UnitsList.AddItem "Pixels"
    UnitsList.ListIndex = 0
    
    ScaleText = "Scale 1/"
    CurrentScaleText = "(Current) 1/"
    
    ScaleList.AddItem CurrentScaleText & ActiveDocument.WorldScale, 0
    ScaleList.AddItem "1/1"
    ScaleList.AddItem "1/2"
    ScaleList.AddItem "1/4"
    ScaleList.AddItem "1/5"
    ScaleList.AddItem "1/10"
    ScaleList.AddItem "1/15"
    ScaleList.AddItem "1/20"
    ScaleList.AddItem "1/40"
    ScaleList.AddItem "1/50"
    ScaleList.ListIndex = 0
    
    'Notifications
    MsgNoObject = "Select one or more shapes in advance!"
    MsgOnlyNumbers = "Enter only numbers!"
    MsgEmpty = "This field can't be empty!"
    
End Sub

Private Sub Espanol_ES()
    On Error Resume Next
    
    ' Fonction ajout� le 24/10/16 pour r�cup�rer la version automatiquement depuis le GMS
    Me.Caption = "Dimensi�n Autom�tica | ES " & APP_VERSION
        
    FrameParameters.Caption = "Par�metros"
    FrameTextType.Caption = "Tipo de Texto"
    LabelUnits.Caption = "Unidades"
    
    UrlCorelNaVeia.ControlTipText = "Acceder a otras macros en CorelNaVeia.com..."
    UrlKafard.ControlTipText = "Acceda al Site de Fabrice para obtener m�s macros..."
    UrlFerreira.ControlTipText = "Contactar a Ferreira Felipe por WhatsApp para otras macros..."
    UrlFabricePayPal.ControlTipText = "P�gale un caf� a Fabrice..."
    UnitsList.ControlTipText = "Elija las Unidades de Medici�n"
    DimensionDistance.ControlTipText = "Distancia de Dimensi�n"
    DimensionTextSize.ControlTipText = "Fuente de Dimensi�n"
    Quantity.ControlTipText = "Incluir cantidades y detalles"
    QuantityPrefix = "Cantidad: "
    DimensionLineThickness.ControlTipText = "L�nea de Dimensi�n"
    ScaleList.ControlTipText = "Seleccionar/Incluir Escala"
    
    LabelScale.Caption = "Escalas"
    OptionScale.Caption = "Incluir Escala"
    OptionAjustText.Caption = "Ajustar Leyenda"
    OptionUnits.Caption = "Incluir Leyenda"
    LabelQuantity.Caption = "Cuantitativo"
    OptionTechnic.Caption = "T�cnico"
    OptionSimple.Caption = "Sencillo"
    OptionGroupDimensions.Caption = "Agrupar Dimensi�ns"
    OptionSaveScale.Caption = "�Guardar La Escala Actual!"
    ButtonApply.Caption = "APLICAR"
    ButtonReset.Caption = "ELIMINAR"
    
    UnitsList.AddItem "Mil�metros"
    UnitsList.AddItem "Cent�metros"
    UnitsList.AddItem "Metros"
    UnitsList.AddItem "Pulgadas"
    UnitsList.AddItem "P�xeles"
    UnitsList.ListIndex = 0
    
    ScaleText = "Escala 1/"
    CurrentScaleText = "(Actual) 1/"
    
    ScaleList.AddItem CurrentScaleText & ActiveDocument.WorldScale, 0
    ScaleList.AddItem "1/1"
    ScaleList.AddItem "1/2"
    ScaleList.AddItem "1/4"
    ScaleList.AddItem "1/5"
    ScaleList.AddItem "1/10"
    ScaleList.AddItem "1/15"
    ScaleList.AddItem "1/20"
    ScaleList.AddItem "1/40"
    ScaleList.AddItem "1/50"
    ScaleList.ListIndex = 0
    
    'Notifications
    MsgNoObject = "�Seleccione una o m�s formas por adelantado!"
    MsgOnlyNumbers = "�Introduce solo n�meros!"
    MsgEmpty = "�Este campo no puede estar vac�o!"
    
End Sub

Private Sub Russian_RU()
    On Error Resume Next
    
    ' Fonction ajout� le 24/10/16 pour r�cup�rer la version automatiquement depuis le GMS
    Me.Caption = "����������� | RU " & APP_VERSION
        
    FrameParameters.Caption = "���������"
    FrameTextType.Caption = "��� ��������"
    LabelUnits.Caption = "�������"
    
    UrlCorelNaVeia.ControlTipText = "������ ������� �� CorelNaVeia.com ..."
    UrlKafard.ControlTipText = "������� �� ���� Fabrice ����� ���������� ������ �������..."
    UrlFerreira.ControlTipText = "�������� Ferreira Felipe �� WhatsApp �� ������ ��������..."
    UrlFabricePayPal.ControlTipText = "�������� Fabrice ����..."
    UnitsList.ControlTipText = "�������� ������� ���������"
    DimensionDistance.ControlTipText = "���������� �� ��������� �����"
    DimensionTextSize.ControlTipText = "������ ������"
    Quantity.ControlTipText = "�������� �������� ����������"
    QuantityPrefix = "����������: "
    DimensionLineThickness.ControlTipText = "������� ��������� �����"
    ScaleList.ControlTipText = "������� �������"
    
    LabelScale.Caption = "�������"
    OptionScale.Caption = "�������� �������"
    OptionAjustText.Caption = "�������� �����"
    OptionUnits.Caption = "�������� �������"
    LabelQuantity.Caption = "����������"
    OptionTechnic.Caption = "�����."
    OptionSimple.Caption = "��������"
    OptionGroupDimensions.Caption = "������������ �������"
    OptionSaveScale.Caption = "��������� �������"
    ButtonApply.Caption = "���������"
    ButtonReset.Caption = "�����"
    
    UnitsList.AddItem "����������"
    UnitsList.AddItem "����������"
    UnitsList.AddItem "�����"
    UnitsList.AddItem "�����"
    UnitsList.AddItem "�������"
    UnitsList.ListIndex = 0
    
    ScaleText = "������� 1/"
    CurrentScaleText = "(�������) 1/"
    
    ScaleList.AddItem CurrentScaleText & ActiveDocument.WorldScale, 0
    ScaleList.AddItem "1/1"
    ScaleList.AddItem "1/2"
    ScaleList.AddItem "1/4"
    ScaleList.AddItem "1/5"
    ScaleList.AddItem "1/10"
    ScaleList.AddItem "1/15"
    ScaleList.AddItem "1/20"
    ScaleList.AddItem "1/40"
    ScaleList.AddItem "1/50"
    ScaleList.ListIndex = 0
    
    'Notifications
    MsgNoObject = "�������� ���� ��� ��������� ��������"
    MsgOnlyNumbers = "����������� ������ �����"
    MsgEmpty = "���� �� ����� ���� ������"
    
End Sub
