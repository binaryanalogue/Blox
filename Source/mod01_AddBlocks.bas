Attribute VB_Name = "mod01_AddBlocks"
Option Explicit

' Blox add-in for PowerPoint
' Created 2022 by @binaryanalogue
' Free to use and distribute

' 2022-06-05 v1.0 Released
'
' 2022-06-06 v1.1 Released
' - sets BlackWhiteMode to white, preventing display of borders
'

Sub AddWhiteBlock()
    
    '
    ' Creates a white-filled text frame in the middle of the slide
    '
    
    Call CreateBlock(False)
    
End Sub

Sub AddSemiBlock()
    
    '
    ' Creates a semi-transparent white text frame in the middle of the slide
    '
    
    Call CreateBlock(True)
    
End Sub

Private Sub CreateBlock(AddSemiTransparency As Boolean)

    '
    ' Creates a new box/block
    ' If user has a shape selected, creates a box scaled and centered on that shape
    ' If no shape selected, creates a box in the centre of the slide
    '

    ' New shape
    Dim shpNew As Shape
    
    ' Selected shape
    Dim blnSelExists As Boolean
    Dim shpSel
    
    ' Location and sizing
    Dim lLeft As Long
    Dim lTop As Long
    Dim lWidth As Long
    Dim lHeight As Long

    
    ' Terminate if there are no presentations open
    If Application.Windows.Count = 0 Then Exit Sub

    ' Do not try to add if user isn't on a valid slide
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then Exit Sub
    
    ' Select the presentation if user is anywhere but on the slide
    If ActiveWindow.Panes.Count > 1 Then ActiveWindow.Panes(2).Activate
    
    ' Check if user has something already selected
    Select Case ActiveWindow.Selection.Type
              
        Case 2
            
            ' Existing shape selected
            
            blnSelExists = True
            Set shpSel = ActiveWindow.Selection.ShapeRange(1)
            
            ' Scale to selected shape with some thresholds
            Select Case shpSel.Height
            
                Case Is < 20
                    lWidth = 10
                    lHeight = 10
                    
                Case Is > 400
                    lWidth = 40
                    lHeight = 40
                    
                Case Else
                    lWidth = shpSel.Width / 3
                    lHeight = shpSel.Height / 3
                    
            End Select
            
            
            ' If transparency is requested, assume this should be a text box; make wider
            'f AddSemiTransparency Then lWidth = lWidth * 2
            
            ' Centre in slide
            lLeft = shpSel.Left + shpSel.Width / 2 - lWidth / 2
            lTop = shpSel.Top + shpSel.Height / 2 - lHeight / 2
            
            ' Something is selected; confine new shape to selected shape's properties
            Set shpNew = ActivePresentation.Slides _
                    (ActiveWindow.Selection.SlideRange(1).SlideIndex) _
                    .Shapes.AddTextbox( _
                        Orientation:=msoTextOrientationHorizontal, _
                        Left:=lLeft, _
                        Top:=lTop, _
                        Width:=lWidth, _
                        Height:=lHeight)
                        
        Case Else
        
            ' Nothing or something other than a shape is selected
            
            ' Square-shaped
            lWidth = 40
            lHeight = 40
            
            ' Centre in slide
            lLeft = Application.ActivePresentation.PageSetup.SlideWidth / 2 - lWidth / 2
            lTop = Application.ActivePresentation.PageSetup.SlideHeight / 2 - lHeight / 2
            
            ' If transparency is requested, assume this should be a text box; make wider
            'If AddSemiTransparency Then lWidth = lWidth * 2
            
            ' Make a new box in the center of the slide
            Set shpNew = ActivePresentation.Slides _
                    (ActiveWindow.Selection.SlideRange(1).SlideIndex) _
                    .Shapes.AddTextbox( _
                        Orientation:=msoTextOrientationHorizontal, _
                        Left:=lLeft, _
                        Top:=lTop, _
                        Width:=lWidth, _
                        Height:=lHeight)
            
    End Select

    With shpNew
    
        .LockAspectRatio = False
        
        ' Reset height/width
        .Height = lHeight
        .Width = lWidth

        ' Set fill parameters
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        If AddSemiTransparency Then
            .Fill.Transparency = 0.5
        Else
            .Fill.Solid
            .Fill.Transparency = 0#
        End If
                                        
        ' Show as white in black/white mode
        .BlackWhiteMode = msoBlackWhiteWhite
        
        ' Set border parameters
        .Line.Weight = 0#
        .Line.Visible = msoFalse
        .Line.ForeColor.RGB = RGB(255, 255, 255)
        .Line.BackColor.RGB = RGB(255, 255, 255)
    
        ' Set text & font size
        With .TextFrame
            
            .WordWrap = msoTrue
            .HorizontalAnchor = msoAnchorCenter
            .VerticalAnchor = msoAnchorMiddle
            .AutoSize = ppAutoSizeNone
            
            With .Ruler.Levels(1)
                .FirstMargin = 0
                .LeftMargin = 0
            End With
        
            With .TextRange
            
                With .Font
                    .Size = 10
                    .Bold = False
                    .Underline = False
                    .Color.RGB = RGB(0, 0, 0)
                End With
                
                With .ParagraphFormat
                    .Alignment = ppAlignCenter
                    .LineRuleWithin = msoTrue
                    .SpaceWithin = 1
                    .LineRuleBefore = msoTrue
                    .SpaceBefore = 0.25
                    .LineRuleAfter = msoTrue
                    .SpaceAfter = 0.25
                End With
                
                ' Enter text range if semi-transparent; user probably wants to add text
                If AddSemiTransparency Then .Select
                
            End With 'TextRange
            
            .MarginTop = 3.5
            .MarginBottom = 3.5
            .MarginLeft = 3.5
            .MarginRight = 3.5
            
        End With 'TextFrame
        
        ' Buggy sizing in PowerPoint
        .Height = lHeight
        .Top = lTop
        .Left = lLeft
        
        ' Leave object selected
        If Not AddSemiTransparency Then .Select
        
    End With ' New shape
    
End Sub
