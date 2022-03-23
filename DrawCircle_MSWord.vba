Function DrawCircle()
    ' Include Shape
    With ActiveDocument.Shapes.AddShape(Type:=msoShapeOval, _
                                      Left:=InchesToPoints(0.1), _
                                      Top:=InchesToPoints(0.1), _
                                      Width:=InchesToPoints(0.5), _
                                      Height:=InchesToPoints(0.5))
        ' Include and format Text
        With .TextFrame.TextRange
            .Text = "My Text"
            .Font.ColorIndex = 12
        End With
        
        ' Format Shape
        .Fill.Visible = msoFalse
        With .Line
            .Weight = 1.75
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0#
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .BackColor.RGB = RGB(255, 255, 255)
        End With
    End With
    
End Function
