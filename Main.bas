Attribute VB_Name = "Main"
Sub PercentBar(Shape As Control, Done As Long, Total As Long)

'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)

On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
x = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(192, 192, 192), BF
Shape.Line (0, 0)-(x - 10, Shape.Height), RGB(0, 0, 127), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(192, 192, 192)
'Shape.Print Percent(Done, Total, 100) & "%"
End Sub


