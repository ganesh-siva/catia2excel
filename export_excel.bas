Attribute VB_Name = "export_excel"
Sub CATMain()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim para As Parameters
Set para = part1.Parameters

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("PartBody")

Dim hsp As HybridShapes
Set hsp = body1.HybridShapes

Dim xlapp As Excel.Application

'getting excel as an object
Set xlapp = GetObject(, "Excel.Application")
xlapp.Application.Visible = True

'setting column headers
xlapp.Application.ActiveCell(1, 1) = "Name"
xlapp.Application.ActiveCell(1, 2) = "X"
xlapp.Application.ActiveCell(1, 3) = "Y"
xlapp.Application.ActiveCell(1, 4) = "Z"
xlapp.Application.Range("A1:D1").Font.Bold = True

'iterating over all points in PartBody &  writing out name and coordinates as a record
Dim i As Integer
For i = 1 To hsp.Count
        xlapp.Application.ActiveCell(i + 1, 1) = hsp.Item(i).Name
        xlapp.Application.ActiveCell(i + 1, 2) = part1.Parameters.Item("PartBody\" & hsp.Item(i).Name & "\X").ValueAsString
        xlapp.Application.ActiveCell(i + 1, 3) = part1.Parameters.Item("PartBody\" & hsp.Item(i).Name & "\Y").ValueAsString
        xlapp.Application.ActiveCell(i + 1, 4) = part1.Parameters.Item("PartBody\" & hsp.Item(i).Name & "\Z").ValueAsString
Next

'Test MsgBox part1.Parameters.Item("PartBody\Point.1\X").ValueAsString
End Sub


