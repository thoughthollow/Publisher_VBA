' Add to the View tab in the ribbon as a custom group for easier zoom level adjustment

Sub zoom_in()
 ActiveDocument.ActiveView.Zoom = ActiveDocument.ActiveView.Zoom + 10
End Sub

Sub zoom_out()
 ActiveDocument.ActiveView.Zoom = ActiveDocument.ActiveView.Zoom - 10
End Sub
