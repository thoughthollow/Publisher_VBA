' Change selected text to upper or lower case
Sub Uppercase()

    Selection.TextRange.text = UCase(Selection.TextRange.text)

End Sub
Sub Lowercase()

    Selection.TextRange.text = LCase(Selection.TextRange.text)

End Sub
