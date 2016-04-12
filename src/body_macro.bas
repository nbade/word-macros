Sub Macro1()
'
' Macro1 Macro
'
'
    Selection.TypeText Text:="typing some things..."
    ActiveWindow.View.ReadingLayout = Not ActiveWindow.View.ReadingLayout
    Application.WindowState = wdWindowStateMinimize
    Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    Selection.TypeText Text:="typing things "
    Selection.Font.Name = "Adobe Caslon Pro"
    Selection.TypeText Text:="and more things."
End Sub
