Attribute VB_Name = "NewMacros"
Sub Macro1()
    Dim sourcePath As String
    Dim sourceName As String
    Dim pdfName As String
    
    ' Get the path of the active document
    sourcePath = ActiveDocument.Path
    
    ' Get the name of the active document without extension
    sourceName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    
    ' Create the PDF name by adding .pdf extension
    pdfName = sourcePath & "\" & sourceName & ".pdf"
    
    ' Export the document as PDF
    ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfName, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
        wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=1, To:=1, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    
    ' Change file open directory to the source document's location
    ChangeFileOpenDirectory sourcePath
End Sub


Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
    Selection.EscapeKey
    Selection.EscapeKey
    Selection.EscapeKey
    Selection.EscapeKey
    Selection.EscapeKey
    ActiveWindow.Close
End Sub
