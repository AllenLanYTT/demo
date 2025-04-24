Sub GeneratePDFsFromTemplate()
    Dim ws As Worksheet
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim templatePath As String
    Dim outputFolder As String
    Dim lastRow As Long
    Dim i As Long
    Dim name As String, congregation As String, partNum As String
    Dim pdfFileName As String
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Assignment")
    
    ' Define the Word template path
    templatePath = ThisWorkbook.Path & "\template_2.docx" ' Update with your Word template file name
    
    ' Define the output folder for PDFs
    outputFolder = ThisWorkbook.Path & "\Generated_PDFs\"
    If Dir(outputFolder, vbDirectory) = "" Then
        MkDir outputFolder ' Create the folder if it doesn't exist
    End If
    
    ' Get the last row with data in the "Assignment" sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize Word application
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = False ' Set to True for debugging
    
    ' Loop through each row in the "Assignment" sheet
    For i = 2 To lastRow ' Assuming row 1 contains headers
        ' Read data from the sheet
        name = ws.Cells(i, 1).Value ' Column A: NAME
        congregation = ws.Cells(i, 2).Value ' Column B: CONGREGATION
        partNum = ws.Cells(i, 3).Value ' Column C: PART_NUM
        
        ' Open the Word template
        Set WordDoc = WordApp.Documents.Open(templatePath)
        
        ' Replace placeholders in the Word template
        With WordDoc.Content.Find
            .Text = "<<NAME>>"
            .Replacement.Text = name
            .Execute Replace:=2
            
            .Text = "<<CONGREGATION>>"
            .Replacement.Text = congregation
            .Execute Replace:=2
            
            .Text = "<<PART_NUM>>"
            .Replacement.Text = partNum
            .Execute Replace:=2
            ' Replace:=2 (or wdReplaceAll): Replace all occurrences of the found text.
            
            If .Found Then
                ReplaceStatus = 0 ' Success
            Else
                ReplaceStatus = 1 ' Failure
            End If
                
         End With
    

        ' Define the PDF file name
        pdfFileName = outputFolder & name & "_" & partNum & ".pdf"
        
        ' Save the Word document as a PDF
        WordDoc.ExportAsFixedFormat OutputFileName:=pdfFileName, ExportFormat:=17 ' 17 = wdExportFormatPDF
        
        ' Close the Word document without saving changes
        WordDoc.Close SaveChanges:=False
 
    Next i
    
    ' Quit Word application
    WordApp.Quit
    
    ' Release objects
    Set WordDoc = Nothing
    Set WordApp = Nothing
    
    ' Notify the user
    MsgBox "PDF generation completed! Files are saved in: " & outputFolder, vbInformation
End Sub
