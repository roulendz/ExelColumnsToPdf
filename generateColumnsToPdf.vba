Sub GenerateBusinessCards()
    'Declare variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim card As Range
    Dim pdfName As String
    Dim pdfPath As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    'Set workbook and worksheet objects
    Set wb = ThisWorkbook
    Set ws = wb.Sheets("Peackoc gel polish") 'Change sheet name as needed
    
    'Set range object for the data in column A
    Set rng = ws.Range("A1", ws.Range("A1").End(xlDown))
    
    'Set range object for the business card template in column F
    Set card = ws.Range("F1")
    
    'Add a new sheet for the business cards
    Sheets.Add After:=Sheets(Sheets.Count)
    
    'Set the column width and row height to fit the card size
    Columns("A:Z").ColumnWidth = 10
    Rows("1:100").RowHeight = 50
    
    'Loop through each cell in the range
    For i = 1 To rng.Count
        
        'Copy the template to the new sheet
        card.Copy
        
        'Paste the template in a cell based on the loop counter
        If i Mod 10 = 0 Then 'If i is divisible by 10, paste in column J
            k = i / 10 'k is the row number
            Cells(k, 10).PasteSpecial xlPasteAll 'Paste in cell Jk
        Else 'If i is not divisible by 10, paste in column A to I
            j = i Mod 10 'j is the column number
            k = (i - j) / 10 + 1 'k is the row number
            Cells(k, j).PasteSpecial xlPasteAll 'Paste in cell jk
        End If
        
        'Replace the placeholder with the data from column A
        With Cells(k, j)
            .Replace What:="[Name]", Replacement:=rng.Cells(i, 1).Value, LookAt:=xlPart, MatchCase:=False
        End With
        
    Next i
    
    'Set the pdf file name and path
    pdfName = "BusinessCards.pdf"
    pdfPath = "C:\Users\rolan\Documents\" & pdfName 'Change folder path as needed
    
    'Set the page margins to 5mm (14.17 points)
    ActiveSheet.PageSetup.LeftMargin = 14.17
    ActiveSheet.PageSetup.RightMargin = 14.17
    ActiveSheet.PageSetup.TopMargin = 14.17
    ActiveSheet.PageSetup.BottomMargin = 14.17
    
    'Export the sheet as a pdf file
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
    
    'Delete the sheet
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
    'Inform the user that the process is completed
    MsgBox "Business cards have been generated and saved as one pdf file in the specified folder.", vbInformation
    
End Sub

