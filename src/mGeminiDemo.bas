Attribute VB_Name = "mGeminiDemo"
' ===================================================================
' Module: mGeminiDemo
' Purpose: Creates demo worksheets to showcase Gemini AI capabilities
' Usage: Run CreateGeminiDemoSheet from the VBA Editor or Alt+F8
' ===================================================================

Option Explicit

Sub CreateGeminiDemoSheet()
    ' Creates a demo sheet with practical business examples
    ' to demonstrate different Gemini input modes and use cases
    
    Dim ws As Worksheet
    Dim wsName As String
    wsName = "GEMINI_DEMO"
    
    ' Delete existing demo sheet if exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ActiveWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new demo sheet
    Set ws = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.count))
    ws.Name = wsName
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 40
    
    ' =========== HEADER ===========
    ws.Range("A1").Value = "GEMINI AI DEMO"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1:E1").Merge
    ws.Range("A1:E1").Interior.Color = RGB(66, 133, 244) ' Google Blue
    ws.Range("A1:E1").Font.Color = RGB(255, 255, 255)
    
    ws.Range("A2").Value = "Select data below and run the Gemini macro (Alt+F8 > Gemini)"
    ws.Range("A2:E2").Merge
    ws.Range("A2").Font.Italic = True
    
    ' =========== EXAMPLE 1: Sales Data ===========
    ws.Range("A4").Value = "EXAMPLE 1: Sales Analysis"
    ws.Range("A4").Font.Bold = True
    ws.Range("A4:E4").Interior.Color = RGB(234, 234, 234)
    
    ws.Range("A5").Value = "Product"
    ws.Range("B5").Value = "Q1 Sales"
    ws.Range("C5").Value = "Q2 Sales"
    ws.Range("D5").Value = "Q3 Sales"
    ws.Range("A5:D5").Font.Bold = True
    
    ws.Range("A6").Value = "Laptop Pro"
    ws.Range("B6").Value = 45000
    ws.Range("C6").Value = 52000
    ws.Range("D6").Value = 48000
    
    ws.Range("A7").Value = "Wireless Mouse"
    ws.Range("B7").Value = 12000
    ws.Range("C7").Value = 15000
    ws.Range("D7").Value = 18000
    
    ws.Range("A8").Value = "USB-C Hub"
    ws.Range("B8").Value = 8000
    ws.Range("C8").Value = 7500
    ws.Range("D8").Value = 9200
    
    ws.Range("A9").Value = "Monitor 27"""
    ws.Range("B9").Value = 32000
    ws.Range("C9").Value = 28000
    ws.Range("D9").Value = 35000
    
    ws.Range("B6:D9").NumberFormat = "$#,##0"
    
    ws.Range("E5").Value = "Try: Select A5:D9, run Gemini, type:"
    ws.Range("E6").Value = "- 'Analyze sales trends'"
    ws.Range("E7").Value = "- 'Which product is growing fastest?'"
    ws.Range("E8").Value = "- 'Create a summary report'"
    ws.Range("E5:E8").Font.Color = RGB(100, 100, 100)
    
    ' =========== EXAMPLE 2: Customer Feedback ===========
    ws.Range("A11").Value = "EXAMPLE 2: Customer Feedback"
    ws.Range("A11").Font.Bold = True
    ws.Range("A11:E11").Interior.Color = RGB(234, 234, 234)
    
    ws.Range("A12").Value = "Customer"
    ws.Range("B12").Value = "Rating"
    ws.Range("C12:D12").Merge
    ws.Range("C12").Value = "Feedback"
    ws.Range("A12:D12").Font.Bold = True
    
    ws.Range("A13").Value = "John D."
    ws.Range("B13").Value = 5
    ws.Range("C13:D13").Merge
    ws.Range("C13").Value = "Excellent product, fast shipping!"
    
    ws.Range("A14").Value = "Sarah M."
    ws.Range("B14").Value = 3
    ws.Range("C14:D14").Merge
    ws.Range("C14").Value = "Good quality but took too long to arrive"
    
    ws.Range("A15").Value = "Mike R."
    ws.Range("B15").Value = 4
    ws.Range("C15:D15").Merge
    ws.Range("C15").Value = "Works great, packaging could be better"
    
    ws.Range("A16").Value = "Lisa K."
    ws.Range("B16").Value = 2
    ws.Range("C16:D16").Merge
    ws.Range("C16").Value = "Product arrived damaged, waiting for replacement"
    
    ws.Range("E12").Value = "Try: Select A12:D16, run Gemini, type:"
    ws.Range("E13").Value = "- 'Summarize customer sentiment'"
    ws.Range("E14").Value = "- 'What are the main complaints?'"
    ws.Range("E15").Value = "- 'Suggest improvements'"
    ws.Range("E12:E15").Font.Color = RGB(100, 100, 100)
    
    ' =========== EXAMPLE 3: Data Transformation ===========
    ws.Range("A18").Value = "EXAMPLE 3: Data Tasks"
    ws.Range("A18").Font.Bold = True
    ws.Range("A18:E18").Interior.Color = RGB(234, 234, 234)
    
    ws.Range("A19").Value = "Raw Data"
    ws.Range("A19").Font.Bold = True
    
    ws.Range("A20").Value = "john.doe@email.com"
    ws.Range("A21").Value = "jane.smith@company.org"
    ws.Range("A22").Value = "bob.wilson@test.net"
    ws.Range("A23").Value = "alice.jones@example.com"
    
    ws.Range("E19").Value = "Try: Select A20:A23, run Gemini, type:"
    ws.Range("E20").Value = "- 'Extract first names'"
    ws.Range("E21").Value = "- 'Convert to table format'"
    ws.Range("E22").Value = "- 'Create Excel formula to extract domain'"
    ws.Range("E19:E22").Font.Color = RGB(100, 100, 100)
    
    ' =========== EXAMPLE 4: AskGemini Function ===========
    ws.Range("A25").Value = "EXAMPLE 4: AskGemini Formula"
    ws.Range("A25").Font.Bold = True
    ws.Range("A25:E25").Interior.Color = RGB(234, 234, 234)
    
    ws.Range("A26").Value = "Prompt"
    ws.Range("B26:D26").Merge
    ws.Range("B26").Value = "Result (using =AskGemini(A#))"
    ws.Range("A26:D26").Font.Bold = True
    
    ws.Range("A27").Value = "What is 2+2?"
    ws.Range("B27:D27").Merge
    ws.Range("B27").Value = "'=AskGemini(A27)"
    ws.Range("B27").Font.Italic = True
    ws.Range("B27").Font.Color = RGB(100, 100, 100)
    
    ws.Range("A28").Value = "Capital of France?"
    ws.Range("B28:D28").Merge
    ws.Range("B28").Value = "'=AskGemini(A28)"
    ws.Range("B28").Font.Italic = True
    ws.Range("B28").Font.Color = RGB(100, 100, 100)
    
    ws.Range("A29").Value = "Translate 'Hello' to Spanish"
    ws.Range("B29:D29").Merge
    ws.Range("B29").Value = "'=AskGemini(A29)"
    ws.Range("B29").Font.Italic = True
    ws.Range("B29").Font.Color = RGB(100, 100, 100)
    
    ws.Range("E26").Value = "Use AskGemini as a formula:"
    ws.Range("E27").Value = "- =AskGemini(""Your question"")"
    ws.Range("E28").Value = "- =AskGemini(A1)"
    ws.Range("E29").Value = "- =AskGemini(""Translate: "" & A1)"
    ws.Range("E26:E29").Font.Color = RGB(100, 100, 100)
    
    ' =========== EXAMPLE 5: Quick Questions ===========
    ws.Range("A31").Value = "EXAMPLE 5: Quick Questions (no selection needed)"
    ws.Range("A31").Font.Bold = True
    ws.Range("A31:E31").Interior.Color = RGB(234, 234, 234)
    
    ws.Range("A32:D35").Merge
    ws.Range("A32").Value = "For quick questions, just run Gemini without selecting any data. " & _
                            "The input box will appear where you can type any question."
    ws.Range("A32").WrapText = True
    ws.Range("A32").VerticalAlignment = xlTop
    
    ws.Range("E32").Value = "Try running Gemini with no selection:"
    ws.Range("E33").Value = "- 'Write a VLOOKUP formula example'"
    ws.Range("E34").Value = "- 'How do I create a pivot table?'"
    ws.Range("E35").Value = "- 'Explain SUMIFS function'"
    ws.Range("E32:E35").Font.Color = RGB(100, 100, 100)
    
    ' =========== FOOTER ===========
    ws.Range("A37").Value = "TIP: Change GEMINI_INPUT_MODE in mGemini module to control input behavior"
    ws.Range("A37").Font.Italic = True
    ws.Range("A37").Font.Color = RGB(100, 100, 100)
    
    ' Turn off gridlines for cleaner look
    ActiveWindow.DisplayGridlines = False
    
    ' Add light borders to all data areas
    AddBorders ws.Range("A5:D9")
    AddBorders ws.Range("A12:D16")
    AddBorders ws.Range("A19:A23")
    AddBorders ws.Range("A26:D29")
    
    ' Add outside border to tip sections
    AddLightBorder ws.Range("E5:E9")
    AddLightBorder ws.Range("E12:E16")
    AddLightBorder ws.Range("E19:E22")
    AddLightBorder ws.Range("E26:E29")
    AddLightBorder ws.Range("E32:E35")
    
    ' Activate the demo sheet
    ws.Activate
    ws.Range("A1").Select
    
    MsgBox "Demo sheet created!" & vbCrLf & vbCrLf & _
           "Try these steps:" & vbCrLf & _
           "1. Select a data range (e.g., A5:D9)" & vbCrLf & _
           "2. Press Alt+F8, select 'Gemini', click Run" & vbCrLf & _
           "3. Type an instruction like 'Analyze this data'" & vbCrLf & _
           "4. Check GEMINI_OUTPUT sheet for results", _
           vbInformation, "Gemini Demo Ready"
End Sub

Private Sub AddBorders(rng As Range)
    ' Adds light gray borders to all cells in range
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(200, 200, 200)
    End With
End Sub

Private Sub AddLightBorder(rng As Range)
    ' Adds light outside border only (for tip boxes)
    With rng
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(220, 220, 220)
        .Interior.Color = RGB(250, 250, 250) ' Very light gray background
    End With
End Sub




