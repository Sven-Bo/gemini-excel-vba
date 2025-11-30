Attribute VB_Name = "mGemini"
' ===================================================================
' Project Name: Google Gemini AI in Excel VBA
' Author: Sven Bosau
' Website: https://pythonandvba.com
' YouTube: https://youtube.com/@codingisfun
' Date Created: 2025/11/29
' Last Modified: 2025/11/29
' Version: 1.0
' ===================================================================
'
' Description:
' This VBA module enables users to interact with Google's Gemini AI
' models directly from Excel by sending text prompts via the Gemini API.
' It allows capturing AI model outputs and displaying them in Excel.
'
' ===================================================================
'
' DISCLAIMER:
' This code is distributed "as is" and the author makes no representations
' or warranties, express or implied, regarding the functionality, operability,
' or use of the code, including, without limitation, any implied warranties of
' merchantability or fitness for a particular purpose. The user of this code
' assumes the entire risk as to its quality and performance. Should any part
' of the code prove defective, the user assumes the entire cost of all necessary
' servicing or repair.
'
' The user must comply with all applicable local laws and regulations in
' using the code, including, without limitation, all intellectual property laws.
'
' Furthermore, by using this code, the user acknowledges and agrees that
' they have read and understand Google's Gemini API Terms of Service and
' agree to abide by them.
' Google's Terms of Service: https://ai.google.dev/terms
'
' The API key is confidential and should be kept secure. Sharing or exposing
' the API key is strictly prohibited. Use the API key responsibly and ensure
' it is stored, transmitted, and used securely.
'
' ===================================================================
'
' AVAILABLE MODELS (as of November 2025):
' Source: https://ai.google.dev/gemini-api/docs/models
'
' GEMINI 3 SERIES (Most Powerful - Thinking Models):
'   - gemini-3-pro-preview        : Best multimodal understanding, 1M input, 64k output
'                                   Uses thinkingLevel: "low" or "high" (cannot disable)
'
' GEMINI 2.5 SERIES (Best Price-Performance):
'   - gemini-2.5-flash            : STABLE - Best balance of speed/quality, 1M input, 64k output
'   - gemini-2.5-flash-lite       : STABLE - Fastest, most cost-efficient, 1M input, 64k output
'                                   Uses thinkingBudget: 0 to disable, or token count
'
' GEMINI 2.0 SERIES (Previous Generation):
'   - gemini-2.0-flash            : Fast and capable, good for most tasks
'   - gemini-2.0-flash-lite       : Ultra-fast, basic tasks
'
' RECOMMENDATION:
'   - For Excel assistant tasks (formulas, quick answers): gemini-2.5-flash-lite
'   - For balanced quality/speed: gemini-2.5-flash
'   - For complex reasoning/analysis: gemini-3-pro-preview
'
' ===================================================================
'
' HOW TO USE:
'
'   1. Set your API key in GEMINI_API_KEY constant below
'   2. Run the "Gemini" macro (Alt+F8 > Gemini > Run)
'   3. Enter your prompt in the input box (or select cells first)
'   4. Response appears in "GEMINI_OUTPUT" sheet
'
' CONFIGURATION:
'   - Change GEMINI_MODEL to switch between models
'   - Change GEMINI_THINKING_MODE to control reasoning depth
'   - Change GEMINI_SYSTEM_INSTRUCTION to customize AI behavior
'   - Change GEMINI_MAX_TOKENS for longer/shorter responses
'
' ===================================================================

Option Explicit

' =========================== CONFIGURATION ===========================

' API Key - Get yours from: https://aistudio.google.com/apikey
' IMPORTANT: Keep this key secure and never share it publicly!
Const GEMINI_API_KEY As String = "YOUR_API_KEY" '<<<<< CHANGE ME !!!

' Model Selection (see AVAILABLE MODELS above)
' Options: "gemini-3-pro-preview", "gemini-2.5-flash", "gemini-2.5-flash-lite"
Const GEMINI_MODEL As String = "gemini-2.5-flash"

' API Version: Use "v1beta" for latest features, "v1" for stable
Const GEMINI_API_VERSION As String = "v1beta"

' Max output tokens (models support up to 65,536)
Const GEMINI_MAX_TOKENS As Long = 8192

' Temperature: Controls randomness (0.0 = deterministic, 2.0 = very random)
' Recommended: 1.0 for general use, 0.0-0.3 for factual, 0.8+ for creative
Const GEMINI_TEMPERATURE As Double = 1#

' Sampling parameters
Const GEMINI_TOP_P As Double = 0.95
Const GEMINI_TOP_K As Long = 40

' Thinking Configuration:
' For Gemini 3: Use thinkingLevel = "low" or "high" (cannot be disabled)
' For Gemini 2.5: Use thinkingBudget = 0 to disable, or positive number for token budget
'
' GEMINI_THINKING_MODE options:
'   "level_high"  - Gemini 3 only: Maximum reasoning depth
'   "level_low"   - Gemini 3 only: Minimal latency
'   "budget_0"    - Gemini 2.5 only: Disable thinking (fastest)
'   "budget_auto" - Gemini 2.5 only: Let model decide (default)
'   ""            - No thinking config (use model defaults)
Const GEMINI_THINKING_MODE As String = ""

' System instruction to set assistant behavior
Const GEMINI_SYSTEM_INSTRUCTION As String = "You are a helpful assistant"

' Prompt Input Mode - Controls how the prompt is collected
' Options:
'   "selection"  - Use selected cells only (no input box)
'   "inputbox"   - Always show input box (ignore selection)
'   "both"       - Use selection as context + show input box for additional prompt
'   "auto"       - If selection has text use it, otherwise show input box
'
' EXAMPLES:
'   "selection" : Select cells A1:A3 with data, run macro -> sends cell contents to AI
'   "inputbox"  : Run macro -> type "What is 2+2?" -> sends typed text to AI
'   "both"      : Select cells with sales data, run macro -> type "Summarize this"
'                 -> sends: "Summarize this" + cell contents to AI
'   "auto"      : If cells selected have text, use them; if empty, show input box
Const GEMINI_INPUT_MODE As String = "both"

' Output Mode - Controls how the response is displayed
' Options:
'   "lines"     - Each line in a separate cell (A1, A2, A3, etc.)
'   "single"    - All text in one cell (A1) with line breaks preserved
'
' EXAMPLES:
'   "lines"  : Response "Hello\nWorld" -> A1="Hello", A2="World"
'   "single" : Response "Hello\nWorld" -> A1="Hello[newline]World"
Const GEMINI_OUTPUT_MODE As String = "lines"

' ====================== END CONFIGURATION ============================

#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
        (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#End If

Sub Gemini()
    On Error GoTo ErrorHandler
    
    #If Mac Then
        MsgBox "This module is designed for Windows only and is not compatible with macOS.", _
              vbOKOnly, "Windows Only"
        Exit Sub
    #End If
    
    Dim HasInternet As Boolean
    HasInternet = GetInternetConnectedState()
    If Not HasInternet Then
        MsgBox "Internet connection is required. Please connect and try again.", _
              vbOKOnly Or vbInformation, "No Internet"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim prompt As String
    Dim selectionText As String
    Dim userInput As String
    
    prompt = ""
    selectionText = ""
    userInput = ""
    
    ' Get text from selected cells (preserving table structure)
    If TypeName(Selection) = "Range" Then
        Dim cell As Range
        Dim currentRow As Long
        currentRow = 0
        
        For Each cell In Selection
            If currentRow = 0 Then currentRow = cell.Row
            
            ' Add line break when moving to new row
            If cell.Row <> currentRow Then
                selectionText = selectionText & vbLf
                currentRow = cell.Row
            ElseIf selectionText <> "" Then
                selectionText = selectionText & vbTab
            End If
            
            selectionText = selectionText & CStr(cell.Value)
        Next cell
        selectionText = Trim(selectionText)
    End If
    
    ' Build prompt based on input mode setting
    Select Case LCase(GEMINI_INPUT_MODE)
        Case "selection"
            ' Use selected cells only
            If selectionText = "" Then
                MsgBox "Please select cells containing your prompt.", vbExclamation, "No Selection"
                Application.ScreenUpdating = True
                Exit Sub
            End If
            prompt = selectionText
            
        Case "inputbox"
            ' Always show input box, ignore selection
            userInput = InputBox("Enter your prompt for Gemini:", "Gemini AI", "")
            If Trim(userInput) = "" Then
                Application.ScreenUpdating = True
                Exit Sub
            End If
            prompt = userInput
            
        Case "both"
            ' Use selection as context + input box for instruction
            If selectionText <> "" Then
                userInput = InputBox("Selected data will be sent as context." & vbCrLf & vbCrLf & _
                                    "Enter your instruction (e.g., 'Summarize this', 'Translate to Spanish'):", _
                                    "Gemini AI - Add Instruction", "")
                If Trim(userInput) = "" Then
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
                prompt = userInput & vbCrLf & vbCrLf & "Data:" & vbCrLf & selectionText
            Else
                userInput = InputBox("Enter your prompt for Gemini:", "Gemini AI", "")
                If Trim(userInput) = "" Then
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
                prompt = userInput
            End If
            
        Case Else ' "auto" or any other value
            ' If selection has text use it, otherwise show input box
            If selectionText <> "" Then
                prompt = selectionText
            Else
                userInput = InputBox("Enter your prompt for Gemini:", "Gemini AI", "")
                If Trim(userInput) = "" Then
                    Application.ScreenUpdating = True
                    Exit Sub
                End If
                prompt = userInput
            End If
    End Select
    
    Application.StatusBar = "Processing Gemini request..."
    
    ' Create HTTP request object
    ' If WinHttp fails, try alternatives: "MSXML2.ServerXMLHTTP" or "MSXML2.XMLHTTP"
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim requestBody As String
    requestBody = BuildGeminiRequest(prompt)
    
    Dim apiUrl As String
    apiUrl = "https://generativelanguage.googleapis.com/" & GEMINI_API_VERSION & "/models/" & GEMINI_MODEL & ":generateContent?key=" & GEMINI_API_KEY
    
    With httpRequest
        .Open "POST", apiUrl, False
        .SetRequestHeader "Content-Type", "application/json"
        .send (requestBody)
    End With
    
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText
        
        Dim completion As String
        completion = ParseGeminiResponse(response)
        
        Dim outputWs As Worksheet
        Set outputWs = GetOrCreateSheet(ActiveWorkbook, "GEMINI_OUTPUT")
        
        Dim outputRange As Range
        Set outputRange = outputWs.Range("A1")
        
        Dim lastRow As Long
        
        ' Output based on mode setting
        If LCase(GEMINI_OUTPUT_MODE) = "single" Then
            ' All text in one cell with line breaks preserved
            outputWs.Cells.Clear
            outputRange.Value = completion
            outputRange.WrapText = True
            lastRow = 1
        Else
            ' Default "lines" mode - each line in separate cell
            Dim lines As Variant
            lines = Split(Replace(Replace(completion, vbCrLf, vbLf), vbCr, vbLf), vbLf)
            lastRow = WriteLinesToRange(lines, outputRange)
        End If
        
        If lastRow > 0 Then
            outputRange.Parent.Activate
            outputRange.Resize(RowSize:=lastRow).Select
            outputRange.Resize(RowSize:=lastRow).EntireColumn.AutoFit
        End If
    Else
        Dim errorMsg As String
        errorMsg = ParseErrorResponse(httpRequest.responseText, httpRequest.Status)
        MsgBox errorMsg, vbCritical, "Request Failed"
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Dim errMsg As String
    errMsg = "Error: " & Err.Description & vbCrLf & _
             "Error number: " & Err.Number
    If Erl > 0 Then errMsg = errMsg & vbCrLf & "Line: " & Erl
    MsgBox errMsg, vbCritical, "Error"
End Sub

Private Function BuildGeminiRequest(ByVal userPrompt As String) As String
    Dim json As Object
    Set json = New Dictionary
    
    Dim contents As Object
    Set contents = New Collection
    
    Dim userMessage As Object
    Set userMessage = New Dictionary
    userMessage.Add "role", "user"
    
    Dim parts As Object
    Set parts = New Collection
    
    Dim textPart As Object
    Set textPart = New Dictionary
    textPart.Add "text", CleanJSONString(userPrompt)
    
    parts.Add textPart
    userMessage.Add "parts", parts
    contents.Add userMessage
    
    json.Add "contents", contents
    
    If GEMINI_SYSTEM_INSTRUCTION <> "" Then
        Dim systemInstruction As Object
        Set systemInstruction = New Dictionary
        
        Dim sysParts As Object
        Set sysParts = New Collection
        
        Dim sysTextPart As Object
        Set sysTextPart = New Dictionary
        sysTextPart.Add "text", GEMINI_SYSTEM_INSTRUCTION
        
        sysParts.Add sysTextPart
        systemInstruction.Add "parts", sysParts
        
        json.Add "systemInstruction", systemInstruction
    End If
    
    Dim generationConfig As Object
    Set generationConfig = New Dictionary
    generationConfig.Add "temperature", GEMINI_TEMPERATURE
    generationConfig.Add "topP", GEMINI_TOP_P
    generationConfig.Add "topK", GEMINI_TOP_K
    generationConfig.Add "maxOutputTokens", GEMINI_MAX_TOKENS
    
    ' Add thinking configuration based on model type
    ' Gemini 3: uses thinkingLevel ("low" or "high")
    ' Gemini 2.5: uses thinkingBudget (0 to disable, or token count)
    If GEMINI_THINKING_MODE <> "" Then
        Dim thinkingConfig As Object
        Set thinkingConfig = New Dictionary
        
        Select Case GEMINI_THINKING_MODE
            Case "level_high"
                thinkingConfig.Add "thinkingLevel", "high"
            Case "level_low"
                thinkingConfig.Add "thinkingLevel", "low"
            Case "budget_0"
                thinkingConfig.Add "thinkingBudget", 0
            Case "budget_auto"
                ' Don't add anything, let model decide
            Case Else
                ' Invalid mode, skip thinking config
        End Select
        
        If thinkingConfig.count > 0 Then
            generationConfig.Add "thinkingConfig", thinkingConfig
        End If
    End If
    
    json.Add "generationConfig", generationConfig
    
    BuildGeminiRequest = JsonConverter.ConvertToJson(json)
End Function

' ===================================================================
' ParseErrorResponse - Parses API error responses with helpful guidance
' Common error codes from Gemini API:
'   400 - Bad Request (invalid parameters)
'   401 - Unauthorized (invalid API key)
'   403 - Forbidden (API key lacks permissions)
'   404 - Not Found (invalid model name)
'   429 - Too Many Requests (rate limit exceeded)
'   500 - Internal Server Error
'   503 - Service Unavailable
' ===================================================================
Private Function ParseErrorResponse(ByVal response As String, ByVal statusCode As Long) As String
    On Error Resume Next
    
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    
    Dim errorMsg As String
    errorMsg = "Gemini API Error (HTTP " & statusCode & ")" & vbCrLf & vbCrLf
    
    ' Add helpful guidance based on status code
    Select Case statusCode
        Case 400
            errorMsg = errorMsg & "BAD REQUEST: Check your request parameters." & vbCrLf
        Case 401
            errorMsg = errorMsg & "UNAUTHORIZED: Your API key is invalid or missing." & vbCrLf & _
                      "Get a new key at: https://aistudio.google.com/apikey" & vbCrLf
        Case 403
            errorMsg = errorMsg & "FORBIDDEN: Your API key lacks required permissions." & vbCrLf
        Case 404
            errorMsg = errorMsg & "NOT FOUND: The model '" & GEMINI_MODEL & "' may not exist." & vbCrLf & _
                      "Check available models at: https://ai.google.dev/gemini-api/docs/models" & vbCrLf
        Case 429
            errorMsg = errorMsg & "RATE LIMITED: Too many requests. Please wait and try again." & vbCrLf & _
                      "See rate limits: https://ai.google.dev/gemini-api/docs/rate-limits" & vbCrLf
        Case 500
            errorMsg = errorMsg & "SERVER ERROR: Google's servers encountered an error." & vbCrLf
        Case 503
            errorMsg = errorMsg & "SERVICE UNAVAILABLE: The API is temporarily down." & vbCrLf
    End Select
    
    errorMsg = errorMsg & vbCrLf
    
    ' Parse JSON error details if available
    If Not json Is Nothing Then
        If json.Exists("error") Then
            Dim errorObj As Object
            Set errorObj = json("error")
            
            If errorObj.Exists("message") Then
                errorMsg = errorMsg & "Details: " & errorObj("message") & vbCrLf
            End If
            
            If errorObj.Exists("status") Then
                errorMsg = errorMsg & "Status: " & errorObj("status") & vbCrLf
            End If
        Else
            errorMsg = errorMsg & "Raw response: " & Left(response, 500)
        End If
    Else
        errorMsg = errorMsg & "Raw response: " & Left(response, 500)
    End If
    
    ParseErrorResponse = errorMsg
    On Error GoTo 0
End Function

' ===================================================================
' ParseGeminiResponse - Extracts text from Gemini API response
' Handles various response scenarios including:
'   - Normal text responses
'   - Safety-blocked content (SAFETY, RECITATION, OTHER)
'   - Empty candidates
'   - Thinking model responses (filters out thought parts)
' ===================================================================
Private Function ParseGeminiResponse(ByVal response As String) As String
    On Error GoTo ErrorHandler
    
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    
    ' Check for prompt feedback (blocked before generation)
    If json.Exists("promptFeedback") Then
        Dim promptFeedback As Object
        Set promptFeedback = json("promptFeedback")
        
        If promptFeedback.Exists("blockReason") Then
            Dim blockReason As String
            blockReason = promptFeedback("blockReason")
            
            Select Case blockReason
                Case "SAFETY"
                    ParseGeminiResponse = "[BLOCKED] Your prompt was blocked due to safety concerns. " & _
                                         "Please revise your prompt and try again."
                Case "OTHER"
                    ParseGeminiResponse = "[BLOCKED] Your prompt was blocked. This may violate " & _
                                         "Google's Terms of Service. See: https://ai.google.dev/terms"
                Case Else
                    ParseGeminiResponse = "[BLOCKED] Prompt blocked. Reason: " & blockReason
            End Select
            Exit Function
        End If
    End If
    
    ' Check for candidates array
    If Not json.Exists("candidates") Then
        ParseGeminiResponse = "[ERROR] Response missing 'candidates' - the model may not have generated output."
        Exit Function
    End If
    
    If json("candidates").count = 0 Then
        ParseGeminiResponse = "[ERROR] No candidates returned - the model did not generate any response."
        Exit Function
    End If
    
    Dim candidate As Object
    Set candidate = json("candidates")(1)
    
    ' Check finish reason for issues
    If candidate.Exists("finishReason") Then
        Dim finishReason As String
        finishReason = candidate("finishReason")
        
        Select Case finishReason
            Case "SAFETY"
                ParseGeminiResponse = "[BLOCKED] Response was blocked due to safety filters. " & _
                                     "Try rephrasing your prompt."
                Exit Function
            Case "RECITATION"
                ParseGeminiResponse = "[BLOCKED] Response blocked due to recitation concerns. " & _
                                     "The output may have resembled copyrighted content. " & _
                                     "Try making your prompt more unique or use higher temperature."
                Exit Function
            Case "OTHER"
                ParseGeminiResponse = "[BLOCKED] Response blocked. Reason: OTHER. " & _
                                     "This may indicate a Terms of Service violation."
                Exit Function
            Case "MAX_TOKENS"
                ' Continue processing but note truncation
            Case "STOP"
                ' Normal completion, continue
        End Select
    End If
    
    ' Check for content
    If Not candidate.Exists("content") Then
        ' Gemini 3 thinking models may have content in a different location
        ' Try to get any text from the raw response
        ParseGeminiResponse = "[ERROR] Candidate missing 'content' - no text was generated."
        Exit Function
    End If
    
    Dim content As Object
    Set content = candidate("content")
    
    ' Handle case where content exists but parts might be missing or named differently
    Dim parts As Object
    If content.Exists("parts") Then
        Set parts = content("parts")
    Else
        ' Try alternative structure for thinking models
        ParseGeminiResponse = "[ERROR] Content missing 'parts' - response structure unexpected."
        Exit Function
    End If
    
    If parts.count = 0 Then
        ParseGeminiResponse = "[ERROR] Parts array is empty - no content generated."
        Exit Function
    End If
    
    ' Extract text from parts (skip "thought" parts from thinking models)
    ' Gemini 3 returns parts with "thought":true for thinking, we want "text" parts only
    Dim textContent As String
    textContent = ""
    
    Dim part As Variant
    For Each part In parts
        If TypeName(part) = "Dictionary" Then
            ' Skip thought parts (internal reasoning from thinking models)
            If part.Exists("thought") Then
                ' This is a thinking part, skip it (we only want the final answer)
            ElseIf part.Exists("text") Then
                textContent = textContent & part("text")
            End If
        End If
    Next part
    
    ' Check if we got any text
    If Len(Trim(textContent)) = 0 Then
        ParseGeminiResponse = "[WARNING] Response contained no text content."
        Exit Function
    End If
    
    ' Add truncation warning if applicable
    If candidate.Exists("finishReason") Then
        If candidate("finishReason") = "MAX_TOKENS" Then
            textContent = textContent & vbCrLf & vbCrLf & _
                         "[NOTE: Response was truncated due to max token limit. " & _
                         "Increase GEMINI_MAX_TOKENS for longer responses.]"
        End If
    End If
    
    ParseGeminiResponse = textContent
    Exit Function
    
ErrorHandler:
    ParseGeminiResponse = "[ERROR] Failed to parse response: " & Err.Description & _
                         vbCrLf & "This may indicate an unexpected API response format."
End Function

' Sanitizes user input for safe inclusion in JSON request
' - Replaces line breaks with spaces
' - Converts double quotes to single quotes (prevents JSON breaking)
' - Escapes backslashes
Private Function CleanJSONString(inputStr As String) As String
    On Error Resume Next
    ' Normalize line breaks to LF (JsonConverter handles escaping to \n)
    CleanJSONString = Replace(inputStr, vbCrLf, vbLf)
    CleanJSONString = Replace(CleanJSONString, vbCr, vbLf)
    ' Preserve vbLf for table structure
    CleanJSONString = Replace(CleanJSONString, """", "'")
    CleanJSONString = Replace(CleanJSONString, "\", "\\")
    On Error GoTo 0
End Function

' Converts escaped quotes (\" ) back to normal quotes (") in API response
' JSON responses contain escaped characters that need unescaping for display
Private Function ReplaceBackslash(text As Variant) As String
    On Error Resume Next
    Dim i As Integer
    Dim newText As String
    newText = ""
    
    For i = 1 To Len(text)
        If Mid(text, i, 2) = "\" & Chr(34) Then
            newText = newText & Chr(34)
            i = i + 1
        Else
            newText = newText & Mid(text, i, 1)
        End If
    Next i
    
    ReplaceBackslash = newText
    On Error GoTo 0
End Function

' Writes an array of text lines to consecutive cells in a worksheet
' - Clears existing content first
' - Escapes lines starting with "=" to prevent Excel formula interpretation
' - Returns the number of lines written
Private Function WriteLinesToRange(lines As Variant, rng As Range) As Long
    Dim i As Long
    
    rng.Worksheet.Cells.ClearContents
    
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = ReplaceBackslash(lines(i))
        
        ' Prefix with apostrophe to prevent formula execution
        If Left(line, 1) = "=" Then
            line = "'" & line
        End If
        
        rng.Cells(i + 1, 1).Value = line
    Next i
    
    WriteLinesToRange = i
End Function

' Checks if the computer has an active internet connection
' Uses Windows API (wininet.dll)
Private Function GetInternetConnectedState() As Boolean
    On Error Resume Next
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function

' Returns a worksheet by name, creating it if it doesn't exist
' Used for output sheets (GEMINI_OUTPUT, etc.)
Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim sheet As Worksheet
    
    For Each sheet In wb.Sheets
        If sheet.Name = sheetName Then
            Set GetOrCreateSheet = sheet
            Exit Function
        End If
    Next sheet
    
    Set GetOrCreateSheet = Sheets.Add(After:=Sheets(Sheets.count))
    GetOrCreateSheet.Name = sheetName
End Function

' ===================================================================
' AskGemini - User Defined Function for use in Excel formulas
'
' Usage:
'   =AskGemini("What is 2+2?")
'   =AskGemini(A1)
'   =AskGemini("Translate to French: " & A1)
'
' Note: Each call makes an API request. Avoid using on large ranges
' to prevent hitting rate limits.
' ===================================================================
Function AskGemini(prompt As String) As String
    On Error GoTo ErrorHandler
    
    If Trim(prompt) = "" Then
        AskGemini = ""
        Exit Function
    End If
    
    ' Check internet connection
    If Not GetInternetConnectedState() Then
        AskGemini = "#NO_INTERNET"
        Exit Function
    End If
    
    Dim httpRequest As Object
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim requestBody As String
    requestBody = BuildGeminiRequest(prompt)
    
    Dim apiUrl As String
    apiUrl = "https://generativelanguage.googleapis.com/" & GEMINI_API_VERSION & _
             "/models/" & GEMINI_MODEL & ":generateContent?key=" & GEMINI_API_KEY
    
    With httpRequest
        .Open "POST", apiUrl, False
        .SetRequestHeader "Content-Type", "application/json"
        .send requestBody
    End With
    
    If httpRequest.Status = 200 Then
        AskGemini = ParseGeminiResponse(httpRequest.responseText)
    Else
        AskGemini = "#API_ERROR_" & httpRequest.Status
    End If
    
    Exit Function
ErrorHandler:
    AskGemini = "#ERROR: " & Err.Description
End Function


