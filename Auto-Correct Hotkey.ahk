#Requires AutoHotkey v2.0
#SingleInstance Force

; ==============================================================================
; CONFIGURATION
; ==============================================================================
API_KEY := "gsk_ApiKeyGoesHere"
MODEL   := "moonshotai/kimi-k2-instruct-0905" ; Using the model from your Python script
URL     := "https://api.groq.com/openai/v1/chat/completions"

; Debug mode - set to true to log to file for troubleshooting
DEBUG_MODE := false
DEBUG_FILE := A_ScriptDir . "\debug_log.txt"

; ==============================================================================
; HOTKEY: Ctrl + Alt + R
; ==============================================================================
^!r:: {
    global DEBUG_MODE, DEBUG_FILE

    if (DEBUG_MODE)
        FileAppend("`n`n=== New Request at " . A_Now . " ===`n", DEBUG_FILE)

    ; 1. Preserve original clipboard
    SavedClip := ClipboardAll()
    A_Clipboard := "" ; Clear clipboard to detect copy success

    ; 2. Copy selected text
    Send "^c"
    if !ClipWait(0.5) {
        ToolTip "Error: Failed to copy text."
        SetTimer () => ToolTip(), -2000
        A_Clipboard := SavedClip
        return
    }

    UserText := A_Clipboard

    if (DEBUG_MODE)
        FileAppend("Input text: " . UserText . "`n", DEBUG_FILE)

    ; 3. Check for embedded instructions (like your Python script)
    ; Looks for text wrapped in { } at the end or beginning
    SystemPrompt := "You are a grammar correction and alignment agent. Your sole and only goal here is to fix any and all grammatical errors with minimal augmentation to the original text or meaning, including spelling, styling, and punctuation. Correct the text while keeping it as original as feasibly possible, not altering correct words or phrases in any way. Return only the final corrected text in your response; do not add any preamble, explanation, or other text."

    ; Regex to find {Instructions} inside the text
    if RegExMatch(UserText, "s)\{(.*?)\}", &Match) {
        SystemPrompt := Match[1]
        ; Remove the instruction from the text sent to LLM
        UserText := StrReplace(UserText, Match[0], "")
    }

    ; 4. UI Feedback
    ToolTip "Processing with Groq..."

    ; 5. Prepare JSON Payload
    ; We must escape quotes and newlines for the JSON string
    SafeText := EscapeJSON(UserText)
    SafeSystem := EscapeJSON(SystemPrompt)

    Body := '{"model": "' . MODEL . '", "messages": [{"role": "system", "content": "' . SafeSystem . '"}, {"role": "user", "content": "' . SafeText . '"}], "temperature": 0.2}'

    if (DEBUG_MODE)
        FileAppend("Request body: " . Body . "`n", DEBUG_FILE)

    ; 6. Send API Request
    try {
        WebRequest := ComObject("WinHttp.WinHttpRequest.5.1")
        WebRequest.Open("POST", URL, true) ; Async=true to prevent UI freeze (handled by WaitForResponse)
        WebRequest.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        WebRequest.SetRequestHeader("Authorization", "Bearer " . API_KEY)
        WebRequest.SetRequestHeader("Accept", "application/json")
        WebRequest.SetRequestHeader("Accept-Charset", "utf-8")
        WebRequest.Send(Body)
        WebRequest.WaitForResponse()

        ; Get response with proper UTF-8 handling using ADODB.Stream
        Response := GetUTF8Response(WebRequest)

        if (DEBUG_MODE)
            FileAppend("Raw response: " . Response . "`n", DEBUG_FILE)

        ; 7. Parse Response - Extract content properly handling escaped quotes
        CorrectedText := ExtractJSONContent(Response)

        if (DEBUG_MODE)
            FileAppend("Extracted content: " . CorrectedText . "`n", DEBUG_FILE)

        if (CorrectedText != "") {
            ; 8. Paste Result
            A_Clipboard := CorrectedText
            Sleep 50 ; Small delay to ensure clipboard is ready
            Send "^v"

            ToolTip "Done!"
            SetTimer () => ToolTip(), -1000
        } else {
            ; Check for errors in response
            ToolTip "API Error: " . SubStr(Response, 1, 100)
            SetTimer () => ToolTip(), -3000
        }

    } catch as e {
        if (DEBUG_MODE)
            FileAppend("Error: " . e.Message . "`n", DEBUG_FILE)
        ToolTip "Connection Error: " . e.Message
        SetTimer () => ToolTip(), -3000
    }

    ; 9. Restore clipboard (Optional - sometimes you want to keep the corrected text)
    Sleep 500
    A_Clipboard := SavedClip
}

; ==============================================================================
; HELPER FUNCTIONS
; ==============================================================================

; Get UTF-8 response from WinHttp using ADODB.Stream for proper encoding
GetUTF8Response(WebRequest) {
    try {
        ; Use ADODB.Stream for proper UTF-8 decoding
        ResponseBody := WebRequest.ResponseBody

        Stream := ComObject("ADODB.Stream")
        Stream.Type := 1  ; Binary
        Stream.Open()
        Stream.Write(ResponseBody)
        Stream.Position := 0
        Stream.Type := 2  ; Text
        Stream.Charset := "UTF-8"
        Result := Stream.ReadText()
        Stream.Close()

        return Result
    } catch {
        ; Fallback to ResponseText if ADODB fails
        return WebRequest.ResponseText
    }
}

; Escape string for JSON - handles all required escape sequences
EscapeJSON(str) {
    ; Must escape backslash FIRST before adding more backslashes
    str := StrReplace(str, "\", "\\")
    ; Escape double quotes
    str := StrReplace(str, '"', '\"')
    ; Escape control characters
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`t", "\t")
    ; Escape other control characters that could cause issues
    str := StrReplace(str, Chr(8), "\b")   ; Backspace
    str := StrReplace(str, Chr(12), "\f")  ; Form feed
    return str
}

; Properly extract JSON string content, handling escaped characters
ExtractJSONContent(response) {
    ; Find the "content": " marker
    startMarker := '"content":'
    startPos := InStr(response, startMarker)
    if (!startPos)
        return ""

    ; Move past the marker and any whitespace
    startPos += StrLen(startMarker)

    ; Skip whitespace to find the opening quote
    while (startPos <= StrLen(response)) {
        char := SubStr(response, startPos, 1)
        if (char = '"')
            break
        if (char != " " && char != "`t" && char != "`n" && char != "`r")
            return "" ; Unexpected character
        startPos++
    }

    if (SubStr(response, startPos, 1) != '"')
        return ""

    ; Move past the opening quote
    startPos++

    ; Now find the closing quote, accounting for escaped characters
    content := ""
    pos := startPos
    while (pos <= StrLen(response)) {
        char := SubStr(response, pos, 1)

        if (char = "\") {
            ; Escape sequence - grab next char too
            nextChar := SubStr(response, pos + 1, 1)
            content .= char . nextChar
            pos += 2
            continue
        }

        if (char = '"') {
            ; End of string (unescaped quote)
            break
        }

        content .= char
        pos++
    }

    ; Now unescape the content
    return UnescapeJSON(content)
}

; Unescape JSON string - handles all standard JSON escape sequences
UnescapeJSON(str) {
    ; IMPORTANT: Process \\ FIRST to avoid double-processing
    ; Use a placeholder that won't appear in normal text
    placeholder := Chr(1)
    str := StrReplace(str, "\\", placeholder)

    ; Handle standard JSON escape sequences
    str := StrReplace(str, '\"', '"')
    str := StrReplace(str, "\n", "`n")
    str := StrReplace(str, "\r", "`r")
    str := StrReplace(str, "\t", "`t")
    str := StrReplace(str, "\b", Chr(8))   ; Backspace
    str := StrReplace(str, "\f", Chr(12))  ; Form feed
    str := StrReplace(str, "\/", "/")      ; Escaped forward slash (optional in JSON)

    ; Handle Unicode escapes \uXXXX
    while RegExMatch(str, "\\u([0-9A-Fa-f]{4})", &match) {
        codePoint := Integer("0x" . match[1])
        str := StrReplace(str, match[0], Chr(codePoint), , 1)
    }

    ; Restore placeholder to actual backslash
    str := StrReplace(str, placeholder, "\")

    return str
}
