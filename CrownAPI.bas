Attribute VB_Name = "CrownAPI"
Option Explicit

' =====================================================================
' Crown Superior - talking to the website directly
' ---------------------------------------------------------------------
' Replaces the parts of this tool that open crownsuperior.com in Chrome
' and read the page, or type into the web forms. Nothing here needs a
' browser, a login, or a page to stay the same shape.
'
' Putting this into the tool, without touching any code:
'
'   1. In the tool, click the Developer tab, then Visual Basic.
'      (Alt+F11 does the same thing, but on a lot of laptops the F
'      keys need Fn held down, which is why it may look dead.)
'   2. In the window that opens: File, Import File, and pick this file.
'   3. Close that window. You are back in the tool.
'   4. Developer tab, Macros, run CrownTestConnection.
'
' The first time you run it, it asks you for your key and remembers it.
' There is nothing to edit in here and nothing to paste. Save the
' workbook once afterwards so it keeps the key.
'
' Your key comes from
'   https://crownsuperior.com/index.php?cf_action=apikey
' signed in there as an administrator. To change it later, run the
' CrownSetKey macro.
'
' Developer, Macros, and the four things you can run:
'   CrownTestConnection  - is the website answering, is my key right
'   CrownLoadQuotes      - put the newest quotes on a sheet
'   CrownWriteQuoteBack  - send a finished quote back to the website
'   CrownSetKey          - type in a new key
'
' The key is a password. Anyone who has it can read and write quotes
' and policies, so do not email it around, and if it gets out, press
' "Make a new key" on that same page.
' =====================================================================

Private Const CROWN_URL As String = "https://crownsuperior.com/index.php?cf_action=api"

' Where the key is kept: a hidden name inside this workbook, so it is
' typed once and never goes near the code.
Private Const CROWN_KEY_NAME As String = "CrownSuperiorKey"

' Forms this reaches. 18 is Get Your Quote, 11 is the policy.
Public Const CROWN_FORM_QUOTE As Long = 18
Public Const CROWN_FORM_POLICY As Long = 11


' ---------------------------------------------------------------------
' The key
'
' Kept in the workbook itself rather than in this code, so nobody has
' to find a line and paste into it. Asked for once, then remembered.
' ---------------------------------------------------------------------

' Which copy of this file is in the workbook.
'
' File > Import does not replace a module of the same name - the old one
' can still be sitting there - so there has to be a way to ask. Run
' CrownVersion in the CrownPayments module to see this and its own.
Public Function CrownApiStamp() As String
    CrownApiStamp = "29 Sep - retries a dropped connection, and CrownClone"
End Function

Private Function CrownKey(Optional ByVal askIfMissing As Boolean = True) As String
    Dim nm As Object
    Dim value As String

    On Error Resume Next
    Set nm = ThisWorkbook.Names(CROWN_KEY_NAME)
    On Error GoTo 0

    If Not nm Is Nothing Then
        value = nm.RefersTo                  ' comes back looking like ="thekey"
        If Left$(value, 1) = "=" Then value = Mid$(value, 2)
        value = Replace(value, """", "")
        value = Trim$(value)
    End If

    If Len(value) = 0 And askIfMissing Then
        value = Trim$(InputBox( _
            "Paste your Crown Superior key." & vbCrLf & vbCrLf & _
            "You get it from crownsuperior.com/index.php?cf_action=apikey" & vbCrLf & _
            "while signed in there as an administrator." & vbCrLf & vbCrLf & _
            "It is only asked for once.", "Crown Superior"))

        If Len(value) > 0 Then CrownStoreKey value
    End If

    CrownKey = value
End Function

Private Sub CrownStoreKey(ByVal value As String)
    On Error Resume Next
    ThisWorkbook.Names(CROWN_KEY_NAME).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:=CROWN_KEY_NAME, RefersTo:="=""" & value & """", Visible:=False
End Sub

' Run this from Developer > Macros to put a different key in.
Public Sub CrownSetKey()
    Dim current As String, value As String

    current = CrownKey(False)

    value = Trim$(InputBox( _
        "Paste your Crown Superior key." & vbCrLf & vbCrLf & _
        "From crownsuperior.com/index.php?cf_action=apikey, signed in there as an administrator.", _
        "Crown Superior", current))

    If Len(value) = 0 Then Exit Sub

    CrownStoreKey value
    MsgBox "Saved. Save the workbook once so it is still here next time.", vbInformation, "Crown Superior"
End Sub


' ---------------------------------------------------------------------
' The plumbing
' ---------------------------------------------------------------------

' Tries three times before giving up.
'
' Sending several hundred of these in a row, one will occasionally come
' back as "the server returned an invalid or unrecognized response" -
' a connection that was reused after the far end had quietly dropped it.
' A fresh object and another go clears it. Asking for the connection to
' be closed each time makes it rarer still.
'
' Returns "" if all three fail, so the caller can carry on to the next
' record instead of the whole run stopping on one bad moment.
Public Function CrownPost(ByVal body As String) As String
    Dim http As Object
    Dim attempt As Long
    Dim answer As String
    Dim wentThrough As Boolean
    Dim pause As Double

    For attempt = 1 To 3
        answer = ""
        wentThrough = False

        On Error Resume Next
        Err.Clear

        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts 15000, 15000, 30000, 120000
        http.Open "POST", CROWN_URL, False
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.setRequestHeader "Connection", "close"
        http.send body

        If Err.Number = 0 Then
            answer = http.responseText
            wentThrough = (Err.Number = 0)
        End If

        Err.Clear
        On Error GoTo 0

        Set http = Nothing

        If wentThrough Then Exit For

        ' A second between tries, without locking up Excel.
        pause = Timer
        Do While Timer < pause + 1
            DoEvents
        Loop
    Next attempt

    CrownPost = answer
End Function

' Percent-encoding. Written out rather than borrowed, because the
' library ones differ between 32 and 64 bit Office.
Public Function CrownEnc(ByVal text As String) As String
    Dim bytes() As Byte
    Dim raw As Variant
    Dim stream As Object
    Dim i As Long, b As Long, out As String

    ' Nothing to encode. This has to come first: with an empty string the
    ' stream holds only its byte order mark, so reading past that returns
    ' Null, and Null into a Byte array is a type mismatch.
    If Len(text) = 0 Then Exit Function

    ' UTF-8 first, so accents and long dashes survive the trip.
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3                    ' step over the byte order mark
    raw = stream.Read
    stream.Close

    If IsNull(raw) Or IsEmpty(raw) Then Exit Function

    bytes = raw

    For i = LBound(bytes) To UBound(bytes)
        b = bytes(i)

        If (b >= 48 And b <= 57) Or (b >= 65 And b <= 90) Or (b >= 97 And b <= 122) _
           Or b = 45 Or b = 46 Or b = 95 Or b = 126 Then
            out = out & Chr$(b)
        Else
            out = out & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i

    CrownEnc = out
End Function

Public Function CrownBody(ByVal action As String) As String
    CrownBody = "do=" & CrownEnc(action) & "&key=" & CrownEnc(CrownKey())
End Function

' Quote a value for JSON.
Private Function JsonStr(ByVal text As String) As String
    Dim out As String

    out = Replace(text, "\", "\\")
    out = Replace(out, """", "\""")
    out = Replace(out, vbCrLf, "\n")
    out = Replace(out, vbCr, "\n")
    out = Replace(out, vbLf, "\n")
    out = Replace(out, vbTab, "\t")

    JsonStr = """" & out & """"
End Function


' ---------------------------------------------------------------------
' The calls
' ---------------------------------------------------------------------

' Is the door open and is the key right?
Public Function CrownPing() As String
    CrownPing = CrownPost(CrownBody("ping"))
    Debug.Print CrownPing
End Function

' Quotes or policies as CSV text. Name the columns you want - the field
' names are exactly as they appear on the form, e.g. "First Name",
' "Vin_Number", "Policy_No".
'
'   sinceId  0 for everything, or the last id you already have
'   limit    up to 500
Public Function CrownListCsv(ByVal formId As Long, ByVal sinceId As Long, _
                             ByVal limit As Long, ByVal fields As String) As String
    Dim body As String

    body = CrownBody("list") _
         & "&form=" & formId _
         & "&since_id=" & sinceId _
         & "&limit=" & limit _
         & "&format=csv" _
         & "&fields=" & CrownEnc(fields)

    CrownListCsv = CrownPost(body)
End Function

' One record as CSV text: a header row and a single row of values.
Public Function CrownGetCsv(ByVal formId As Long, ByVal recordId As Long, _
                            ByVal fields As String) As String
    Dim body As String

    body = CrownBody("get") _
         & "&form=" & formId _
         & "&id=" & recordId _
         & "&format=csv" _
         & "&fields=" & CrownEnc(fields)

    CrownGetCsv = CrownPost(body)
End Function

' One answer from one record. Blank if it has never been filled in.
Public Function CrownGetValue(ByVal formId As Long, ByVal recordId As Long, _
                              ByVal fieldName As String) As String
    Dim csv As String, lines() As String, parts() As String

    csv = CrownGetCsv(formId, recordId, fieldName)
    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)

    If UBound(lines) < 1 Then
        CrownGetValue = ""
        Exit Function
    End If

    parts = CrownSplitCsvLine(lines(1))

    If UBound(parts) >= 2 Then
        CrownGetValue = parts(2)
    Else
        CrownGetValue = ""
    End If
End Function

' Write answers back onto a record that already exists.
'
'   names   e.g. Array("Policy_No", "Insurance_Co_Name", "Amount ")
'   values  e.g. Array("GAI-211007", "United Auto", "431.00")
'
' Answers the server's reply. "written" is how many were saved and
' "unknown_fields" lists any name the form does not have - always worth
' a look, because a name that is not there is simply not written.
Public Function CrownUpdate(ByVal formId As Long, ByVal recordId As Long, _
                            ByRef names As Variant, ByRef values As Variant) As String
    Dim body As String

    body = CrownBody("update") _
         & "&form=" & formId _
         & "&id=" & recordId _
         & "&values=" & CrownEnc(CrownJson(names, values))

    CrownUpdate = CrownPost(body)
End Function

' Put a brand new quote in. Same two arrays. The reply carries the new id.
Public Function CrownCreate(ByVal formId As Long, _
                            ByRef names As Variant, ByRef values As Variant) As String
    Dim body As String

    body = CrownBody("create") _
         & "&form=" & formId _
         & "&values=" & CrownEnc(CrownJson(names, values))

    CrownCreate = CrownPost(body)
End Function

' Copy a record that already exists, changing only what is named.
'
' A renewed policy is the same customer, the same car and the same cover
' with a new number and new dates. Rather than rebuild all of that from
' here, this asks the website to take a copy of the old record and put the
' handful of differences on top of it.
'
'   CrownClone 11, 12083, Array("Policy number", "Description"), _
'                         Array("GAI -10202602", "Renewal")
'
' The reply carries the new record's id.
Public Function CrownClone(ByVal formId As Long, ByVal sourceId As Long, _
                           ByRef names As Variant, ByRef values As Variant) As String
    Dim body As String

    body = CrownBody("clone") _
         & "&form=" & formId _
         & "&id=" & sourceId _
         & "&values=" & CrownEnc(CrownJson(names, values))

    CrownClone = CrownPost(body)
End Function

Private Function CrownJson(ByRef names As Variant, ByRef values As Variant) As String
    Dim i As Long, out As String

    out = "{"

    For i = LBound(names) To UBound(names)
        If i > LBound(names) Then out = out & ","
        out = out & JsonStr(CStr(names(i))) & ":" & JsonStr(CStr(values(i)))
    Next i

    CrownJson = out & "}"
End Function


' ---------------------------------------------------------------------
' Getting it onto a sheet
' ---------------------------------------------------------------------

' Fills a sheet with quotes, header row included. Returns how many rows
' were written, and puts the highest id it saw into lastId so the next
' run can carry on from there.
'
'   CrownListToSheet Worksheets("Input"), 18, 0, 100, _
'       "First Name,Last Name,Date of Birth ,Drivers License Number,DL_state,Vin_Number", lastId
Public Function CrownListToSheet(ByVal target As Object, ByVal formId As Long, _
                                 ByVal sinceId As Long, ByVal limit As Long, _
                                 ByVal fields As String, ByRef lastId As Long, _
                                 Optional ByVal firstRow As Long = 1, _
                                 Optional ByVal firstCol As Long = 1) As Long
    Dim csv As String, lines() As String, parts() As String
    Dim r As Long, c As Long, written As Long

    csv = CrownListCsv(formId, sinceId, limit, fields)
    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)

    If UBound(lines) < 0 Then Exit Function

    If Left$(lines(0), 5) = "error" Then
        MsgBox "The website answered: " & lines(0), vbExclamation
        Exit Function
    End If

    ' Nothing is cleared. It writes over whatever is in the block it needs
    ' and leaves the rest of the sheet alone, because wiping a sheet that
    ' somebody has been working in is not a mistake you can undo.
    For r = LBound(lines) To UBound(lines)
        If Len(Trim$(lines(r))) > 0 Then
            parts = CrownSplitCsvLine(lines(r))

            For c = LBound(parts) To UBound(parts)
                target.Cells(firstRow + r, firstCol + c).Value = parts(c)
            Next c

            If r > 0 Then
                written = written + 1

                If IsNumeric(parts(0)) Then
                    If CLng(parts(0)) > lastId Then lastId = CLng(parts(0))
                End If
            End If
        End If
    Next r

    CrownListToSheet = written
End Function

' One CSV line into its cells. Handles the doubled quotes and the commas
' inside a value, which Split on its own would get wrong.
Public Function CrownSplitCsvLine(ByVal line As String) As String()
    Dim out() As String
    Dim i As Long, ch As String
    Dim inQuotes As Boolean, current As String
    Dim count As Long

    ReDim out(0 To 0)
    count = 0

    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)

        If inQuotes Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    current = current & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                current = current & ch
            End If
        Else
            If ch = """" Then
                inQuotes = True
            ElseIf ch = "," Then
                ReDim Preserve out(0 To count)
                out(count) = current
                count = count + 1
                current = ""
            Else
                current = current & ch
            End If
        End If
    Next i

    ReDim Preserve out(0 To count)
    out(count) = current

    CrownSplitCsvLine = out
End Function


' ---------------------------------------------------------------------
' The three you can run from Developer > Macros
'
' Nothing here needs anything typed into it. They ask for what they
' need as they go, and tell you what happened when they finish.
' ---------------------------------------------------------------------

' Is the website answering, and is the key right?
Public Sub CrownTestConnection()
    Dim answer As String

    If Len(CrownKey()) = 0 Then
        MsgBox "No key, so there is nothing to test yet." & vbCrLf & vbCrLf & _
               "Run this again and paste the key when it asks.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    answer = CrownPost(CrownBody("ping"))

    If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
        MsgBox "Connected. The website answered and the key is right." & vbCrLf & vbCrLf & _
               "Save the workbook once so it keeps the key.", vbInformation, "Crown Superior"
    ElseIf InStr(1, answer, "bad-key", vbTextCompare) > 0 Then
        MsgBox "The website answered, but it does not recognise that key." & vbCrLf & vbCrLf & _
               "Check it against crownsuperior.com/index.php?cf_action=apikey", vbExclamation, "Crown Superior"
    ElseIf Len(Trim$(answer)) = 0 Then
        MsgBox "No answer at all. That is usually no internet, or a firewall in the way.", vbExclamation, "Crown Superior"
    Else
        MsgBox "The website said:" & vbCrLf & vbCrLf & answer, vbExclamation, "Crown Superior"
    End If
End Sub

' The newest quotes, onto a sheet of their own.
' Retired, and left here saying so.
'
' It asked "how many of the newest quotes" and then asked the website for
' them starting from the beginning - and the website answers a question
' like that oldest first, because it is built for walking forward from
' where you got to last time. So asking for a hundred of the newest
' handed back the oldest hundred, which is why quotes from 2024 arrived.
'
' It also wrote them to a sheet of its own, which is not the sheet the
' rest of this works from. Two lists of quotes is one too many.
Public Sub CrownLoadQuotes()
    MsgBox "This one has been replaced." & vbCrLf & vbCrLf & _
           "Use CrownFetchNewQuotes instead. It brings in the quotes that have " & _
           "come in since the last time and puts them on Crown Quote List, " & _
           "which is the sheet CrownQuoteToEditData reads from." & vbCrLf & vbCrLf & _
           "The Crown Quotes tab this used to make can be deleted - nothing " & _
           "reads it.", vbInformation, "Crown Superior"
End Sub

' A finished quote, back onto the website.
Public Sub CrownWriteQuoteBack()
    Dim id As String, policyNo As String, company As String, amount As String
    Dim answer As String
    Dim names As Variant, values As Variant

    If Len(CrownKey()) = 0 Then
        MsgBox "No key yet - run CrownTestConnection and paste it when it asks.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    id = InputBox("Which quote number?" & vbCrLf & _
                  "(the first column on the Crown Quotes sheet)", "Crown Superior")
    If Len(Trim$(id)) = 0 Then Exit Sub
    If Not IsNumeric(id) Then
        MsgBox "That is not a quote number.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    company = InputBox("Insurance company?", "Crown Superior")
    policyNo = InputBox("Policy number?", "Crown Superior")
    amount = InputBox("Total amount due?", "Crown Superior")

    names = Array("Insurance_Co_Name", "Policy_No", "Amount ")
    values = Array(company, policyNo, amount)

    answer = CrownUpdate(CROWN_FORM_QUOTE, CLng(id), names, values)

    If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
        MsgBox "Written onto quote " & id & " on the website.", vbInformation, "Crown Superior"
    Else
        MsgBox "It did not go through. The website said:" & vbCrLf & vbCrLf & answer, vbExclamation, "Crown Superior"
    End If
End Sub

' A sheet by that name, made if it is not there yet.
Public Function CrownSheet(ByVal sheetName As String) As Object
    Dim ws As Object

    For Each ws In ActiveWorkbook.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set CrownSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ActiveWorkbook.Worksheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
    ws.Name = sheetName
    Set CrownSheet = ws
End Function
