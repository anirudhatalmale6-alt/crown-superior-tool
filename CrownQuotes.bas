Attribute VB_Name = "CrownQuotes"
Option Explicit

' =====================================================================
' Crown Superior - quotes into the tool
' ---------------------------------------------------------------------
' Two ways in, both ending where you already work:
'
'   CrownFetchNewQuotes    the website's own quotes, only the ones that
'                          have come in since the last fetch
'   CrownImportQuoteFile   a Quotes sheet saved as a csv, picked by hand,
'                          for when you would rather do it the old way
'
' Neither of them does the placing. Both write a file in the layout the
' website's quote export already has, and then hand it to your own
' ImportquoteToAlloutputsheet, which fills Output, and to
' RetrieveDataByRowNumber, which brings a row onto Edit Data. Your 145
' column mappings are the ones doing the work in both cases - there is no
' second copy of them in here to drift out of step with the first.
' =====================================================================

' The sheet the picked file is expected to look like - the Quotes tab,
' the one the chatbot writes into.
Private Const QUOTE_SHEET_FIRST As String = "saved_at"

' Where the website's quote number lands once your own import has placed
' it - m54 sends SubmissionId to this column. It is what tells a quote
' already on the list from a new one.
Private Const QUOTE_ID_HEADING As String = "Submitter's User ID"


' ---------------------------------------------------------------------
' 1. The website's quotes, only what is new
' ---------------------------------------------------------------------

Public Sub CrownFetchNewQuotes()
    Dim csv As String, path As String, kept As String
    Dim lines() As String
    Dim highest As Long, since As Long
    Dim brought As Long

    since = CrownLastQuoteId()

    csv = CrownPost(CrownBody("quotecsv") & "&limit=200&since_id=" & since)

    If Len(csv) = 0 Then
        MsgBox "The website did not answer. Nothing has changed." & vbCrLf & vbCrLf & _
               "If it stays like this, use CrownImportQuoteFile and pick a file instead.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    If Left$(csv, 5) = "error" Then
        MsgBox "The website answered: " & csv, vbExclamation, "Crown Superior"
        Exit Sub
    End If

    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)
    brought = CrownCountRows(lines)

    If brought = 0 Then
        MsgBox "No new quotes since the last fetch." & vbCrLf & vbCrLf & _
               "The last one brought in was number " & since & ".", _
               vbInformation, "Crown Superior"
        Exit Sub
    End If

    highest = CrownHighestId(lines)
    path = CrownWriteTemp(csv, "crown-quotes-")

    CrownHandToTool path

    kept = CrownKeepThem()

    ' Only remember the number if they actually landed. Saved regardless,
    ' a failure in the middle of the placing would skip those quotes for
    ' good and nothing would ever say so.
    If InStr(1, kept, "NOT added") = 0 And highest > since Then
        CrownSaveLastQuoteId highest
    End If

    MsgBox brought & " new quote(s) brought in." & vbCrLf & _
           kept & vbCrLf & _
           "Edit Data is showing the row you had open." & vbCrLf & vbCrLf & _
           "Next time this will start from quote number " & highest & ".", _
           vbInformation, "Crown Superior"
End Sub

' The last quote number brought in, remembered in the workbook itself so
' it survives closing it.
Private Function CrownLastQuoteId() As Long
    Dim nm As Object
    Dim value As String

    On Error Resume Next
    Set nm = ThisWorkbook.Names("CrownLastQuoteId")
    On Error GoTo 0

    If nm Is Nothing Then Exit Function

    value = nm.RefersTo
    If Left$(value, 1) = "=" Then value = Mid$(value, 2)
    value = Replace(value, """", "")

    If IsNumeric(value) Then CrownLastQuoteId = CLng(Val(value))
End Function

Private Sub CrownSaveLastQuoteId(ByVal id As Long)
    On Error Resume Next
    ThisWorkbook.Names("CrownLastQuoteId").Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add Name:="CrownLastQuoteId", _
                           RefersTo:="=""" & CStr(id) & """", Visible:=False
End Sub

' How many quotes are in what came back - the header line is not one.
Private Function CrownCountRows(ByRef lines() As String) As Long
    Dim i As Long

    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then CrownCountRows = CrownCountRows + 1
    Next i
End Function

' The biggest quote number in what came back. Column A is the number.
Private Function CrownHighestId(ByRef lines() As String) As Long
    Dim parts() As String
    Dim i As Long, id As Long

    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            parts = CrownSplitCsvLine(lines(i))

            If UBound(parts) >= 0 Then
                id = CLng(Val(parts(0)))
                If id > CrownHighestId Then CrownHighestId = id
            End If
        End If
    Next i
End Function


' Start again from an older quote.
'
' CrownFetchNewQuotes counts forward from the last one it brought in, which
' is what stops it fetching the same quotes twice. If you ever need the
' older ones back - a sheet rebuilt, or something missed - this is how you
' tell it where to start. The next fetch brings in everything after the
' number you give it.
Public Sub CrownFetchQuotesFrom()
    Dim answer As String
    Dim since As Long

    since = CrownLastQuoteId()

    answer = InputBox( _
        "The next fetch will bring in quotes NEWER than this number." & vbCrLf & vbCrLf & _
        "It is " & since & " at the moment." & vbCrLf & _
        "Put in a lower number to pick up older quotes, or 0 for all of them." & vbCrLf & vbCrLf & _
        "Nothing is fetched now - run CrownFetchNewQuotes afterwards.", _
        "Crown Superior", CStr(since))

    If Len(Trim$(answer)) = 0 Then Exit Sub

    If Not IsNumeric(answer) Then
        MsgBox "That is not a quote number.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    CrownSaveLastQuoteId CLng(Val(answer))

    MsgBox "Set to " & CLng(Val(answer)) & "." & vbCrLf & vbCrLf & _
           "Run CrownFetchNewQuotes and it will bring in everything after that." & vbCrLf & _
           "Anything already on Crown Quote List is recognised and not added twice.", _
           vbInformation, "Crown Superior"
End Sub


' Put the two heading rows side by side so the difference can be seen.
'
' When the append refuses it names the first column where Output and Crown
' Quote List disagree, which is enough to know something is wrong but not
' always enough to know why. This writes both rows out in full, marks every
' column where they differ, and leaves it on a sheet you can send me.
'
' It reads only. Nothing is copied, changed or cleared by it.
Public Sub CrownCompareQuoteHeadings()
    Dim output As Worksheet, list As Worksheet, out As Worksheet
    Dim wide As Long, c As Long, differences As Long
    Dim a As String, b As String

    On Error Resume Next
    Set output = ThisWorkbook.Sheets("Output")
    Set list = ThisWorkbook.Sheets("Crown Quote List")
    On Error GoTo 0

    If output Is Nothing Or list Is Nothing Then
        MsgBox "This needs both the Output and Crown Quote List sheets.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    wide = output.Cells(1, output.Columns.Count).End(xlToLeft).Column

    If list.Cells(1, list.Columns.Count).End(xlToLeft).Column > wide Then
        wide = list.Cells(1, list.Columns.Count).End(xlToLeft).Column
    End If

    If wide < 1 Then
        MsgBox "Neither sheet has any headings on it.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set out = CrownSheet("Quote Headings")
    out.Cells.ClearContents
    out.Range("A1:D1").value = Array("Column", "Output says", "Crown Quote List says", "Same?")

    For c = 1 To wide
        a = Trim$(CStr(output.Cells(1, c).value))
        b = Trim$(CStr(list.Cells(1, c).value))

        out.Cells(c + 1, 1).value = c
        out.Cells(c + 1, 2).value = a
        out.Cells(c + 1, 3).value = b

        If StrComp(a, b, vbTextCompare) = 0 Then
            out.Cells(c + 1, 4).value = "yes"
        Else
            out.Cells(c + 1, 4).value = "NO"
            differences = differences + 1
        End If
    Next c

    out.Columns("A:D").AutoFit
    out.Activate

    MsgBox wide & " columns compared." & vbCrLf & _
           differences & " of them differ." & vbCrLf & vbCrLf & _
           "The Quote Headings sheet has both rows side by side. Send me that " & _
           "sheet and I will tell you which import to run.", _
           vbInformation, "Crown Superior"
End Sub


' ---------------------------------------------------------------------
' 2. A file, picked by hand, the old way
' ---------------------------------------------------------------------

Public Sub CrownImportQuoteFile()
    Dim picked As Variant
    Dim head As String, csv As String, path As String

    picked = Application.GetOpenFilename( _
        "Comma separated (*.csv), *.csv, All files (*.*), *.*", , _
        "Pick the Quotes file to bring in")

    If VarType(picked) = vbBoolean Then Exit Sub

    ' The website's own export can go straight through - it is already in
    ' the shape the tool reads.
    If CrownLooksLikeExport(CStr(picked)) Then
        CrownHandToTool CStr(picked)
        MsgBox "Brought in." & vbCrLf & CrownKeepThem() & vbCrLf & _
               "Edit Data is showing the row you had open.", _
               vbInformation, "Crown Superior"
        Exit Sub
    End If

    head = CrownQuoteHeader()

    If Len(head) = 0 Then
        MsgBox "That file is a Quotes sheet rather than a website export, so it " & _
               "has to be rearranged before the tool can read it - and the website " & _
               "has to say what the columns are." & vbCrLf & vbCrLf & _
               "The website did not answer just now. Try again in a moment.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    csv = CrownQuoteSheetToExport(CStr(picked), head)

    If Len(csv) = 0 Then
        MsgBox "Nothing in that file looked like a quote." & vbCrLf & vbCrLf & _
               "It should be the Quotes tab saved as a csv - the one whose first " & _
               "column is " & QUOTE_SHEET_FIRST & ".", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    path = CrownWriteTemp(csv, "crown-quotefile-")
    CrownHandToTool path

    MsgBox "Brought in." & vbCrLf & CrownKeepThem() & vbCrLf & _
           "Edit Data is showing the row you had open.", vbInformation, "Crown Superior"
End Sub

' Is this already a website export? Column A of one says SubmissionId.
Private Function CrownLooksLikeExport(ByVal path As String) As Boolean
    Dim fileNumber As Integer
    Dim line As String
    Dim parts() As String

    On Error GoTo Finished

    fileNumber = FreeFile
    Open path For Input As #fileNumber
    If Not EOF(fileNumber) Then Line Input #fileNumber, line
    Close #fileNumber

    parts = CrownSplitCsvLine(line)

    If UBound(parts) >= 0 Then
        CrownLooksLikeExport = (StrComp(Trim$(parts(0)), "SubmissionId", vbTextCompare) = 0)
    End If

    Exit Function

Finished:
    On Error Resume Next
    Close #fileNumber
End Function

' The website's export header row, and nothing else.
Private Function CrownQuoteHeader() As String
    Dim csv As String
    Dim lines() As String

    csv = CrownPost(CrownBody("quotecsv") & "&limit=1")

    If Len(csv) = 0 Or Left$(csv, 5) = "error" Then Exit Function

    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)

    If UBound(lines) >= 0 Then CrownQuoteHeader = lines(0)
End Function

' ---------------------------------------------------------------------
' 3. Turning a Quotes sheet into a website export
' ---------------------------------------------------------------------
'
' The Quotes tab is the chatbot's own layout - saved_at, quote_type,
' first_name - and the tool reads the website's, which is wider and in a
' different order. This puts each heading under the column the website
' would have put it in.
'
' Matched by heading, never by position: a column added to the sheet
' should not move everything after it into the wrong box.

Private Function CrownQuotePairs() As Object
    Dim map As Object

    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = 1                     ' headings, so case does not matter

    map.Add "quote_type", "Coverage"
    map.Add "how_did_you_hear", "Source"
    map.Add "first_name", "First Name"
    map.Add "middle_name", "Middle name"
    map.Add "last_name", "Last Name"
    map.Add "phone", "Phone number "
    map.Add "email", "Email"
    map.Add "dob", "Date of Birth "
    map.Add "address", "Address"
    map.Add "city", "CIty"
    map.Add "state", "State"
    map.Add "zip", "Zip_Code"
    map.Add "drivers_license_number", "Drivers License Number"
    map.Add "license_state", "DL_state"
    map.Add "marital_status", "Marital Status "
    map.Add "gender", "Gender"
    map.Add "occupation", "JOB_TITLE_"
    map.Add "industry", "Industry_"
    map.Add "previous_company", "Name of previous insurance Carrier "
    map.Add "new_purchase", "New Purchase"
    map.Add "garaging_same_as_home", "garaging_address_the_same_as_home"
    map.Add "garaging_address", "Garagining_address"
    map.Add "garaging_city", "Garaging_city_"
    map.Add "garaging_state", "Garaging_state_"
    map.Add "garaging_zip", "Garaging_zip_"
    map.Add "second_driver", "Second driver"
    map.Add "driver_2_excluded", "Exclude "
    map.Add "driver_2_first_name", "Driver 2 First"
    map.Add "driver_2_last_name", "Driver 2 Last"
    map.Add "driver_2_dob", "Driver 2 dob"
    map.Add "driver_2_gender", "driver 2 Gender "
    map.Add "driver_2_marital_status", "Driver 2 Marital Status "
    map.Add "driver_2_relationship", "Driver_2_relation"
    map.Add "driver_2_license", "Driver 2 DL"
    map.Add "driver_2_license_state", "DL_state_Drv 2"
    map.Add "driver_2_industry", "Driver 2 industry "
    map.Add "driver_2_occupation", "Driver 2 JOB TITLE"
    map.Add "vehicle_1_year", "Year"
    map.Add "vehicle_1_make", "Make"
    map.Add "vehicle_1_model", "Model"
    map.Add "vehicle_1_vin", "Vin_Number"
    map.Add "vehicle_1_mileage", "Mileage "
    map.Add "vehicle_1_commercial", "vehicle 1 commercial use"
    map.Add "vehicle_1_deductibles", "Vehicle 1 Deductibles"
    map.Add "notes", "Notes"

    Set CrownQuotePairs = map
End Function

' Put a value into the words the website would have used.
'
' The website translates these on its way in - Georgia becomes GA, $250
' becomes 250/250 - so a file brought in by hand has to arrive saying the
' same things. Two ways in that disagree about what a state is called is
' worse than either way on its own.
Private Function CrownQuoteTidy(ByVal heading As String, ByVal value As String) As String
    Dim low As String
    Dim when As Date

    CrownQuoteTidy = value
    low = LCase$(Trim$(heading))

    If low = "state" Or low = "license_state" Or low = "garaging_state" _
       Or low = "driver_2_license_state" Then
        CrownQuoteTidy = CrownStateCode(value)
        Exit Function
    End If

    If low = "dob" Or low = "driver_2_dob" Then
        On Error Resume Next
        when = CDate(value)
        On Error GoTo 0

        If Year(when) > 1900 Then CrownQuoteTidy = Format$(when, "m/d/yyyy")
        Exit Function
    End If

    If low = "quote_type" Then
        If InStr(1, value, "full", vbTextCompare) > 0 Then
            CrownQuoteTidy = "Full Coverage Auto Insurance"
        ElseIf InStr(1, value, "liab", vbTextCompare) > 0 Then
            CrownQuoteTidy = "Liability Auto Insurance"
        End If
        Exit Function
    End If

    If low = "vehicle_1_deductibles" Then
        Dim digits As String
        digits = CrownDigits(value)

        If Len(digits) > 0 Then CrownQuoteTidy = digits & "/" & digits
        Exit Function
    End If

    If low = "new_purchase" Then
        If StrComp(Trim$(value), "yes", vbTextCompare) = 0 Then
            CrownQuoteTidy = "Yes (I am purchasing this vehicle now or in the near future)"
        Else
            CrownQuoteTidy = "No"
        End If
    End If
End Function

' Two letters for a state written out in full. Anything already two
' letters, or not a state at all, comes back as it was.
Private Function CrownStateCode(ByVal value As String) As String
    Dim names As Variant, codes As Variant
    Dim i As Long

    CrownStateCode = Trim$(value)
    If Len(CrownStateCode) <= 2 Then Exit Function

    names = Array("alabama", "alaska", "arizona", "arkansas", "california", "colorado", _
        "connecticut", "delaware", "district of columbia", "florida", "georgia", "hawaii", _
        "idaho", "illinois", "indiana", "iowa", "kansas", "kentucky", "louisiana", "maine", _
        "maryland", "massachusetts", "michigan", "minnesota", "mississippi", "missouri", _
        "montana", "nebraska", "nevada", "new hampshire", "new jersey", "new mexico", _
        "new york", "north carolina", "north dakota", "ohio", "oklahoma", "oregon", _
        "pennsylvania", "rhode island", "south carolina", "south dakota", "tennessee", _
        "texas", "utah", "vermont", "virginia", "washington", "west virginia", "wisconsin", _
        "wyoming")

    codes = Array("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "HI", _
        "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", _
        "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", _
        "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY")

    For i = LBound(names) To UBound(names)
        If StrComp(CrownStateCode, CStr(names(i)), vbTextCompare) = 0 Then
            CrownStateCode = CStr(codes(i))
            Exit Function
        End If
    Next i
End Function

Private Function CrownQuoteSheetToExport(ByVal path As String, ByVal head As String) As String
    Dim pairs As Object, where As Object
    Dim columns() As String, sheetHead() As String, parts() As String
    Dim fileNumber As Integer
    Dim line As String, out As String, cell As String
    Dim i As Long, at As Long, rows As Long
    Dim line2() As String

    Set pairs = CrownQuotePairs()
    columns = CrownSplitCsvLine(head)

    ' Which column of the export each form field sits in.
    Set where = CreateObject("Scripting.Dictionary")
    where.CompareMode = 1

    For i = LBound(columns) To UBound(columns)
        If Not where.Exists(Trim$(columns(i))) Then where.Add Trim$(columns(i)), i
    Next i

    On Error GoTo Finished

    fileNumber = FreeFile
    Open path For Input As #fileNumber

    If EOF(fileNumber) Then GoTo Finished

    Line Input #fileNumber, line
    sheetHead = CrownSplitCsvLine(line)
    out = head

    Do Until EOF(fileNumber)
        Line Input #fileNumber, line

        If Len(Trim$(line)) > 0 Then
            parts = CrownSplitCsvLine(line)

            ReDim line2(LBound(columns) To UBound(columns))

            For i = LBound(sheetHead) To UBound(sheetHead)
                If i <= UBound(parts) Then
                    cell = Trim$(parts(i))

                    ' "none" is how that sheet writes an empty box.
                    If StrComp(cell, "none", vbTextCompare) = 0 Then cell = ""

                    If Len(cell) > 0 And pairs.Exists(Trim$(sheetHead(i))) Then
                        If where.Exists(pairs(Trim$(sheetHead(i)))) Then
                            at = where(pairs(Trim$(sheetHead(i))))
                            line2(at) = CrownQuoteTidy(sheetHead(i), cell)
                        End If
                    End If
                End If
            Next i

            out = out & vbLf & CrownJoinCsv(line2)
            rows = rows + 1
        End If
    Loop

Finished:
    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0

    If rows > 0 Then CrownQuoteSheetToExport = out
End Function

' One row back into csv, quoting anything that needs it.
Private Function CrownJoinCsv(ByRef cells() As String) As String
    Dim i As Long, out As String, cell As String

    For i = LBound(cells) To UBound(cells)
        cell = cells(i)

        If InStr(cell, """") > 0 Or InStr(cell, ",") > 0 _
           Or InStr(cell, vbLf) > 0 Or InStr(cell, vbCr) > 0 Then
            cell = """" & Replace(cell, """", """""") & """"
        End If

        If i > LBound(cells) Then out = out & ","
        out = out & cell
    Next i

    CrownJoinCsv = out
End Function


' ---------------------------------------------------------------------
' 3. Keeping them on Crown Quote List
' ---------------------------------------------------------------------
'
' Output is whatever was brought in just now - it gets cleared and rebuilt
' every time. Crown Quote List is the one that keeps them.
'
' Nothing is mapped here. Both sheets are built from the same list of
' headings by your own code, so a row can be copied straight across. That
' is checked before anything moves: if the two heading rows ever stop
' agreeing, this stops and says so rather than putting a date of birth
' under Gender.
'
' The website's quote number arrives in the "Submitter's User ID" column -
' that is where your m54 puts it - so a quote already on the list is
' recognised and left alone.
Private Function CrownAppendToQuoteList(ByRef added As Long, ByRef already As Long, _
                                        ByRef why As String) As Boolean
    Dim source As Worksheet, target As Worksheet
    Dim seen As Object
    Dim lastSourceRow As Long, lastTargetRow As Long, columns As Long
    Dim idColumn As Long
    Dim r As Long, c As Long
    Dim id As String

    added = 0
    already = 0
    why = ""

    On Error Resume Next
    Set source = ThisWorkbook.Sheets("Output")
    Set target = ThisWorkbook.Sheets("Crown Quote List")
    On Error GoTo 0

    If source Is Nothing Then
        why = "there is no Output sheet"
        Exit Function
    End If

    If target Is Nothing Then
        why = "there is no Crown Quote List sheet"
        Exit Function
    End If

    columns = source.Cells(1, source.Columns.Count).End(xlToLeft).Column

    If columns < 2 Then
        why = "Output has no headings on it"
        Exit Function
    End If

    lastSourceRow = CrownLastUsedRow(source, columns)

    If lastSourceRow < 2 Then
        CrownAppendToQuoteList = True              ' nothing came in; nothing to do
        Exit Function
    End If

    ' An empty list gets Output's headings, so the first run sets it up.
    If Len(Trim$(CStr(target.Cells(1, 1).value))) = 0 Then
        source.Range(source.Cells(1, 1), source.Cells(1, columns)).Copy _
            target.Cells(1, 1)
        Application.CutCopyMode = False
    End If

    If Not CrownSameHeadings(source, target, columns, why) Then Exit Function

    idColumn = CrownColumnNamed(target, columns, QUOTE_ID_HEADING)
    lastTargetRow = CrownLastUsedRow(target, columns)

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    If idColumn > 0 Then
        For r = 2 To lastTargetRow
            id = Trim$(CStr(target.Cells(r, idColumn).value))
            If Len(id) > 0 And Not seen.Exists(id) Then seen.Add id, True
        Next r
    End If

    For r = 2 To lastSourceRow
        id = ""
        If idColumn > 0 Then id = Trim$(CStr(source.Cells(r, idColumn).value))

        If Len(id) > 0 And seen.Exists(id) Then
            already = already + 1
        Else
            lastTargetRow = lastTargetRow + 1

            source.Range(source.Cells(r, 1), source.Cells(r, columns)).Copy
            target.Cells(lastTargetRow, 1).PasteSpecial xlPasteValues
            Application.CutCopyMode = False

            If Len(id) > 0 Then seen.Add id, True
            added = added + 1
        End If
    Next r

    CrownAppendToQuoteList = True
End Function

' The last row with anything on it, looked for across the whole width.
' Column A alone is not safe - a quote with no source on it leaves A empty.
Private Function CrownLastUsedRow(ByVal ws As Worksheet, ByVal columns As Long) As Long
    Dim found As Range

    On Error Resume Next
    Set found = ws.Range(ws.Cells(1, 1), ws.Cells(ws.Rows.Count, columns)) _
        .Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If Not found Is Nothing Then CrownLastUsedRow = found.Row
End Function

Private Function CrownSameHeadings(ByVal source As Worksheet, ByVal target As Worksheet, _
                                   ByVal columns As Long, ByRef why As String) As Boolean
    Dim c As Long
    Dim a As String, b As String

    For c = 1 To columns
        a = Trim$(CStr(source.Cells(1, c).value))
        b = Trim$(CStr(target.Cells(1, c).value))

        If StrComp(a, b, vbTextCompare) <> 0 Then
            why = "the headings on Output and Crown Quote List stopped agreeing at column " & c & _
                  " - Output says """ & a & """ and the list says """ & b & """." & vbCrLf & _
                  "Nothing has been copied. Rebuild the list with your usual import and try again."
            Exit Function
        End If
    Next c

    CrownSameHeadings = True
End Function

Private Function CrownColumnNamed(ByVal ws As Worksheet, ByVal columns As Long, _
                                  ByVal heading As String) As Long
    Dim c As Long

    For c = 1 To columns
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), heading, vbTextCompare) = 0 Then
            CrownColumnNamed = c
            Exit Function
        End If
    Next c
End Function

' Everything after the tool has filled Output: put the new ones on the
' list, and say what happened.
Private Function CrownKeepThem() As String
    Dim added As Long, already As Long
    Dim why As String

    If Not CrownAppendToQuoteList(added, already, why) Then
        CrownKeepThem = "NOT added to Crown Quote List - " & why
        Exit Function
    End If

    CrownKeepThem = added & " added to Crown Quote List"

    If already > 0 Then
        CrownKeepThem = CrownKeepThem & ", " & already & " were already on it"
    End If

    CrownKeepThem = CrownKeepThem & "."
End Function

' ---------------------------------------------------------------------
' 4. One quote onto Edit Data
' ---------------------------------------------------------------------
'
' Stand on the row you want on Crown Quote List and run this. That
' customer goes onto Edit Data and you can quote them.
'
' RetrieveDataByRowNumber reads a row number out of Edit Data C5 and takes
' that row off Output, so this puts the heading row and the one quote on
' Output and points C5 at row 2. That is the same state you were getting
' to by clearing Output down to a single row by hand.

Public Sub CrownQuoteToEditData()
    Dim list As Worksheet, output As Worksheet, edit As Worksheet
    Dim columns As Long, lastRow As Long, wanted As Long
    Dim answer As String
    Dim who As String

    On Error Resume Next
    Set list = ThisWorkbook.Sheets("Crown Quote List")
    Set output = ThisWorkbook.Sheets("Output")
    Set edit = ThisWorkbook.Sheets("Edit Data")
    On Error GoTo 0

    If list Is Nothing Or output Is Nothing Or edit Is Nothing Then
        MsgBox "This needs the Crown Quote List, Output and Edit Data sheets.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    columns = list.Cells(1, list.Columns.Count).End(xlToLeft).Column
    lastRow = CrownLastUsedRow(list, columns)

    If columns < 2 Or lastRow < 2 Then
        MsgBox "There are no quotes on Crown Quote List yet." & vbCrLf & vbCrLf & _
               "Run CrownFetchNewQuotes or CrownImportQuoteFile first.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' The row the cursor is on, when the cursor is on that sheet. Otherwise
    ' ask, because guessing which customer he meant is not a thing to do.
    wanted = 0

    ' Guarded: Selection is not always a range - click a picture or a chart
    ' and asking it for a row stops the macro with an error instead of
    ' asking which customer was meant.
    If ActiveSheet Is list Then
        On Error Resume Next

        If TypeName(Selection) = "Range" Then
            If Selection.Row >= 2 And Selection.Row <= lastRow Then wanted = Selection.Row
        End If

        On Error GoTo 0
    End If

    If wanted = 0 Then
        answer = InputBox("Which row of Crown Quote List?" & vbCrLf & vbCrLf & _
                          "Quotes are on rows 2 to " & lastRow & "." & vbCrLf & _
                          "(Or close this, click the row you want on that sheet, and run it again.)", _
                          "Crown Superior")

        If Len(Trim$(answer)) = 0 Then Exit Sub
        If Not IsNumeric(answer) Then
            MsgBox "That is not a row number.", vbExclamation, "Crown Superior"
            Exit Sub
        End If

        wanted = CLng(answer)
    End If

    If wanted < 2 Or wanted > lastRow Then
        MsgBox "Row " & wanted & " is not one of the quotes. They are on rows 2 to " & _
               lastRow & ".", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    who = CrownWhoIsOnRow(list, columns, wanted)

    Application.ScreenUpdating = False

    output.Cells.Clear
    list.Range(list.Cells(1, 1), list.Cells(1, columns)).Copy output.Cells(1, 1)
    list.Range(list.Cells(wanted, 1), list.Cells(wanted, columns)).Copy
    output.Cells(2, 1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    edit.Range("C5").value = 2

    Application.ScreenUpdating = True

    RetrieveDataByRowNumber

    On Error Resume Next
    edit.Activate
    On Error GoTo 0

    MsgBox "On Edit Data now:" & vbCrLf & vbCrLf & "   " & who & vbCrLf & vbCrLf & _
           "Output is holding that one quote. Crown Quote List still has all of them.", _
           vbInformation, "Crown Superior"
End Sub

' Whose quote is on this row, for the message - so he can see at a glance
' that the right customer landed.
Private Function CrownWhoIsOnRow(ByVal ws As Worksheet, ByVal columns As Long, _
                                 ByVal row As Long) As String
    Dim first As Long, last As Long, email As Long

    first = CrownColumnNamed(ws, columns, "First Name")
    last = CrownColumnNamed(ws, columns, "Last Name")
    email = CrownColumnNamed(ws, columns, "E-mail")

    If email = 0 Then email = CrownColumnNamed(ws, columns, "Email")

    If first > 0 Then CrownWhoIsOnRow = Trim$(CStr(ws.Cells(row, first).value))
    If last > 0 Then CrownWhoIsOnRow = Trim$(CrownWhoIsOnRow & " " & CStr(ws.Cells(row, last).value))

    If email > 0 And Len(Trim$(CStr(ws.Cells(row, email).value))) > 0 Then
        CrownWhoIsOnRow = Trim$(CrownWhoIsOnRow & "   " & CStr(ws.Cells(row, email).value))
    End If

    If Len(Trim$(CrownWhoIsOnRow)) = 0 Then CrownWhoIsOnRow = "row " & row
End Function

' ---------------------------------------------------------------------
' 5. Handing it to the tool
' ---------------------------------------------------------------------

Private Function CrownWriteTemp(ByVal csv As String, ByVal prefix As String) As String
    Dim stream As Object
    Dim path As String

    path = Environ$("TEMP") & "\" & prefix & Format$(Now, "yyyymmdd-hhnnss") & ".csv"

    ' Written through a stream rather than Print #, so a name with an
    ' accent in it arrives as the name and not as rubbish.
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText csv
    stream.SaveToFile path, 2
    stream.Close

    CrownWriteTemp = path
End Function

' Your own import, unchanged. ImportquoteToAlloutputsheet takes its file
' from the ImportedFilePath global rather than from its argument, which is
' what lets this point it at a file nobody picked by hand.
Private Sub CrownHandToTool(ByVal path As String)
    ImportedFilePath = path

    ImportquoteToAlloutputsheet ThisWorkbook.Sheets(1)

    On Error Resume Next
    ThisWorkbook.Sheets("Edit Data").Activate
    On Error GoTo 0

    RetrieveDataByRowNumber
End Sub
