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


' ---------------------------------------------------------------------
' 1. The website's quotes, only what is new
' ---------------------------------------------------------------------

Public Sub CrownFetchNewQuotes()
    Dim csv As String, path As String
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

    If highest > since Then CrownSaveLastQuoteId highest

    MsgBox brought & " new quote(s) brought in and put on Output." & vbCrLf & _
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
        MsgBox "Brought in. Edit Data is showing the row you had open.", _
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

    MsgBox "Brought in and put on Output." & vbCrLf & _
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
' 4. Handing it to the tool
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
