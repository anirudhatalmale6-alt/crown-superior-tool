Attribute VB_Name = "CrownPayments"
Option Explicit

' =====================================================================
' Crown Superior - payment information, both directions
' ---------------------------------------------------------------------
' Needs the CrownAPI module to be imported as well; the key is shared.
'
' Two things live here.
'
'   CrownUploadPaymentsDue
'       Takes the UploadPaymentdue sheet exactly as it is today and writes
'       every row straight onto the website. No browser, no logging in,
'       no searching the policy list, no form validation to argue with.
'       Works for every carrier, not just one.
'
'   CrownUaigPaymentDue
'       Works through OUR United Auto policies - the list comes from the
'       website, not from whichever UAIG report happens to be in front of
'       it - looks each policy up in Policy Inquiry, reads what is owed,
'       and writes it to the CrownPayments sheet AND back to the website.
'
' The website is the record. The Google Sheet reads the website on its
' own about once an hour, so anything written here reaches the phone
' system without anybody exporting or uploading anything.
' =====================================================================

' The columns on UploadPaymentdue, as it already is.
Private Const COL_POLICY As Long = 1     ' A  Policy Number
Private Const COL_SITE As Long = 2       ' B  Website (the carrier)
Private Const COL_FIRST As Long = 3      ' C  First Name
Private Const COL_LAST As Long = 4       ' D  Last Name
Private Const COL_CANCEL As Long = 5     ' E  Cancel Date
Private Const COL_AMOUNT As Long = 6     ' F  Amount Due
Private Const COL_DUE As Long = 7        ' G  Due Date
Private Const COL_STATUS As Long = 8     ' H  Status


' ---------------------------------------------------------------------
' Our policies, from the website
' ---------------------------------------------------------------------

' Every policy the website holds, as a lookup: the digits of the policy
' number -> the record number to write back to.
'
' Digits only, because the same policy is written "GAI 212619",
' "GAI212633" and "GAI - 210184" in different records, and the carriers
' ask for it as a number anyway.
'
' carrier is optional. "United Auto", "UAIG" and "United" all mean the
' same insurer to the website.
Public Function CrownPolicyRows(Optional ByVal carrier As String = "") As Collection
    Dim csv As String, body As String, lines() As String, parts() As String
    Dim rows As New Collection
    Dim i As Long

    ' Built with a plain If, not IIf: IIf works out BOTH of its answers
    ' before choosing one, so the carrier would be encoded even when there
    ' is no carrier to encode.
    body = CrownBody("policylist")
    If Len(carrier) > 0 Then body = body & "&carrier=" & CrownEnc(carrier)

    csv = CrownPost(body)

    If Len(csv) = 0 Then
        MsgBox "The website did not answer. Check your internet, then run CrownTestConnection.", _
               vbExclamation, "Crown Superior"
        Set CrownPolicyRows = rows
        Exit Function
    End If

    If Left$(csv, 5) = "error" Then
        MsgBox "The website answered: " & csv, vbExclamation, "Crown Superior"
        Set CrownPolicyRows = rows
        Exit Function
    End If

    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)

    For i = 1 To UBound(lines)                    ' 0 is the header
        If Len(Trim$(lines(i))) > 0 Then
            parts = CrownSplitCsvLine(lines(i))

            ' 0 record number, 2 policy number, 3 carrier, 4 first, 5 last
            If UBound(parts) >= 5 Then rows.Add parts
        End If
    Next i

    Set CrownPolicyRows = rows
End Function

' The policies, by the digits in their policy number.
Public Function CrownPolicyIndex(Optional ByVal carrier As String = "") As Object
    Dim rows As Collection, parts As Variant
    Dim index As Object
    Dim digits As String

    Set index = CreateObject("Scripting.Dictionary")
    Set rows = CrownPolicyRows(carrier)

    For Each parts In rows
        digits = CrownDigits(parts(2))

        ' The list comes back newest first, so the first one seen for a
        ' number is the newest record with it.
        If Len(digits) > 0 And Not index.Exists(digits) Then
            index.Add digits, parts(0)
        End If
    Next parts

    Set CrownPolicyIndex = index
End Function

' The policy number in a cell, as it was really typed.
'
' A long one like 100332072503 is held by Excel as a number and comes
' back from CStr as "1.00332072503E+11", which has no policy number in
' it at all. Formatting it as a plain integer gets the digits back.
' Takes the cell itself, not its contents: passing a Range into a Variant
' hands over the value, and a value has no .Value to read.
Public Function CrownCellText(ByVal cell As Range) As String
    Dim value As Variant

    value = cell.value

    If IsEmpty(value) Then Exit Function

    If IsNumeric(value) And Not IsDate(value) Then
        If value = Int(value) And Abs(value) < 1E+15 Then
            CrownCellText = Format$(value, "0")
            Exit Function
        End If
    End If

    CrownCellText = Trim$(CStr(value))
End Function

' The digits in a policy number, and nothing else.
Public Function CrownDigits(ByVal text As String) As String
    Dim i As Long, ch As String, out As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i

    CrownDigits = out
End Function

' The website wants MM-DD-YYYY. Anything it cannot read is sent through
' untouched rather than turned into today by accident.
Private Function CrownDate(ByVal value As Variant) As String
    If Trim$(CStr(value)) = "" Then Exit Function

    If IsDate(value) Then
        CrownDate = Format$(CDate(value), "mm-dd-yyyy")
    Else
        CrownDate = Trim$(CStr(value))
    End If
End Function


' ---------------------------------------------------------------------
' 1. The upload sheet, straight onto the website
' ---------------------------------------------------------------------

Public Sub CrownUploadPaymentsDue()
    Dim wsUpload As Worksheet, wsLog As Worksheet
    Dim rows As Collection, parts As Variant
    Dim byNumber As Object, byName As Object
    Dim lastRow As Long, i As Long, logRow As Long
    Dim policyText As String, digits As String, tail As String
    Dim recordId As String, matchedOn As String, note As String
    Dim carrier As String, nameKey As String, answer As String
    Dim done As Long, missing As Long, refused As Long
    Dim names As Variant, values As Variant
    Dim candidates As Variant, j As Long

    On Error Resume Next
    Set wsUpload = ThisWorkbook.Worksheets("UploadPaymentdue")
    On Error GoTo 0

    If wsUpload Is Nothing Then
        MsgBox "There is no UploadPaymentdue sheet in this workbook.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set rows = CrownPolicyRows()
    If rows.Count = 0 Then Exit Sub

    ' Two ways of finding a record: by the digits in the policy number, and
    ' by the customer's name. The numbers on this sheet are the carrier's,
    ' and the carrier issues a new one on a rewrite, so a policy we hold can
    ' easily be under a different number here.
    Set byNumber = CreateObject("Scripting.Dictionary")
    Set byName = CreateObject("Scripting.Dictionary")

    For Each parts In rows
        digits = CrownDigits(parts(2))
        If Len(digits) > 0 And Not byNumber.Exists(digits) Then byNumber.Add digits, parts

        nameKey = UCase$(Trim$(parts(4)) & "|" & Trim$(parts(5)))

        If Len(Replace(nameKey, "|", "")) > 0 Then
            If byName.Exists(nameKey) Then
                candidates = byName(nameKey)
                ReDim Preserve candidates(UBound(candidates) + 1)
                candidates(UBound(candidates)) = parts
                byName(nameKey) = candidates
            Else
                byName.Add nameKey, Array(parts)
            End If
        End If
    Next parts

    Set wsLog = CrownSheet("CrownUploadLog")
    wsLog.Cells.ClearContents
    wsLog.Columns(1).NumberFormat = "@"          ' policy numbers are text, not sums
    wsLog.Range("A1:K1").value = Array("Policy Number", "Website", "First Name", "Last Name", _
                                       "Cancel Date", "Amount Due", "Due Date", "Status", _
                                       "When", "Matched on", "What happened")
    logRow = 2

    lastRow = wsUpload.Cells(wsUpload.rows.Count, COL_POLICY).End(xlUp).row

    For i = 2 To lastRow
        policyText = CrownCellText(wsUpload.Cells(i, COL_POLICY))

        If Len(Trim$(policyText)) > 0 Then
            digits = CrownDigits(policyText)
            carrier = Trim$(CStr(wsUpload.Cells(i, COL_SITE).value))
            nameKey = UCase$(Trim$(CStr(wsUpload.Cells(i, COL_FIRST).value)) & "|" & _
                             Trim$(CStr(wsUpload.Cells(i, COL_LAST).value)))
            recordId = ""
            matchedOn = ""
            note = ""

            ' 1. the whole policy number
            If byNumber.Exists(digits) Then
                parts = byNumber(digits)
                recordId = parts(0)
                matchedOn = "policy number"
            End If

            ' 2. the part after the last hyphen, which is what the old macro
            '    searched on for United Auto
            If Len(recordId) = 0 And InStr(policyText, "-") > 0 Then
                tail = CrownDigits(Mid$(policyText, InStrRev(policyText, "-") + 1))

                If Len(tail) > 0 And byNumber.Exists(tail) Then
                    parts = byNumber(tail)
                    recordId = parts(0)
                    matchedOn = "end of the policy number"
                End If
            End If

            ' 3. the customer's name, and only when it points at one policy
            If Len(recordId) = 0 And byName.Exists(nameKey) Then
                candidates = byName(nameKey)
                note = ""

                For j = LBound(candidates) To UBound(candidates)
                    parts = candidates(j)

                    If Len(carrier) = 0 Or CrownSameCarrier(carrier, CStr(parts(3))) Then
                        If Len(recordId) = 0 Then
                            recordId = parts(0)
                            matchedOn = "name"
                            If Len(carrier) > 0 Then matchedOn = "name and carrier"
                        Else
                            ' more than one - too risky to pick
                            recordId = ""
                            matchedOn = ""
                            note = "That name has more than one policy here: "
                            Exit For
                        End If
                    End If
                Next j

                If Len(note) > 0 Then
                    For j = LBound(candidates) To UBound(candidates)
                        parts = candidates(j)
                        note = note & parts(0) & " " & parts(2) & " (" & parts(3) & ")  "
                    Next j
                End If
            End If

            wsLog.Cells(logRow, 1).value = policyText
            wsLog.Cells(logRow, 2).value = wsUpload.Cells(i, COL_SITE).value
            wsLog.Cells(logRow, 3).value = wsUpload.Cells(i, COL_FIRST).value
            wsLog.Cells(logRow, 4).value = wsUpload.Cells(i, COL_LAST).value
            wsLog.Cells(logRow, 5).value = wsUpload.Cells(i, COL_CANCEL).value
            wsLog.Cells(logRow, 6).value = wsUpload.Cells(i, COL_AMOUNT).value
            wsLog.Cells(logRow, 7).value = wsUpload.Cells(i, COL_DUE).value
            wsLog.Cells(logRow, 8).value = wsUpload.Cells(i, COL_STATUS).value
            wsLog.Cells(logRow, 9).value = Now
            wsLog.Cells(logRow, 10).value = matchedOn

            If Len(recordId) = 0 Then
                If Len(note) = 0 Then note = "No policy on the website with that number or that name"
                wsLog.Cells(logRow, 11).value = note
                missing = missing + 1
            Else
                names = Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                              "Status_", "Update_due_dates", "Last_Updated")
                values = Array(Trim$(CStr(wsUpload.Cells(i, COL_AMOUNT).value)), _
                               CrownDate(wsUpload.Cells(i, COL_DUE).value), _
                               CrownDate(wsUpload.Cells(i, COL_CANCEL).value), _
                               Trim$(CStr(wsUpload.Cells(i, COL_STATUS).value)), _
                               "Yes", _
                               Format$(Now, "mm-dd-yyyy hh:mm AM/PM"))

                answer = CrownUpdate(CROWN_FORM_POLICY, CLng(recordId), names, values)

                If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                    wsLog.Cells(logRow, 11).value = "Written to record " & recordId
                    done = done + 1
                ElseIf Len(Trim$(answer)) = 0 Then
                    wsLog.Cells(logRow, 11).value = "The website did not answer - try this row again"
                    refused = refused + 1
                Else
                    wsLog.Cells(logRow, 11).value = "Refused: " & answer
                    refused = refused + 1
                End If
            End If

            logRow = logRow + 1
        End If
    Next i

    wsLog.Columns("A:K").AutoFit
    wsLog.Activate

    MsgBox done & " policy record(s) updated on the website." & vbCrLf & _
           missing & " could not be matched to a policy there." & vbCrLf & _
           refused & " were refused or did not go through." & vbCrLf & vbCrLf & _
           "The CrownUploadLog sheet says what happened to every row, and how each " & _
           "one was matched." & vbCrLf & vbCrLf & _
           "The Google Sheet picks the changes up within the hour on its own.", _
           vbInformation, "Crown Superior"
End Sub

' Two ways of writing the same insurer. Loose on purpose: the sheet says
' "United Au" or "UAIG" where the website says "United Auto".
Public Function CrownSameCarrier(ByVal a As String, ByVal b As String) As Boolean
    Dim x As String, y As String

    x = UCase$(Trim$(a))
    y = UCase$(Trim$(b))

    If Len(x) = 0 Or Len(y) = 0 Then
        CrownSameCarrier = True
        Exit Function
    End If

    If x = y Then
        CrownSameCarrier = True
        Exit Function
    End If

    If InStr(1, y, x, vbTextCompare) = 1 Or InStr(1, x, y, vbTextCompare) = 1 Then
        CrownSameCarrier = True
        Exit Function
    End If

    If (InStr(1, x, "UNITED", vbTextCompare) > 0 Or InStr(1, x, "UAIG", vbTextCompare) > 0) _
       And InStr(1, y, "UNITED", vbTextCompare) > 0 Then
        CrownSameCarrier = True
        Exit Function
    End If

    If (InStr(1, x, "VERVE", vbTextCompare) > 0 Or InStr(1, x, "TRISURA", vbTextCompare) > 0) _
       And (InStr(1, y, "VERVE", vbTextCompare) > 0 Or InStr(1, y, "TRISURA", vbTextCompare) > 0) Then
        CrownSameCarrier = True
        Exit Function
    End If

    CrownSameCarrier = False
End Function


' ---------------------------------------------------------------------
' 2. United Auto, policy by policy, from OUR list
' ---------------------------------------------------------------------

Public Sub CrownUaigPaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim By As New Selenium.By
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Worksheet
    Dim index As Object, key As Variant
    Dim policyDigits As String, recordId As String
    Dim dueAmount As String, dueDate As String, cancelDate As String, policyStatus As String
    Dim answer As String, jsScript As String
    Dim row As Long, done As Long, blank As Long

    Set wsInput = ThisWorkbook.Worksheets("Input")

    ' Our United Auto policies, newest first. If the website has none, or
    ' cannot be reached, stop here rather than open a browser for nothing.
    Set index = CrownPolicyIndex("United Auto")
    If index.Count = 0 Then
        MsgBox "No United Auto policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set ws = CrownSheet("CrownPayments")
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"       ' a long policy number is not a sum
    ws.Range("A1:H1").value = Array("Record number", "Policy number", "Amount due", "Due date", _
                                    "Cancel date", "Policy status", "When", "Written back")
    row = 2

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--force-device-scale-factor=0.70"
    drv.Start
    On Error GoTo 0

    drv.Get wsInput.Range("URL_26").value
    EnterData drv, Keys, "tbxUserID", wsInput.Range("USER_26").value, "ID"
    EnterData drv, Keys, "tbxPassword", wsInput.Range("PASS_26").value, "ID"
    drv.Keyboard.SendKeys Keys.Enter

    LoopElementUntilFoundBYXPATH drv, "//a[normalize-space(text())='Quote']"
    LoopElementUntilFound drv, "rpthref"

    ' Straight to Policy Inquiry. No reports, no downloads, no tabs - the
    ' list of what to look up came from us, so nothing here depends on
    ' which report UAIG happens to show first.
    ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
    ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"

    For Each key In index.Keys
        policyDigits = CStr(key)
        recordId = index(key)

        dueAmount = ""
        dueDate = ""
        cancelDate = ""
        policyStatus = ""

        On Error Resume Next

        LoopElementUntilFound drv, "tbxPolicyNo"
        drv.FindElementById("tbxPolicyNo").Clear
        drv.FindElementById("tbxPolicyNo").SendKeys policyDigits
        ClickElement drv, "btnSubmitPol1", "ID"
        drv.Wait 2000

        If drv.IsElementPresent(By.ID("CURAMTDUE")) Then
            dueAmount = drv.FindElementById("CURAMTDUE").value
        End If

        policyStatus = drv.FindElementByXPath( _
            "//td[.//font[contains(.,'Policy Status')]]/following-sibling::td//strong").text

        If Trim$(policyStatus) = "" Then
            policyStatus = drv.FindElementByXPath( _
                "//font[contains(text(),'Policy Status')]/ancestor::td/following-sibling::td//font").text
        End If

        policyStatus = Replace(policyStatus, vbLf, " ")
        policyStatus = Application.WorksheetFunction.Trim(policyStatus)

        ' The last starred instalment line is the one that is due.
        jsScript = "var rows = document.evaluate(" & _
                   """" & "//tr[td//font[contains(.,'Installment') and contains(.,'*')]]" & """" & _
                   ", document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);" & _
                   "if (rows.snapshotLength > 0) {" & _
                   "  var lastRow = rows.snapshotItem(rows.snapshotLength - 1);" & _
                   "  var cells = lastRow.getElementsByTagName('td');" & _
                   "  if (cells.length >= 7) { return cells[6].innerText.trim(); }" & _
                   "} return '';"
        dueDate = drv.ExecuteScript(jsScript)

        cancelDate = drv.ExecuteScript( _
            "var c = document.querySelector('#spnShowSchedule span.pay-dt');" & _
            "return c ? c.innerText.trim() : '';")

        On Error GoTo 0

        ws.Cells(row, 1).value = recordId
        ws.Cells(row, 2).value = policyDigits
        ws.Cells(row, 3).value = dueAmount
        ws.Cells(row, 4).value = dueDate
        ws.Cells(row, 5).value = cancelDate
        ws.Cells(row, 6).value = policyStatus
        ws.Cells(row, 7).value = Now

        ' Nothing at all came back: the policy is not at this carrier, or
        ' the page did not load. Do not write emptiness over what the
        ' website already holds.
        If Len(Trim$(dueAmount)) = 0 And Len(Trim$(dueDate)) = 0 _
           And Len(Trim$(cancelDate)) = 0 And Len(Trim$(policyStatus)) = 0 Then
            ws.Cells(row, 8).value = "nothing found - left alone"
            blank = blank + 1
        Else
            answer = CrownUpdate(CROWN_FORM_POLICY, CLng(recordId), _
                Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                      "Status_", "Update_due_dates", "Last_Updated"), _
                Array(dueAmount, CrownDate(dueDate), CrownDate(cancelDate), _
                      policyStatus, "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM")))

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 8).value = "written"
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 8).value = "website did not answer - run again for this one"
            Else
                ws.Cells(row, 8).value = "refused: " & answer
            End If
        End If

        row = row + 1
        drv.GoBack
        drv.Wait 1000
    Next key

    ws.Activate

    MsgBox index.Count & " United Auto policies looked up." & vbCrLf & _
           done & " written back to the website." & vbCrLf & _
           blank & " had nothing to read and were left as they were." & vbCrLf & vbCrLf & _
           "See the CrownPayments sheet for the detail. The Google Sheet " & _
           "picks the changes up within the hour.", vbInformation, "Crown Superior"
End Sub
