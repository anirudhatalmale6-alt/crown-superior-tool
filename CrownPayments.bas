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
' A term cancelled or expired longer ago than this is of no use to anyone.
Private Const UAIG_STALE_DAYS As Long = 60

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

' ---------------------------------------------------------------------
' 2. United Auto, one policy at a time - renewals included
' ---------------------------------------------------------------------
'
' Policy Inquiry does not always go straight to a policy. When a policy
' has been renewed, United holds more than one term under it and answers
' with a list instead - "Policy Inquiry Result", a line per term, each
' with its own number, its own effective date and its own status. The
' newer term is given a new number: usually the old one with 10 in front
' of it, sometimes a number with nothing in common with it at all.
'
' That list is why four policies came back empty last time. This reads
' it, takes the term with the newest effective date - ignoring anything
' cancelled or expired more than two months ago - and opens that one.
'
' Where the newer number is not on our website at all, the policy has
' been renewed without us knowing. Those are gathered up and offered at
' the end as new policy records: a copy of the one we hold, with the new
' number, the new dates and "Renewal" as the description. Nothing is
' created without being listed and agreed to first.

' A date the way United writes it, as something that can be compared.
'
' It writes them two ways on the same screen depending on where you are:
' 02/23/2026 on one, 2026-02-23 on another. Four digits at the front is
' the year, so month and day follow it; anything else is month first, the
' American way. Zero for anything that is not a date at all, so a blank
' or an "N/A" sorts below every real date instead of above them.
Public Function CrownUsDate(ByVal text As String) As Double
    Dim parts() As String
    Dim d As Long, m As Long, y As Long

    text = Trim$(Replace(Replace(text, "-", "/"), ".", "/"))
    If Len(text) = 0 Then Exit Function

    parts = Split(text, "/")
    If UBound(parts) <> 2 Then Exit Function

    If Len(Trim$(parts(0))) = 4 Then
        y = Val(parts(0))
        m = Val(parts(1))
        d = Val(parts(2))
    Else
        m = Val(parts(0))
        d = Val(parts(1))
        y = Val(parts(2))
    End If

    If y < 100 Then y = 2000 + y
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    If y < 1900 Or y > 2200 Then Exit Function

    On Error Resume Next
    CrownUsDate = CDbl(DateSerial(y, m, d))
End Function

' Which column holds what. Found by its heading rather than by counting,
' because the order of them is United's business, not ours.
Private Function CrownColumnOf(ByRef header() As String, ByVal wanted As String) As Long
    Dim i As Long

    CrownColumnOf = -1

    For i = LBound(header) To UBound(header)
        If InStr(1, header(i), wanted, vbTextCompare) > 0 Then
            CrownColumnOf = i
            Exit Function
        End If
    Next i
End Function

' The result table as plain text: rows separated by ~, cells by |.
'
' The page is several tables deep inside one another and every one of
' them contains the words we are looking for, so the smallest one that
' does is the real table.
Private Function CrownUaigListScript() As String
    CrownUaigListScript = _
        "var best=null,bl=0,tabs=document.getElementsByTagName('table');" & _
        "for(var i=0;i<tabs.length;i++){var t=tabs[i],x=t.innerText||'';" & _
        "if(x.indexOf('Term Effective Date')<0)continue;" & _
        "if(best===null||x.length<bl){best=t;bl=x.length;}}" & _
        "if(!best)return '';var out=[];" & _
        "for(var r=0;r<best.rows.length;r++){var cs=best.rows[r].cells,line=[];" & _
        "for(var c=0;c<cs.length;c++){line.push((cs[c].innerText||'')" & _
        ".replace(/[|~]/g,' ').replace(/\s+/g,' ').replace(/^ /,'').replace(/ $/,''));}" & _
        "out.push(line.join('|'));}return out.join('~');"
End Function

' Open one line of that list, by clicking its policy number where it sits
' rather than searching again - searching again only brings the list back.
Private Function CrownUaigClickScript(ByVal rowIndex As Long) As String
    CrownUaigClickScript = _
        "var best=null,bl=0,tabs=document.getElementsByTagName('table');" & _
        "for(var i=0;i<tabs.length;i++){var t=tabs[i],x=t.innerText||'';" & _
        "if(x.indexOf('Term Effective Date')<0)continue;" & _
        "if(best===null||x.length<bl){best=t;bl=x.length;}}" & _
        "if(!best)return 'no table';var row=best.rows[" & rowIndex & "];" & _
        "if(!row)return 'no row';var a=row.getElementsByTagName('a')[0];" & _
        "if(!a)return 'no link';a.click();return 'clicked';"
End Function

' Are we looking at one policy, or still at the list?
Private Function CrownUaigOnPolicy(ByVal drv As Object) As Boolean
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "var t=document.body?(document.body.innerText||''):'';" & _
        "return (document.getElementById('CURAMTDUE')" & _
        "||t.indexOf('Policy Status')>=0)?'1':'0';")
    On Error GoTo 0

    CrownUaigOnPolicy = (answer = "1")
End Function

' The line of the list to act on.
'
' "policies with a newer effective date in almost every case will be the
' policy to consider" - so the newest effective date wins, and anything
' cancelled or expired more than two months ago is passed over.
'
' Answers the row number to click, or 0 for nothing usable.
Private Function CrownUaigPickRow(ByVal listText As String, _
                                  ByRef bestNumber As String, ByRef bestEff As String, _
                                  ByRef bestExp As String, ByRef bestStatus As String, _
                                  ByRef note As String) As Long
    Dim rows() As String, header() As String, cells() As String
    Dim colPolicy As Long, colEff As Long, colExp As Long, colStatus As Long
    Dim r As Long, best As Long, skipped As Long
    Dim eff As Double, expiry As Double, bestEffOn As Double
    Dim status As String

    bestNumber = ""
    bestEff = ""
    bestExp = ""
    bestStatus = ""

    rows = Split(listText, "~")
    If UBound(rows) < 1 Then Exit Function          ' a heading and nothing under it

    header = Split(rows(0), "|")
    colPolicy = CrownColumnOf(header, "Policy No")
    colEff = CrownColumnOf(header, "Term Effective")
    colExp = CrownColumnOf(header, "Term Expiration")
    colStatus = CrownColumnOf(header, "Status")

    If colPolicy < 0 Or colEff < 0 Then
        note = "the list did not have the columns expected"
        Exit Function
    End If

    For r = 1 To UBound(rows)
        cells = Split(rows(r), "|")

        If UBound(cells) >= colPolicy And UBound(cells) >= colEff Then
            status = ""
            If colStatus >= 0 And UBound(cells) >= colStatus Then status = cells(colStatus)

            eff = CrownUsDate(cells(colEff))
            expiry = 0
            If colExp >= 0 And UBound(cells) >= colExp Then expiry = CrownUsDate(cells(colExp))

            ' Long gone. Skipped whatever its dates say about being newest.
            If (InStr(1, status, "cancel", vbTextCompare) > 0 _
                Or InStr(1, status, "expire", vbTextCompare) > 0) _
               And expiry > 0 And expiry < CDbl(Date) - UAIG_STALE_DAYS Then
                skipped = skipped + 1
            ElseIf Len(Trim$(cells(colPolicy))) > 0 And eff >= bestEffOn Then
                bestEffOn = eff
                best = r
                bestNumber = Trim$(cells(colPolicy))
                bestEff = Trim$(cells(colEff))
                bestStatus = Trim$(status)
                If colExp >= 0 And UBound(cells) >= colExp Then bestExp = Trim$(cells(colExp))
            End If
        End If
    Next r

    If best > 0 Then
        note = UBound(rows) & " terms on United; took " & bestNumber & " effective " & bestEff
        If UBound(rows) = 1 Then note = "one term on United; took " & bestNumber & " effective " & bestEff
        If skipped > 0 Then note = note & "; " & skipped & " long expired"
    ElseIf skipped > 0 Then
        note = "all " & skipped & " terms cancelled or expired more than two months ago"
    End If

    CrownUaigPickRow = best
End Function


Public Sub CrownUaigPaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim By As New Selenium.By
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Worksheet
    Dim policies As Collection, parts As Variant
    Dim index As Object, names As Object
    Dim renewals As Collection
    Dim policyDigits As String, recordId As String, siteNumber As String
    Dim liveNumber As String, liveDigits As String
    Dim termEff As String, termExp As String, listStatus As String
    Dim dueAmount As String, dueDate As String, cancelDate As String
    Dim policyStatus As String, detailStatus As String
    Dim listText As String, note As String, answer As String, jsScript As String
    Dim row As Long, done As Long, blank As Long, pick As Long, madeCount As Long
    Dim i As Long

    Set wsInput = ThisWorkbook.Worksheets("Input")

    ' Our United Auto policies, newest first. If the website has none, or
    ' cannot be reached, stop here rather than open a browser for nothing.
    Set policies = CrownPolicyRows("United Auto")

    If policies.Count = 0 Then
        MsgBox "No United Auto policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' The digits of every policy number we hold, so a number United gives
    ' back can be recognised as one we already have.
    Set index = CreateObject("Scripting.Dictionary")
    Set names = CreateObject("Scripting.Dictionary")

    For Each parts In policies
        policyDigits = CrownDigits(parts(2))

        If Len(policyDigits) > 0 And Not index.Exists(policyDigits) Then
            index.Add policyDigits, parts(0)
            names.Add policyDigits, Trim$(parts(4) & " " & parts(5))
        End If
    Next parts

    Set renewals = New Collection

    Set ws = CrownSheet("CrownPayments")
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"       ' a long policy number is not a sum
    ws.Columns(3).NumberFormat = "@"
    ws.Range("A1:K1").value = Array("Record number", "Policy on our site", "Policy at United", _
                                    "Amount due", "Due date", "Cancel date", "Policy status", _
                                    "Term effective", "Term expires", "When", "What happened")
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
    ' which report United happens to show first.
    ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
    ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"

    For i = 1 To policies.Count
        parts = policies(i)
        recordId = parts(0)
        siteNumber = parts(2)
        policyDigits = CrownDigits(siteNumber)

        If Len(policyDigits) = 0 Then GoTo NextPolicy

        dueAmount = ""
        dueDate = ""
        cancelDate = ""
        policyStatus = ""
        detailStatus = ""
        liveNumber = ""
        termEff = ""
        termExp = ""
        listStatus = ""
        note = ""

        On Error Resume Next

        LoopElementUntilFound drv, "tbxPolicyNo"
        drv.FindElementById("tbxPolicyNo").Clear
        drv.FindElementById("tbxPolicyNo").SendKeys policyDigits
        ClickElement drv, "btnSubmitPol1", "ID"
        drv.Wait 2000

        ' A renewed policy answers with its terms instead of going
        ' straight to one of them.
        '
        ' Asked in this order on purpose: only look for a list when we are
        ' not already on a policy. A policy page that happened to carry the
        ' words "Term Effective Date" anywhere on it would otherwise be
        ' read as a list, and clicked, for all hundred and thirty-nine of
        ' them rather than the four that need it.
        listText = ""

        If Not CrownUaigOnPolicy(drv) Then
            listText = drv.ExecuteScript(CrownUaigListScript())

            If Len(listText) > 0 Then
                pick = CrownUaigPickRow(listText, liveNumber, termEff, termExp, listStatus, note)

                If pick > 0 Then
                    drv.ExecuteScript CrownUaigClickScript(pick)
                    drv.Wait 2500
                End If
            End If
        End If

        ' Only read a policy if we are actually looking at one. Reading
        ' the list as though it were a policy is how empty answers get
        ' written over good ones.
        If CrownUaigOnPolicy(drv) Then
            If drv.IsElementPresent(By.ID("CURAMTDUE")) Then
                dueAmount = drv.FindElementById("CURAMTDUE").value
            End If

            detailStatus = drv.FindElementByXPath( _
                "//td[.//font[contains(.,'Policy Status')]]/following-sibling::td//strong").text

            If Trim$(detailStatus) = "" Then
                detailStatus = drv.FindElementByXPath( _
                    "//font[contains(text(),'Policy Status')]/ancestor::td/following-sibling::td//font").text
            End If

            detailStatus = Replace(detailStatus, vbLf, " ")
            detailStatus = Application.WorksheetFunction.Trim(detailStatus)

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
        ElseIf Len(listText) > 0 Then
            If Len(note) = 0 Then note = "United listed its terms but would not open one"
        End If

        On Error GoTo 0

        ' The list's own wording, unless the policy page gave a better one.
        policyStatus = detailStatus
        If Len(Trim$(policyStatus)) = 0 Then policyStatus = listStatus

        If Len(liveNumber) = 0 Then liveNumber = siteNumber
        liveDigits = CrownDigits(liveNumber)

        ws.Cells(row, 1).value = recordId
        ws.Cells(row, 2).value = siteNumber
        ws.Cells(row, 3).value = liveNumber
        ws.Cells(row, 4).value = dueAmount
        ws.Cells(row, 5).value = dueDate
        ws.Cells(row, 6).value = cancelDate
        ws.Cells(row, 7).value = policyStatus
        ws.Cells(row, 8).value = termEff
        ws.Cells(row, 9).value = termExp
        ws.Cells(row, 10).value = Now

        ' Nothing at all came back: the policy is not at this carrier, or
        ' the page did not load. Do not write emptiness over what the
        ' website already holds.
        If Len(Trim$(dueAmount)) = 0 And Len(Trim$(dueDate)) = 0 _
           And Len(Trim$(cancelDate)) = 0 And Len(Trim$(policyStatus)) = 0 Then
            ws.Cells(row, 11).value = Trim$("nothing found - left alone. " & note)
            blank = blank + 1
        Else
            answer = CrownUpdate(CROWN_FORM_POLICY, CLng(recordId), _
                Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                      "Status_", "Update_due_dates", "Last_Updated", "Current_policy_no"), _
                Array(dueAmount, CrownDate(dueDate), CrownDate(cancelDate), _
                      policyStatus, "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM"), liveNumber))

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 11).value = Trim$("written. " & note)
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 11).value = "website did not answer - run again for this one"
            Else
                ws.Cells(row, 11).value = "refused: " & answer
            End If
        End If

        ' United has moved this policy on to a number we have never seen.
        If Len(liveDigits) > 0 And liveDigits <> policyDigits Then
            If index.Exists(liveDigits) Then
                ws.Cells(row, 11).value = ws.Cells(row, 11).value & _
                    " The newer term is already on the site as record " & index(liveDigits) & "."
            Else
                renewals.Add Array(recordId, liveNumber, termEff, termExp, _
                                   dueAmount, dueDate, cancelDate, policyStatus, _
                                   names(policyDigits), siteNumber)
                index.Add liveDigits, recordId          ' so one run cannot offer it twice
                ws.Cells(row, 11).value = ws.Cells(row, 11).value & " RENEWED - new number."
            End If
        End If

        row = row + 1

NextPolicy:
        On Error Resume Next
        ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
        ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"
        drv.Wait 800
        On Error GoTo 0
    Next i

    madeCount = CrownMakeRenewals(renewals, ws)

    ws.Activate

    MsgBox policies.Count & " United Auto policies looked up." & vbCrLf & _
           done & " written back to the website." & vbCrLf & _
           blank & " had nothing to read and were left as they were." & vbCrLf & _
           renewals.Count & " had been renewed under a new number, " & madeCount & " added." & vbCrLf & vbCrLf & _
           "See the CrownPayments sheet for the detail. The Google Sheet " & _
           "picks the changes up within the hour.", vbInformation, "Crown Superior"
End Sub


' ---------------------------------------------------------------------
' 3. Renewals
' ---------------------------------------------------------------------
'
' Each one becomes a copy of the policy we already hold for that customer
' - same driver, same car, same cover - carrying the new number, the new
' term dates and "Renewal" as its description. The copy is made on the
' website, so nothing has to be typed back in.
'
' They are listed and agreed to first. A policy record is not something
' to create behind someone's back.
Private Function CrownMakeRenewals(ByVal renewals As Collection, ByVal ws As Object) As Long
    Dim renewal As Variant
    Dim listing As String, answer As String
    Dim i As Long, made As Long

    If renewals.Count = 0 Then Exit Function

    For i = 1 To renewals.Count
        renewal = renewals(i)
        listing = listing & renewal(8) & " - " & renewal(9) & " is now " & renewal(1)

        If Len(renewal(2)) > 0 Then listing = listing & ", effective " & renewal(2)

        listing = listing & vbCrLf
    Next i

    If MsgBox(renewals.Count & " of these policies have been renewed at United under a new " & _
              "number that is not on our website:" & vbCrLf & vbCrLf & listing & vbCrLf & _
              "Add each one as a new policy, copied from the one we hold, with the new " & _
              "number and dates and ""Renewal"" as the description?", _
              vbYesNo + vbQuestion, "Crown Superior") <> vbYes Then
        Exit Function
    End If

    For i = 1 To renewals.Count
        renewal = renewals(i)

        ' "March 23, 2023" is how that form writes a policy date. Sent in
        ' any other shape it goes in as text the calendar cannot read.
        answer = CrownClone(CROWN_FORM_POLICY, CLng(renewal(0)), _
            Array("Policy number", "Current_policy_no", "Description", _
                  "today_date_amin", "exp_date_amin", _
                  "Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                  "Status_", "Update_due_dates", "Last_Updated"), _
            Array(renewal(1), renewal(1), "Renewal", _
                  CrownLongDate(renewal(2)), CrownLongDate(renewal(3)), _
                  renewal(4), CrownDate(renewal(5)), CrownDate(renewal(6)), _
                  renewal(7), "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM")))

        If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
            made = made + 1
            CrownNoteRenewal ws, CStr(renewal(0)), "renewal added as record " & CrownIdFrom(answer)
        ElseIf Len(Trim$(answer)) = 0 Then
            CrownNoteRenewal ws, CStr(renewal(0)), "renewal NOT added - the website did not answer"
        Else
            CrownNoteRenewal ws, CStr(renewal(0)), "renewal NOT added: " & answer
        End If
    Next i

    CrownMakeRenewals = made
End Function

' The new record number out of the website's reply.
Private Function CrownIdFrom(ByVal answer As String) As String
    Dim at As Long, i As Long, ch As String

    at = InStr(1, answer, """id"":")
    If at = 0 Then Exit Function

    For i = at + 5 To Len(answer)
        ch = Mid$(answer, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
        CrownIdFrom = CrownIdFrom & ch
    Next i
End Function

' Put a word about the renewal on the line the policy is already on.
Private Sub CrownNoteRenewal(ByVal ws As Object, ByVal recordId As String, ByVal words As String)
    Dim r As Long

    For r = 2 To 5000
        If Len(Trim$(CStr(ws.Cells(r, 1).value))) = 0 Then Exit For

        If CStr(ws.Cells(r, 1).value) = recordId Then
            ws.Cells(r, 11).value = ws.Cells(r, 11).value & " " & words
            Exit Sub
        End If
    Next r
End Sub

' The policy form keeps its effective and expiration dates written out -
' "March 23, 2023". Anything that is not a date is left exactly as it is
' rather than being turned into today by accident.
Public Function CrownLongDate(ByVal text As String) As String
    Dim when As Double

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    when = CrownUsDate(text)

    If when = 0 Then
        CrownLongDate = text
    Else
        CrownLongDate = Format$(CDate(when), "mmmm d, yyyy")
    End If
End Function
