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
Public Function CrownPolicyIndex(Optional ByVal carrier As String = "") As Object
    Dim csv As String, lines() As String, parts() As String
    Dim index As Object
    Dim i As Long, digits As String

    Set index = CreateObject("Scripting.Dictionary")

    csv = CrownPost(CrownBody("policylist") & IIf(Len(carrier) > 0, "&carrier=" & CrownEnc(carrier), ""))

    If Len(csv) = 0 Then
        MsgBox "The website did not answer. Check your internet, then run CrownTestConnection.", _
               vbExclamation, "Crown Superior"
        Set CrownPolicyIndex = index
        Exit Function
    End If

    If Left$(csv, 5) = "error" Then
        MsgBox "The website answered: " & csv, vbExclamation, "Crown Superior"
        Set CrownPolicyIndex = index
        Exit Function
    End If

    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)

    For i = 1 To UBound(lines)                    ' 0 is the header
        If Len(Trim$(lines(i))) > 0 Then
            parts = CrownSplitCsvLine(lines(i))

            If UBound(parts) >= 2 Then
                digits = CrownDigits(parts(2))    ' the policy number

                ' The newest record wins: the list comes back newest first,
                ' so the first one seen for a number is the one to write to.
                If Len(digits) > 0 And Not index.Exists(digits) Then
                    index.Add digits, parts(0)    ' -> record number
                End If
            End If
        End If
    Next i

    Set CrownPolicyIndex = index
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
    Dim index As Object
    Dim lastRow As Long, i As Long, logRow As Long
    Dim digits As String, recordId As String, answer As String
    Dim done As Long, missing As Long, refused As Long
    Dim names As Variant, values As Variant

    On Error Resume Next
    Set wsUpload = ThisWorkbook.Worksheets("UploadPaymentdue")
    On Error GoTo 0

    If wsUpload Is Nothing Then
        MsgBox "There is no UploadPaymentdue sheet in this workbook.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set index = CrownPolicyIndex()
    If index.Count = 0 Then Exit Sub

    Set wsLog = CrownSheet("CrownUploadLog")
    wsLog.Cells.ClearContents
    wsLog.Range("A1:J1").value = Array("Policy Number", "Website", "First Name", "Last Name", _
                                       "Cancel Date", "Amount Due", "Due Date", "Status", "When", "What happened")
    logRow = 2

    lastRow = wsUpload.Cells(wsUpload.rows.Count, COL_POLICY).End(xlUp).row

    For i = 2 To lastRow
        If Trim$(CStr(wsUpload.Cells(i, COL_POLICY).value)) <> "" Then
            digits = CrownDigits(CStr(wsUpload.Cells(i, COL_POLICY).value))

            wsLog.Cells(logRow, 1).value = wsUpload.Cells(i, COL_POLICY).value
            wsLog.Cells(logRow, 2).value = wsUpload.Cells(i, COL_SITE).value
            wsLog.Cells(logRow, 3).value = wsUpload.Cells(i, COL_FIRST).value
            wsLog.Cells(logRow, 4).value = wsUpload.Cells(i, COL_LAST).value
            wsLog.Cells(logRow, 5).value = wsUpload.Cells(i, COL_CANCEL).value
            wsLog.Cells(logRow, 6).value = wsUpload.Cells(i, COL_AMOUNT).value
            wsLog.Cells(logRow, 7).value = wsUpload.Cells(i, COL_DUE).value
            wsLog.Cells(logRow, 8).value = wsUpload.Cells(i, COL_STATUS).value
            wsLog.Cells(logRow, 9).value = Now

            If Not index.Exists(digits) Then
                wsLog.Cells(logRow, 10).value = "No policy with that number on the website"
                missing = missing + 1
            Else
                recordId = index(digits)

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
                    wsLog.Cells(logRow, 10).value = "Written to record " & recordId
                    done = done + 1
                Else
                    wsLog.Cells(logRow, 10).value = "Refused: " & answer
                    refused = refused + 1
                End If
            End If

            logRow = logRow + 1
        End If
    Next i

    wsLog.Activate

    MsgBox done & " policy record(s) updated on the website." & vbCrLf & _
           missing & " had no matching policy number there." & vbCrLf & _
           refused & " were refused - see the CrownUploadLog sheet." & vbCrLf & vbCrLf & _
           "The Google Sheet picks these up within the hour on its own.", _
           vbInformation, "Crown Superior"
End Sub


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
