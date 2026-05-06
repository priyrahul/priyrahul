Attribute VB_Name = "SalaryExpenseAutomation"
Option Explicit

' ==========================
' Global constants
' ==========================
Public Const MASTER_SHEET As String = "MASTER"
Public Const SALARY_SHEET As String = "SALARY_INPUT"
Public Const EXPENSE_SHEET As String = "EXPENSE_INPUT"
Public Const DASHBOARD_SHEET As String = "DASHBOARD"
Public Const MASTER_PASSWORD As String = "ChangeMe#2026"

Public Const BANK_PYMT_PROD_TYPE As String = "PAB_VENDOR"
Public Const BANK_DEBIT_ACC As String = "431105000740"

' ==========================
' Workbook bootstrap
' ==========================
Public Sub SetupSystem()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    CreateOrResetMaster
    CreateOrResetSalaryInput
    CreateOrResetExpenseInput
    CreateOrResetDashboard
    AddMainButtons
    ProtectMasterSheet True

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Salary & Expense Payment Automation System initialized successfully.", vbInformation
End Sub

Public Sub OpenEmployeeForm()
    frmEmployeeMaster.Show
End Sub

Public Sub OpenSalaryForm()
    frmSalaryEntry.Show
End Sub

Public Sub OpenExpenseForm()
    frmExpenseEntry.Show
End Sub

Public Sub OpenSearchForm()
    frmEmployeeSearch.Show
End Sub

Public Sub GenerateBankFileMenu()
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Yes = Salary Payment file" & vbCrLf & "No = Expense Payment file", vbYesNoCancel + vbQuestion, "Generate Bank File")

    If choice = vbYes Then
        GenerateBankFile True
    ElseIf choice = vbNo Then
        GenerateBankFile False
    End If
End Sub

Public Sub RefreshDashboard()
    Dim wsD As Worksheet
    Set wsD = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    wsD.Range("B2").Value = TotalSalaryPaid()
    wsD.Range("B3").Value = TotalEmployees()
    BuildBranchSummary wsD
    BuildExpenseSummary wsD
    BuildMonthlySummary wsD

    wsD.Columns.AutoFit
End Sub

' ==========================
' Sheet creation / setup
' ==========================
Private Sub CreateOrResetMaster()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(MASTER_SHEET)

    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("Emp Code", "Name", "Account Number", "IFSC Code", "Branch")
    ws.Columns("A:E").EntireColumn.AutoFit
    ws.Columns("C").NumberFormat = "@"

    ws.Rows(1).Font.Bold = True
    ws.EnableSelection = xlUnlockedCells

    LockAllCells ws
    ws.Range("A2:E1048576").Locked = True
End Sub

Private Sub CreateOrResetSalaryInput()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SALARY_SHEET)

    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Date", "Emp Name", "Account Number", "IFSC", "Amount", "Branch", "Narration")
    ws.Columns("C").NumberFormat = "@"
    ws.Columns("E").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
End Sub

Private Sub CreateOrResetExpenseInput()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(EXPENSE_SHEET)

    ws.Cells.Clear
    ws.Range("A1:F1").Value = Array("Date", "Name", "Account Number", "IFSC", "Amount", "Narration")
    ws.Columns("C").NumberFormat = "@"
    ws.Columns("E").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"
    ws.Rows(1).Font.Bold = True
    ws.Columns.AutoFit
End Sub

Private Sub CreateOrResetDashboard()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(DASHBOARD_SHEET)

    ws.Cells.Clear
    ws.Range("A1").Value = "Dashboard"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "Total Salary Paid"
    ws.Range("A3").Value = "Total Employees"

    ws.Range("A5").Value = "Branch-wise Salary Summary"
    ws.Range("A10").Value = "Expense Summary"
    ws.Range("A15").Value = "Monthly Summary"

    ws.Columns("A:B").AutoFit
End Sub

Private Sub AddMainButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    On Error Resume Next
    ws.Shapes.SelectAll
    Selection.Delete
    On Error GoTo 0

    AddButton ws, "Add Employee", "OpenEmployeeForm", 350, 40
    AddButton ws, "Add Salary", "OpenSalaryForm", 350, 80
    AddButton ws, "Add Expense", "OpenExpenseForm", 350, 120
    AddButton ws, "Generate File", "GenerateBankFileMenu", 350, 160
    AddButton ws, "Search Employee", "OpenSearchForm", 350, 200
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 140, 28)
    shp.TextFrame.Characters.Text = caption
    shp.OnAction = macroName
End Sub

' ==========================
' Master insert/search
' ==========================
Public Function AddEmployeeRecord(ByVal empName As String, ByVal accNo As String, ByVal ifsc As String, ByVal branch As String) As Boolean
    On Error GoTo EH

    Dim ws As Worksheet, nextRow As Long, empCode As String
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)

    empName = Trim(empName)
    accNo = Trim(accNo)
    ifsc = UCase$(Trim(ifsc))
    branch = Trim(branch)

    ValidateEmployeeInput empName, accNo, ifsc, branch

    If EmployeeExists(empName) Then
        Err.Raise vbObjectError + 1000, , "Duplicate employee name is not allowed."
    End If

    ProtectMasterSheet False
    nextRow = LastRow(ws, "A") + 1
    empCode = GenerateEmpCode(nextRow - 1)

    ws.Cells(nextRow, 1).Value = empCode
    ws.Cells(nextRow, 2).Value = empName
    ws.Cells(nextRow, 3).NumberFormat = "@"
    ws.Cells(nextRow, 3).Value = "'" & accNo
    ws.Cells(nextRow, 4).Value = ifsc
    ws.Cells(nextRow, 5).Value = branch

    ProtectMasterSheet True
    AddEmployeeRecord = True
    Exit Function

EH:
    ProtectMasterSheet True
    MsgBox Err.Description, vbExclamation, "Employee Save Error"
    AddEmployeeRecord = False
End Function

Public Function FindEmployee(ByVal empName As String, ByRef empCode As String, ByRef accNo As String, ByRef ifsc As String, ByRef branch As String) As Boolean
    Dim ws As Worksheet, lr As Long, i As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)

    lr = LastRow(ws, "A")
    For i = 2 To lr
        If StrComp(Trim(ws.Cells(i, 2).Value), Trim(empName), vbTextCompare) = 0 Then
            empCode = ws.Cells(i, 1).Value
            accNo = CStr(ws.Cells(i, 3).Value)
            ifsc = CStr(ws.Cells(i, 4).Value)
            branch = CStr(ws.Cells(i, 5).Value)
            FindEmployee = True
            Exit Function
        End If
    Next i
End Function

' ==========================
' Salary / expense insert
' ==========================
Public Function AddSalaryEntry(ByVal dt As Date, ByVal empName As String, ByVal amount As Variant, ByVal baseNarr As String) As Boolean
    On Error GoTo EH

    Dim acc As String, ifsc As String, br As String, empCode As String
    Dim ws As Worksheet, nRow As Long

    If Not IsNumeric(amount) Then Err.Raise vbObjectError + 1001, , "Amount must be numeric."
    If Not FindEmployee(empName, empCode, acc, ifsc, br) Then Err.Raise vbObjectError + 1002, , "Employee not found in MASTER."

    Set ws = ThisWorkbook.Worksheets(SALARY_SHEET)
    nRow = LastRow(ws, "A") + 1

    ws.Cells(nRow, 1).Value = dt
    ws.Cells(nRow, 2).Value = empName
    ws.Cells(nRow, 3).NumberFormat = "@"
    ws.Cells(nRow, 3).Value = "'" & acc
    ws.Cells(nRow, 4).Value = ifsc
    ws.Cells(nRow, 5).Value = CDbl(amount)
    ws.Cells(nRow, 6).Value = br
    ws.Cells(nRow, 7).Value = BuildNarration(baseNarr, dt)

    AddSalaryEntry = True
    Exit Function
EH:
    MsgBox Err.Description, vbExclamation, "Salary Save Error"
End Function

Public Function AddExpenseEntry(ByVal dt As Date, ByVal nameVal As String, ByVal acc As String, ByVal ifsc As String, ByVal amount As Variant, ByVal baseNarr As String) As Boolean
    On Error GoTo EH

    If Trim(nameVal) = "" Then Err.Raise vbObjectError + 1003, , "Name is required."
    ValidateAccountNumber acc
    ValidateIFSC ifsc
    If Not IsNumeric(amount) Then Err.Raise vbObjectError + 1004, , "Amount must be numeric."

    Dim ws As Worksheet, nRow As Long
    Set ws = ThisWorkbook.Worksheets(EXPENSE_SHEET)
    nRow = LastRow(ws, "A") + 1

    ws.Cells(nRow, 1).Value = dt
    ws.Cells(nRow, 2).Value = nameVal
    ws.Cells(nRow, 3).NumberFormat = "@"
    ws.Cells(nRow, 3).Value = "'" & Trim(acc)
    ws.Cells(nRow, 4).Value = UCase$(Trim(ifsc))
    ws.Cells(nRow, 5).Value = CDbl(amount)
    ws.Cells(nRow, 6).Value = BuildNarration(baseNarr, dt)

    AddExpenseEntry = True
    Exit Function
EH:
    MsgBox Err.Description, vbExclamation, "Expense Save Error"
End Function

' ==========================
' Bank file export
' ==========================
Public Sub GenerateBankFile(ByVal isSalary As Boolean)
    On Error GoTo EH

    Dim src As Worksheet, wbOut As Workbook, wsOut As Worksheet
    Dim lr As Long, i As Long, oRow As Long
    Dim branchFilter As String, branchVal As String

    If isSalary Then
        Set src = ThisWorkbook.Worksheets(SALARY_SHEET)
        branchFilter = InputBox("Enter Branch name for branch-wise export (leave blank for all):", "Branch Filter")
    Else
        Set src = ThisWorkbook.Worksheets(EXPENSE_SHEET)
        branchFilter = ""
    End If

    lr = LastRow(src, "A")
    If lr < 2 Then
        MsgBox "No entries found to export.", vbInformation
        Exit Sub
    End If

    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)

    wsOut.Range("A1:N1").Value = Array("PYMT_PROD_TYPE_CODE", "PYMT_MODE", "DEBIT_ACC_NO", "BNF_NAME", "BENE_ACC_NO", "BENE_IFSC", "AMOUNT", "DEBIT_NARR", "CREDIT_NARR", "MOBILE_NUM", "EMAIL_ID", "REMARK", "PYMT_DATE", "REF_NO")

    oRow = 2
    For i = 2 To lr
        If isSalary Then
            branchVal = CStr(src.Cells(i, 6).Value)
            If Len(Trim(branchFilter)) > 0 Then
                If StrComp(Trim(branchVal), Trim(branchFilter), vbTextCompare) <> 0 Then GoTo ContinueLoop
            End If

            WriteOutputRow wsOut, oRow, src.Cells(i, 2).Value, src.Cells(i, 3).Value, src.Cells(i, 4).Value, src.Cells(i, 5).Value, src.Cells(i, 7).Value
        Else
            WriteOutputRow wsOut, oRow, src.Cells(i, 2).Value, src.Cells(i, 3).Value, src.Cells(i, 4).Value, src.Cells(i, 5).Value, src.Cells(i, 6).Value
        End If
        oRow = oRow + 1
ContinueLoop:
    Next i

    wsOut.Columns("E").NumberFormat = "@"
    wsOut.Columns.AutoFit

    Dim fname As String, usedBranch As String
    usedBranch = IIf(Len(Trim(branchFilter)) = 0, "ALL_BRANCHES", branchFilter)
    fname = "SALARY PAYMENT AS ON " & Format(Date, "dd-mm-yyyy") & " - " & usedBranch & ".xlsx"

    wbOut.SaveAs ThisWorkbook.Path & Application.PathSeparator & fname, xlOpenXMLWorkbook
    MsgBox "Bank file generated: " & fname, vbInformation
    Exit Sub
EH:
    MsgBox "File generation error: " & Err.Description, vbExclamation
End Sub

Private Sub WriteOutputRow(ByVal ws As Worksheet, ByVal r As Long, ByVal bnfName As String, ByVal beneAcc As String, ByVal beneIfsc As String, ByVal amt As Variant, ByVal narr As String)
    ws.Cells(r, 1).Value = BANK_PYMT_PROD_TYPE
    ws.Cells(r, 2).Value = PaymentMode(CStr(beneIfsc))
    ws.Cells(r, 3).Value = BANK_DEBIT_ACC
    ws.Cells(r, 4).Value = bnfName
    ws.Cells(r, 5).NumberFormat = "@"
    ws.Cells(r, 5).Value = "'" & CStr(beneAcc)
    ws.Cells(r, 6).Value = UCase$(CStr(beneIfsc))
    ws.Cells(r, 7).Value = CDbl(amt)
    ws.Cells(r, 8).Value = narr
    ws.Cells(r, 9).Value = narr
    ws.Cells(r, 10).Value = ""
    ws.Cells(r, 11).Value = ""
    ws.Cells(r, 12).Value = ""
    ws.Cells(r, 13).Value = Date
    ws.Cells(r, 14).Value = ""
End Sub

' ==========================
' Validation/helpers
' ==========================
Public Function PaymentMode(ByVal ifsc As String) As String
    If UCase$(Left$(Trim(ifsc), 4)) = "ICIC" Then
        PaymentMode = "FT"
    Else
        PaymentMode = "NEFT"
    End If
End Function

Public Function BuildNarration(ByVal baseNarr As String, ByVal txnDate As Date) As String
    BuildNarration = Trim(baseNarr) & " " & Format(txnDate, "mmm-yyyy")
End Function

Private Sub ValidateEmployeeInput(ByVal empName As String, ByVal accNo As String, ByVal ifsc As String, ByVal branch As String)
    If empName = "" Then Err.Raise vbObjectError + 2000, , "Employee name is required."
    If branch = "" Then Err.Raise vbObjectError + 2001, , "Branch is required."
    ValidateAccountNumber accNo
    ValidateIFSC ifsc
End Sub

Public Sub ValidateAccountNumber(ByVal accNo As String)
    Dim t As String, i As Long, ch As String
    t = Trim(accNo)
    If t = "" Then Err.Raise vbObjectError + 2002, , "Account number must not be blank."

    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch < "0" Or ch > "9" Then
            Err.Raise vbObjectError + 2003, , "Account number must be numeric characters only."
        End If
    Next i
End Sub

Public Sub ValidateIFSC(ByVal ifsc As String)
    Dim t As String
    t = UCase$(Trim(ifsc))
    If Len(t) <> 11 Then Err.Raise vbObjectError + 2004, , "IFSC must be exactly 11 characters."
    If Not t Like "[A-Z][A-Z][A-Z][A-Z]0[A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9][A-Z0-9]" Then
        Err.Raise vbObjectError + 2005, , "Invalid IFSC format."
    End If
End Sub

Public Function EmployeeExists(ByVal empName As String) As Boolean
    Dim ws As Worksheet, lr As Long, i As Long
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    lr = LastRow(ws, "A")

    For i = 2 To lr
        If StrComp(Trim(ws.Cells(i, 2).Value), Trim(empName), vbTextCompare) = 0 Then
            EmployeeExists = True
            Exit Function
        End If
    Next i
End Function

Public Function GenerateEmpCode(ByVal serialNo As Long) As String
    GenerateEmpCode = "EMP" & Format(serialNo, "0000")
End Function

Public Function LastRow(ByVal ws As Worksheet, ByVal col As String) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If r < 1 Then r = 1
    LastRow = r
End Function

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Public Sub LockAllCells(ByVal ws As Worksheet)
    ws.Cells.Locked = True
End Sub

Public Sub ProtectMasterSheet(ByVal protectIt As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    If protectIt Then
        ws.Protect Password:=MASTER_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    Else
        ws.Unprotect MASTER_PASSWORD
    End If
End Sub

' ==========================
' Dashboard helpers
' ==========================
Private Function TotalSalaryPaid() As Double
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets(SALARY_SHEET)
    lr = LastRow(ws, "A")
    If lr < 2 Then Exit Function
    TotalSalaryPaid = Application.WorksheetFunction.Sum(ws.Range("E2:E" & lr))
End Function

Private Function TotalEmployees() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MASTER_SHEET)
    TotalEmployees = Application.Max(0, LastRow(ws, "A") - 1)
End Function

Private Sub BuildBranchSummary(ByVal wsD As Worksheet)
    wsD.Range("A6:C999").ClearContents
    wsD.Range("A6:C6").Value = Array("Branch", "Total Amount", "Count")
    SummarizeByColumn ThisWorkbook.Worksheets(SALARY_SHEET), 6, 5, wsD, 7
End Sub

Private Sub BuildExpenseSummary(ByVal wsD As Worksheet)
    Dim wsE As Worksheet, lr As Long
    Set wsE = ThisWorkbook.Worksheets(EXPENSE_SHEET)
    lr = LastRow(wsE, "A")

    wsD.Range("A11:B12").ClearContents
    wsD.Range("A11").Value = "Total Expense Paid"
    wsD.Range("B11").Value = IIf(lr < 2, 0, Application.WorksheetFunction.Sum(wsE.Range("E2:E" & lr)))
End Sub

Private Sub BuildMonthlySummary(ByVal wsD As Worksheet)
    wsD.Range("A16:C999").ClearContents
    wsD.Range("A16:C16").Value = Array("Month", "Salary", "Expense")
    wsD.Range("A17").Value = "Use Pivot Table or extend macro for custom month-wise analysis"
End Sub

Private Sub SummarizeByColumn(ByVal wsSrc As Worksheet, ByVal keyCol As Long, ByVal amtCol As Long, ByVal wsOut As Worksheet, ByVal startRow As Long)
    Dim dict As Object, dictCnt As Object
    Dim lr As Long, i As Long, k As Variant, r As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set dictCnt = CreateObject("Scripting.Dictionary")

    lr = LastRow(wsSrc, "A")
    For i = 2 To lr
        Dim key As String
        key = Trim(CStr(wsSrc.Cells(i, keyCol).Value))
        If key <> "" Then
            If Not dict.Exists(key) Then
                dict.Add key, 0#
                dictCnt.Add key, 0&
            End If
            dict(key) = dict(key) + CDbl(Val(wsSrc.Cells(i, amtCol).Value))
            dictCnt(key) = dictCnt(key) + 1
        End If
    Next i

    r = startRow
    For Each k In dict.Keys
        wsOut.Cells(r, 1).Value = k
        wsOut.Cells(r, 2).Value = dict(k)
        wsOut.Cells(r, 3).Value = dictCnt(k)
        r = r + 1
    Next k
End Sub
