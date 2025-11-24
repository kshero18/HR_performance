Attribute VB_Name = "Module1"
Option Explicit

' ======== CONFIG READERS ========
Private Function GetCfg(ByVal name As String) As String
    On Error Resume Next
    GetCfg = CStr(Evaluate(name))
    On Error GoTo 0
End Function

Private Function BuildConnStr() As String
    ' Reads ServerName and DatabaseName from Config sheet
    Dim srv As String, db As String
    srv = GetCfg("ServerName")     ' <— Named cell "ServerName"
    db = GetCfg("DatabaseName")    ' <— Named cell "DatabaseName"

    ' Windows Authentication (Trusted Connection)
    ' Preferred modern driver:
    BuildConnStr = "Provider=MSOLEDBSQL;Data Source=" & srv & _
                   ";Initial Catalog=" & db & ";Integrated Security=SSPI;"
End Function

' ======== GENERIC QUERY RUNNER ========
Private Sub QueryToListObject(ByVal sql As String, ByVal targetSheet As String, ByVal tableName As String)
    Dim cn As Object, rs As Object
    Dim ws As Worksheet, lo As ListObject, rng As Range
    Dim f As Integer, lastRow As Long, lastCol As Long

    Set ws = ThisWorkbook.Worksheets(targetSheet)
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    ' ?? Show connection string for debugging
    Debug.Print "Connecting: " & BuildConnStr

    cn.Open BuildConnStr
    rs.Open sql, cn, 0, 1, 1  ' 0=adOpenForwardOnly, 1=adLockReadOnly, 1=adCmdText

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    For Each lo In ws.ListObjects
        If lo.name = tableName Then lo.Unlist: Exit For
    Next lo
    On Error GoTo 0

    ws.Cells.ClearContents

    If rs.EOF And rs.BOF Then
        ws.Range("A1").Value = "No rows returned."
    Else
        ' Write headers
        For f = 0 To rs.Fields.Count - 1
            ws.Cells(1, f + 1).Value = rs.Fields(f).name
        Next f

        ' Write data
        Set rng = ws.Range("A2")
        rng.CopyFromRecordset rs

        ' Format as table
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
        lo.name = tableName
        lo.TableStyle = "TableStyleMedium2"
    End If

    rs.Close: cn.Close
    Set rs = Nothing: Set cn = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ======== KPI QUERIES ========
Private Function Sql_KpiOverview() As String
    Sql_KpiOverview = "SELECT * FROM dbo.vw_kpi_overview;"
End Function

Private Function Sql_DeptHub(Optional ByVal dept As String = "") As String
    Dim s As String
    s = "SELECT * FROM dbo.vw_kpi_department_hub"
    If Len(dept) > 0 Then
        s = "SELECT * FROM dbo.vw_kpi_department_hub WHERE Department = '" & Replace(dept, "'", "''") & "'"
    End If
    Sql_DeptHub = s & ";"
End Function

Private Function Sql_AttritionTenure() As String
    Sql_AttritionTenure = "SELECT TenureGroup, LeftCount, TotalCount, AttritionRatePct, TenureSort FROM dbo.vw_attrition_tenure_group;"
End Function

Private Function Sql_OvertimeAttrition() As String
    Sql_OvertimeAttrition = "SELECT * FROM dbo.vw_overtime_attrition;"
End Function

' ======== PUBLIC ENTRYPOINTS ========
Public Sub RefreshData()
    Dim dept As String
    dept = GetCfg("DepartmentFilter")

    QueryToListObject Sql_KpiOverview(), "Data", "tblKpiOverview"
    EnsureSheet "Data_KpiDept"
    QueryToListObject Sql_DeptHub(), "Data_KpiDept", "tblKpiDepartment"
    EnsureSheet "Data_Tenure"
    QueryToListObject Sql_AttritionTenure(), "Data_Tenure", "tblAttritionTenure"
    EnsureSheet "Data_Overtime"
    QueryToListObject Sql_OvertimeAttrition(), "Data_Overtime", "tblOvertimeAttrition"

    RefreshAllPivots
    MsgBox "Data refresh complete.", vbInformation
End Sub

Private Sub EnsureSheet(ByVal name As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).name = name
    Else
        ws.Cells.ClearContents
    End If
End Sub

Private Sub RefreshAllPivots()
    Dim ws As Worksheet, pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        On Error Resume Next
        pc.Refresh
        On Error GoTo 0
    Next pc
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Calculate
        On Error GoTo 0
    Next ws
End Sub

Public Sub ExportDashboardToPDF()
    Dim outDir As String, prefix As String, outPath As String
    Dim shNames As Variant

    outDir = GetCfg("ExportFolder")
    prefix = GetCfg("PDFFilePrefix")
    If Len(outDir) = 0 Then outDir = ThisWorkbook.Path

    ' Ensure output folder exists
    If Dir(outDir, vbDirectory) = vbNullString Then
        MkDir outDir
    End If

    outPath = outDir & "\" & prefix & "_" & Format(Date, "yyyy_mm_dd") & ".pdf"

    ' List of sheets to include in the PDF, in order
    shNames = Array("Dashboard", "Data_Issues_Log", "Data_Dictionary", "KPI_Definitions")

    ' Select all sheets to be exported as a single multi-page PDF
    ThisWorkbook.Sheets(shNames).Select

    ' Export all selected sheets
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' Return selection back to a single sheet
    ThisWorkbook.Worksheets("Dashboard").Select

    MsgBox "Exported PDF to:" & vbCrLf & outPath, vbInformation
End Sub

Private Sub FormatDocSheetsAsTables()
    Dim sheetNames As Variant
    Dim sh As Variant
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim lo As ListObject
    Dim rng As Range

    sheetNames = Array("Data_Issues_Log", "Data_Dictionary", "KPI_Definitions")

    For Each sh In sheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(sh))
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' Clear any existing ListObjects so we don't stack them
            On Error Resume Next
            For Each lo In ws.ListObjects
                lo.Unlist
            Next lo
            On Error GoTo 0

            ' Work out the used range
            With ws
                lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
                lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column

                If lastRow >= 1 And lastCol >= 1 Then
                    Set rng = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))

                    ' Create a table from the range
                    Set lo = .ListObjects.Add(SourceType:=xlSrcRange, _
                                              Source:=rng, _
                                              XlListObjectHasHeaders:=xlYes)
                    lo.name = "tbl_" & CStr(sh)
                    lo.TableStyle = "TableStyleMedium2"
                End If
            End With
        End If

        Set ws = Nothing
        Set lo = Nothing
        Set rng = Nothing
    Next sh
End Sub

Public Sub RefreshAndExport()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Call RefreshData
    Call FormatDocSheetsAsTables   ' <— add this line
    Call ExportDashboardToPDF

Done:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
EH:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Done
End Sub

