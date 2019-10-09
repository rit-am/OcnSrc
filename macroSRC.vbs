Option Explicit
Private Sub Workbook_Open()
    MsgBox "SRC Validator"
End Sub
Sub CreateSheet(VarSName As String)
    Dim SheetExists As Boolean
    Dim Sheet As Worksheet
    SheetExists = False
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name = VarSName Then
            Sheet.Delete
            SheetExists = False
        End If
    Next Sheet
    If SheetExists = False Then
        With ThisWorkbook
            .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = VarSName
        End With
    End If
End Sub
Sub HelloWorld()
'--------------------------------------------------------------------------------------------------------
    'REM - Declare Variables Here
    Dim myFile As String, arr() As String, Txt As String
    Dim i As Integer, j As Integer, x As Integer
    Dim RecTypeTable() As Integer
    Dim Flag_HD As Boolean, Flag_TR As Boolean, Flag_HDTR As Boolean
    Dim Count_HD As Integer, Count_TR As Integer, POS_HD As Integer, POS_TR As Integer
    Dim IntRow As Integer, IntCol As Integer
'--------------------------------------------------------------------------------------------------------
    'REM - Define Variables Here
    myFile = "C:\Users\ritam\Downloads\APAC\OPSRCF\Omnipay src\Omnipay src\tmp.src"
    x = 1
    Flag_HD = False
    Flag_TR = False
    Flag_HDTR = False
    ReDim RecTypeTable(16)
'--------------------------------------------------------------------------------------------------------
    'REM - Open And Read and File
    Open myFile For Input As #1
    Do Until EOF(1)
        i = i + 1
        ReDim Preserve arr(i)
        Line Input #1, Txt
    Loop
    Close #1
'--------------------------------------------------------------------------------------------------------
    'REM - Move File Data to Array
    arr() = Split(Txt, vbLf)
'--------------------------------------------------------------------------------------------------------
    'REM - Analyze File Data
    'REM - Analyse Headers
    For i = LBound(arr) To UBound(arr)
        Txt = Mid(arr(i), 1, 3)
        If Trim(Txt) = "HD" Then
            Flag_HD = True
            Count_HD = Count_HD + 1
            POS_HD = i
        ElseIf Trim(Txt) = "TR" Then
            Flag_TR = True
            Count_TR = Count_TR + 1
            POS_TR = i
        Else
            'MsgBox (Int(Txt))
            RecTypeTable(Int(Txt)) = RecTypeTable(Int(Txt)) + 1
        End If
    Next i
'--------------------------------------------------------------------------------------------------------
    'REM - Analyze File Data
    'REM - Report File Header.Trailer
    Txt = ""
    If Flag_TR = True And Flag_HD = True _
        And Count_HD = 1 And Count_TR = 1 _
        And POS_HD = LBound(arr) And POS_TR = UBound(arr) Then
            MsgBox ("Header & Trailer record found")
            Flag_HDTR = True
        Else
            MsgBox ("Header & Trailer record sanity falied")
            Txt = Txt & "Header Found Flag : " & Flag_HD & vbCrLf
            Txt = Txt & "Header Found Count : " & Count_HD & vbCrLf
            Txt = Txt & "Header Found POS : " & POS_HD & vbCrLf
            Txt = Txt & "Trailer Found Flag : " & Flag_TR & vbCrLf
            Txt = Txt & "Trailer Found Count : " & Count_TR & vbCrLf
            Txt = Txt & "Header Found POS : " & POS_TR & vbCrLf
            MsgBox (Txt)
            Debug.Print (Txt)
    End If
'--------------------------------------------------------------------------------------------------------
    'REM - Analyze File Data
    'REM - Report File Record Distribution
    Dim TotalRecords As Integer
    Txt = ""
    TotalRecords = 0
    For i = 0 To 16
        Txt = Txt & "R-" & i & " : " & RecTypeTable(i) & vbCrLf
        TotalRecords = TotalRecords + RecTypeTable(i)
    Next i
    MsgBox (Txt & vbCrLf & "Total Records :" & TotalRecords)
    
 '--------------------------------------------------------------------------------------------------------
    'REM - Optimize File Data
    'REM - Sort Records
    Txt = ""
    For i = LBound(arr) To UBound(arr)
        For j = (LBound(arr) + 1) To UBound(arr)
            If (Trim(Mid(arr(i), 1, 3)) <> "HD") And (Trim(Mid(arr(i), 1, 3)) <> "TR") _
                And (Trim(Mid(arr(j), 1, 3)) <> "HD") And (Trim(Mid(arr(j), 1, 3)) <> "TR") Then
                If Int(Mid(arr(i), 1, 3)) < Int(Mid(arr(j), 1, 3)) Then
                    Txt = arr(i)
                    arr(i) = arr(j)
                    arr(j) = Txt
                End If
            End If
        Next j
    Next i
'--------------------------------------------------------------------------------------------------------
    'REM - Insert Array to XLS
    CreateSheet "AscArray"
    For i = LBound(arr) To UBound(arr)
        'MsgBox (arr(i))
        Worksheets("AscArray").Cells(i + 1, 1) = "'" & "D"
        Worksheets("AscArray").Cells(i + 1, 2) = "'" & arr(i)
    Next i
'--------------------------------------------------------------------------------------------------------
    'REM - Put Record Definations in xls
    CreateSheet "RecordDefinations"
    Dim SWorkSheetdefOP As Worksheet
    Dim SRecType As String, SFieldName As String
    'SWorkSheetdefOP = ThisWorkbook.Worksheets("def_OP")
    Dim SRecordDefn(17, 300)
    Dim Counter As Integer
    Counter = 0
    
    SRecordDefn(1, 1) = "SRecType"
    SRecordDefn(2, 1) = "SFieldName"
    SRecordDefn(3, 1) = "SMOC"
    SRecordDefn(4, 1) = "SPosition"
    SRecordDefn(5, 1) = "SFieldName"
    SRecordDefn(6, 1) = "SLength"
    SRecordDefn(7, 1) = "SDataType"
    SRecordDefn(8, 1) = "STabFieldMapping"
    SRecordDefn(9, 1) = "SValueLogic"
    SRecordDefn(10, 1) = "SDescription"
    SRecordDefn(11, 1) = "SHardDefinedValue"
    SRecordDefn(12, 1) = "SUserDefinedValue"
    
    For IntRow = 2 To 220
        SRecordDefn(1, IntRow) = Worksheets("def_OP").Cells(IntRow, 1)
        SRecordDefn(2, IntRow) = Worksheets("def_OP").Cells(IntRow, 2)
        SRecordDefn(3, IntRow) = Worksheets("def_OP").Cells(IntRow, 3)
        SRecordDefn(4, IntRow) = Worksheets("def_OP").Cells(IntRow, 4)
        SRecordDefn(5, IntRow) = Worksheets("def_OP").Cells(IntRow, 5)
        SRecordDefn(6, IntRow) = Worksheets("def_OP").Cells(IntRow, 6)
        SRecordDefn(7, IntRow) = Worksheets("def_OP").Cells(IntRow, 7)
        SRecordDefn(8, IntRow) = Worksheets("def_OP").Cells(IntRow, 8)
        SRecordDefn(9, IntRow) = Worksheets("def_OP").Cells(IntRow, 9)
        SRecordDefn(10, IntRow) = Worksheets("def_OP").Cells(IntRow, 10)
        SRecordDefn(11, IntRow) = 0
        SRecordDefn(12, IntRow) = 0
    Next IntRow
    
    
    
    For i = 1 To 3
        Txt = ""
        For j = 1 To 12
            Txt = Txt & SRecordDefn(j, 1) & " : " & SRecordDefn(j, i + 1) & vbCrLf
            Worksheets("RecordDefinations").Cells(i, j) = SRecordDefn(j, i + 1)
        Next j
        MsgBox (Txt)
    Next i
    
    
    
    
'--------------------------------------------------------------------------------------------------------
    'REM - Save and Close Warning

'--------------------------------------------------------------------------------------------------------
    'REM - Save and Close
    ThisWorkbook.Save
    ThisWorkbook.Close
'--------------------------------------------------------------------------------------------------------
End Sub
