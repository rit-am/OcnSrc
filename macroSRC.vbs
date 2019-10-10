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
    Dim myFile As String, arr() As String, Txt As String, tmpTxt As String
    Dim myPath As String
    Dim i As Integer, j As Integer, x As Integer
    Dim RecTypeTable() As Integer, IntValCountOfRecordtypes As Integer
    Dim Flag_HD As Boolean, Flag_TR As Boolean, Flag_HDTR As Boolean
    Dim Count_HD As Integer, Count_TR As Integer, POS_HD As Integer, POS_TR As Integer
    Dim IntRow As Integer, IntCol As Integer
    Dim SWorkSheetdefOP As Worksheet
    Dim SRecType As String, SFieldName As String
    Dim SRecordDefn(17, 300) As String, HeaderDefn(300, 17) As String
    Dim Counter As Integer
    Dim TotalRecords As Integer
    Dim valIntOne As Integer
    Dim vRecType As String, vFieldName As String, _
        vMOC As String, vPosition As String, _
        vLength As String, vDataType As String, _
        vTabFieldMapping As String, vValueLogic As String, _
        vDescription As String, vHardDefinedValue As String, _
        vUserDefinedValue As String
'--------------------------------------------------------------------------------------------------------
    'REM - Define Variables Here
    myPath = ThisWorkbook.Path & "\"
    myFile = myPath & "tmp.src"
    x = 1
    Flag_HD = False
    Flag_TR = False
    Flag_HDTR = False
    valIntOne = 1
    IntValCountOfRecordtypes = 16 'Types of record in SRC file other than HD & TR
    ReDim RecTypeTable(16)
    'SWorkSheetdefOP = ThisWorkbook.Worksheets("def_OP")
'--------------------------------------------------------------------------------------------------------
    'REM - Create Required Sheets
    CreateSheet "RecordDefinations"
    CreateSheet "DataSet"
    CreateSheet "AscArray"
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
        Txt = Mid(arr(i), valIntOne, 3)
        If Len(Txt) >= 3 Then
            If Trim(Txt) = "HD" Then
                Flag_HD = True
                Count_HD = Count_HD + valIntOne
                POS_HD = i
            ElseIf Trim(Txt) = "TR" Then
                Flag_TR = True
                Count_TR = Count_TR + valIntOne
                POS_TR = i
            Else
                RecTypeTable(Int(Txt)) = RecTypeTable(Int(Txt)) + valIntOne
            End If
        End If
    Next i
'--------------------------------------------------------------------------------------------------------
    'REM - Analyze File Data
    'REM - Report File Header.Trailer
    Txt = ""
    If Flag_TR = True And Flag_HD = True _
        And Count_HD = valIntOne And Count_TR = valIntOne _
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
    End If
'--------------------------------------------------------------------------------------------------------
    'REM - Analyze File Data
    'REM - Report File Record Distribution
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
        For j = (LBound(arr) + valIntOne) To UBound(arr)
            If (Len(arr(i)) >= 3) And (Len(arr(j)) >= 3) Then
                If (Trim(Mid(arr(i), valIntOne, 3)) <> "HD") And (Trim(Mid(arr(i), valIntOne, 3)) <> "TR") _
                    And (Trim(Mid(arr(j), valIntOne, 3)) <> "HD") And (Trim(Mid(arr(j), valIntOne, 3)) <> "TR") Then
                    If Int(Mid(arr(i), valIntOne, 3)) < Int(Mid(arr(j), valIntOne, 3)) Then
                        Txt = arr(i)
                        arr(i) = arr(j)
                        arr(j) = Txt
                    End If
                End If
            End If
        Next j
    Next i
'--------------------------------------------------------------------------------------------------------
    'REM - Insert Array to XLS
    For i = LBound(arr) To UBound(arr)
        'MsgBox (arr(i))
        Worksheets("AscArray").Cells(i + valIntOne, 1) = "'" & "D"
        Worksheets("AscArray").Cells(i + valIntOne, 2) = "'" & arr(i)
    Next i
'--------------------------------------------------------------------------------------------------------
    'REM - Put Record Definations in xls
  
    SRecordDefn(1, 1) = "SRecType"
    SRecordDefn(2, 1) = "SFieldName"
    SRecordDefn(3, 1) = "SFieldDisplayName"
    SRecordDefn(4, 1) = "SMOC"
    SRecordDefn(5, 1) = "SPosition"
    SRecordDefn(6, 1) = "SLength"
    SRecordDefn(7, 1) = "SDataType"
    SRecordDefn(8, 1) = "STabFieldMapping"
    SRecordDefn(9, 1) = "SValueLogic"
    SRecordDefn(10, 1) = "SDescription"
    SRecordDefn(11, 1) = "SHardDefinedValue"
    SRecordDefn(12, 1) = "SUserDefinedValue"
    
    HeaderDefn(0, 1) = "Record Type"
    HeaderDefn(0, 2) = "Field Name "
    HeaderDefn(0, 3) = "Field Display Name"
    HeaderDefn(0, 4) = "M / O / C"
    HeaderDefn(0, 5) = "Position"
    HeaderDefn(0, 6) = "Length"
    HeaderDefn(0, 7) = "Data Type"
    HeaderDefn(0, 8) = "Tab Field Mapping"
    HeaderDefn(0, 9) = "Value / Logic"
    HeaderDefn(0, 10) = "Description"
    HeaderDefn(0, 11) = "Hard Coded Value"
    HeaderDefn(0, 12) = "User Defined Value"
    
    For IntRow = 2 To Worksheets("def_OP").Cells(Rows.Count, valIntOne).End(xlUp).Row
        For j = valIntOne To UBound(SRecordDefn)
            SRecordDefn(j, IntRow) = Worksheets("def_OP").Cells(IntRow, j)
        Next j
    Next IntRow
    
    For i = valIntOne To Worksheets("def_OP").Cells(Rows.Count, valIntOne).End(xlUp).Row
        Txt = ""
        For j = valIntOne To UBound(SRecordDefn)
            Txt = Txt & SRecordDefn(j, valIntOne) & " : " & SRecordDefn(j, i + valIntOne) & vbCrLf
            Worksheets("RecordDefinations").Cells(i, j) = SRecordDefn(j, i + valIntOne)
        Next j
    Next i
    
    Counter = 0
    IntCol = 0
    IntRow = 0
    tmpTxt = ""
    For Counter = valIntOne To Worksheets("RecordDefinations").Cells(Rows.Count, valIntOne).End(xlUp).Row
        IntCol = IntCol + valIntOne
        Txt = Worksheets("RecordDefinations").Cells(Counter, valIntOne)
        If (tmpTxt <> Txt) Then
            IntRow = IntRow + valIntOne
            IntCol = valIntOne
        End If
        tmpTxt = Txt
        For j = valIntOne To UBound(SRecordDefn)
            HeaderDefn(IntRow, j) = Worksheets("RecordDefinations").Cells(Counter, j)
        Next j
        Txt = ""
        For i = 3 To 3
            'Txt = Txt & HeaderDefn(0, i) & " :: " & HeaderDefn(IntRow, i) & vbCrLf
            Txt = Txt & HeaderDefn(IntRow, i) & vbCrLf
        Next i
        Worksheets("DataSet").Cells(IntRow, valIntOne) = "R"
        Worksheets("DataSet").Cells(IntRow, IntCol + valIntOne) = Txt
    Next Counter
'--------------------------------------------------------------------------------------------------------
    'REM - Update Data from AscArray to DataSet
    
    

'--------------------------------------------------------------------------------------------------------
    'REM - Save and Close
    ThisWorkbook.Save
    'ThisWorkbook.Close
'--------------------------------------------------------------------------------------------------------
End Sub
