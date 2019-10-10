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
    Dim valIntOne As Integer, IntValThree As Integer
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
    IntValThree = 3
    IntValCountOfRecordtypes = 16 'Types of record in SRC file other than HD & TR
    ReDim RecTypeTable(16)
    'SWorkSheetdefOP = ThisWorkbook.Worksheets("def_OP")
'--------------------------------------------------------------------------------------------------------
    'REM - Create Required Sheets
    CreateSheet "RecordDefinations"
    CreateSheet "AscArray"
    CreateSheet "DataSet"
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
        Txt = Mid(arr(i), valIntOne, IntValThree)
        If Len(Txt) >= IntValThree Then
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
            Txt = Txt & "Trailer Found POS : " & POS_TR & vbCrLf
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
            If (Len(arr(i)) >= IntValThree) And (Len(arr(j)) >= IntValThree) Then
                If (Trim(Mid(arr(i), valIntOne, IntValThree)) <> "HD") And (Trim(Mid(arr(i), valIntOne, IntValThree)) <> "TR") _
                    And (Trim(Mid(arr(j), valIntOne, IntValThree)) <> "HD") And (Trim(Mid(arr(j), valIntOne, IntValThree)) <> "TR") Then
                    If Int(Mid(arr(i), valIntOne, IntValThree)) < Int(Mid(arr(j), valIntOne, IntValThree)) Then
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
        For i = IntValThree To IntValThree
            'Txt = Txt & HeaderDefn(0, i) & " :: " & HeaderDefn(IntRow, i) & vbCrLf
            Txt = Txt & HeaderDefn(IntRow, i) & vbCrLf
        Next i
        Worksheets("DataSet").Cells(IntRow, valIntOne) = "R"
        Worksheets("DataSet").Cells(IntRow, IntCol + valIntOne) = Txt
    Next Counter
'--------------------------------------------------------------------------------------------------------
    'REM - Update Data from AscArray to DataSet
    
    Dim Data_Type As String, Header_Type As String
    
    
    For i = 1 To Worksheets("AscArray").Cells(Rows.Count, valIntOne).End(xlUp).Row
        For j = 1 To Worksheets("DataSet").Cells(Rows.Count, valIntOne).End(xlUp).Row
        
        If (Len(Worksheets("AscArray").Cells(i, 2)) > 3 And Len(Worksheets("DataSet").Cells(i, 2)) > 3) Then
        
            Data_Type = Mid(Worksheets("AscArray").Cells(i, 2), 1, 3)
            If Worksheets("DataSet").Cells(j, 2) = "D" Then
                x = 1
            Else
                x = 2
            End If
            Header_Type = Mid(Worksheets("DataSet").Cells(j, 2), x, 3)
            
            If Trim(Data_Type) <> "HD" And _
                Trim(Data_Type) <> "TR" And _
                Trim(Header_Type) <> "HD" And _
                Trim(Header_Type) <> "TR" And _
                Trim(Header_Type) <> "" Then
                
                    If (Int(Data_Type) = Int(Header_Type)) Then
                    
                    'Insert Row Above Row 3
                    Worksheets("DataSet").Rows(j + 1).Insert Shift:=xlUp, _
                            CopyOrigin:=xlFormatFromLeftOrAbove 'xlFormatFromLeftOrAbove 'or xlFormatFromRightOrBelow
                    Worksheets("DataSet").Cells(j + 1, valIntOne) = "D"
                    Worksheets("DataSet").Cells(j + 1, 2) = Worksheets("AscArray").Cells(i, 2)
                
                    End If
                
                
            End If
        
        End If
        
        

        
        
        
        
        Next j
    Next i
    
    
    
    

'--------------------------------------------------------------------------------------------------------
    'REM - Save and Close
    ThisWorkbook.Save
    'ThisWorkbook.Close
'--------------------------------------------------------------------------------------------------------
End Sub
