Attribute VB_Name = "Module1"
Option Explicit

Sub getCSV()
    '=====================
    'CSV�����[�v���Ďg�p���Ď�荞�ރT���v��
    '=====================
    
    Dim timerStart As Long
    timerStart = Timer()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    Dim strPath As String
    strPath = "C:\Users\test\OneDrive\Desktop\dummy.csv" '�t�H���_��Path�����
    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant '�J���}��split���Ċi�[
    
    Open strPath For Input As #1 'CSV���I�[�v��
    
    i = 1
    
    Do Until EOF(1)
    
        Line Input #1, strLine
        arrLine = Split(strLine, ",") 'strLine���J���}��؂�arrLine�Ɋi�[
        For j = 0 To UBound(arrLine)
            ws.Cells(i, j + 1).Value = arrLine(j)
        Next j
        i = i + 1
        Loop
        
        Close #1
    
        Debug.Print Timer() - timerStart & "�b�o��"
End Sub


Private Sub csvImport()
 Dim timerStart As Long
 timerStart = Timer()
    Dim strPath As String
    Dim qtCsv As QueryTable
    
    strPath = "C:\Users\test\OneDrive\Desktop\dummy.csv"
    Set qtCsv = Sheet1.QueryTables.Add(Connection:="TEXT;" & strPath, _
    Destination:=Sheet1.Range("A1")) '��荞��CSV�p�X�ƁA��荞�ݐ�̃V�[�g�A�Z�����w��
    
    With qtCsv
        .TextFileCommaDelimiter = True '�J���}��؂�̎w��
        .TextFileParseType = xlDelimited '��؂蕶���̌`��
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileStartRow = 1 '�J�n�s�̎w��
        .TextFileTextQualifier = xlTextQualifierDoubleQuote '���p���̎w��
        .TextFilePlatform = 932 '�����R�[�h�w��
        .Refresh 'QueryTables�I�u�W�F�N�g���X�V���A�V�[�g��ɏo��
        .Delete 'QueryTables.Add���\�b�h�Ŏ�荞��CSV�Ƃ̐ڑ�������
    End With
     Debug.Print Timer() - timerStart & "�b�o��"
End Sub

Sub �N�G��()
 Dim timerStart As Long
 timerStart = Timer()

    ActiveWorkbook.Queries.Add Name:="dummy", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    �\�[�X = Csv.Document(File.Contents(""C:\Users\test\OneDrive\Desktop\dummy.csv""),[Delimiter="","", Columns=11, Encoding=932, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    ���i���ꂽ�w�b�_�[�� = Table.PromoteHeaders(�\�[�X, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    �ύX���ꂽ�^ = Table.TransformColumnTypes(���i���ꂽ�w�b�_�[��,{{""���O"", type text}, {""�ӂ肪��"", type text}, {""�A�h���X"", type text}, {""����"", type" & _
        " text}, {""�N��"", Int64.Type}, {""�a����"", type date}, {""����"", type text}, {""�s���{��"", type text}, {""�g��"", type text}, {""�L�����A"", type text}, {""�J���[�̐H�ו�"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    �ύX���ꂽ�^" & _
        ""
    Sheets.Add After:=ActiveSheet
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=dummy;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [dummy]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "dummy"
        .Refresh BackgroundQuery:=False
    End With
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Range("A5").Select
Debug.Print Timer() - timerStart & "�b�o��"
End Sub

