Attribute VB_Name = "Module1"
Option Explicit

Sub getCSV()
    '=====================
    'CSVをループして使用して取り込むサンプル
    '=====================
    
    Dim timerStart As Long
    timerStart = Timer()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    
    Dim strPath As String
    strPath = "C:\Users\test\OneDrive\Desktop\dummy.csv" 'フォルダのPathを入力
    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant 'カンマでsplitして格納
    
    Open strPath For Input As #1 'CSVをオープン
    
    i = 1
    
    Do Until EOF(1)
    
        Line Input #1, strLine
        arrLine = Split(strLine, ",") 'strLineをカンマ区切りarrLineに格納
        For j = 0 To UBound(arrLine)
            ws.Cells(i, j + 1).Value = arrLine(j)
        Next j
        i = i + 1
        Loop
        
        Close #1
    
        Debug.Print Timer() - timerStart & "秒経過"
End Sub


Private Sub csvImport()
 Dim timerStart As Long
 timerStart = Timer()
    Dim strPath As String
    Dim qtCsv As QueryTable
    
    strPath = "C:\Users\test\OneDrive\Desktop\dummy.csv"
    Set qtCsv = Sheet1.QueryTables.Add(Connection:="TEXT;" & strPath, _
    Destination:=Sheet1.Range("A1")) '取り込むCSVパスと、取り込み先のシート、セルを指定
    
    With qtCsv
        .TextFileCommaDelimiter = True 'カンマ区切りの指定
        .TextFileParseType = xlDelimited '区切り文字の形式
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileStartRow = 1 '開始行の指定
        .TextFileTextQualifier = xlTextQualifierDoubleQuote '引用符の指定
        .TextFilePlatform = 932 '文字コード指定
        .Refresh 'QueryTablesオブジェクトを更新し、シート上に出力
        .Delete 'QueryTables.Addメソッドで取り込んだCSVとの接続を解除
    End With
     Debug.Print Timer() - timerStart & "秒経過"
End Sub

Sub クエリ()
 Dim timerStart As Long
 timerStart = Timer()

    ActiveWorkbook.Queries.Add Name:="dummy", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    ソース = Csv.Document(File.Contents(""C:\Users\test\OneDrive\Desktop\dummy.csv""),[Delimiter="","", Columns=11, Encoding=932, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    昇格されたヘッダー数 = Table.PromoteHeaders(ソース, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    変更された型 = Table.TransformColumnTypes(昇格されたヘッダー数,{{""名前"", type text}, {""ふりがな"", type text}, {""アドレス"", type text}, {""性別"", type" & _
        " text}, {""年齢"", Int64.Type}, {""誕生日"", type date}, {""婚姻"", type text}, {""都道府県"", type text}, {""携帯"", type text}, {""キャリア"", type text}, {""カレーの食べ方"", type text}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    変更された型" & _
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
Debug.Print Timer() - timerStart & "秒経過"
End Sub

