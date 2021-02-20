Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="dummy", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    ソース = Csv.Document(File.Contents(""C:\Users\test""),[Delimiter="","", Columns=11, Encoding=932, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    昇格されたヘッダー数 = Table.PromoteHeaders(ソース, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    変更された型 = Table.TransformColumnTypes(昇格されたヘッダー数,{{""名前"", type text}, {""ふりがな"", type text}, {""アドレス"", type text}, {""性別"", type" & _
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
End Sub
