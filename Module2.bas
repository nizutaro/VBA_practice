Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveWorkbook.Queries.Add Name:="dummy", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    �\�[�X = Csv.Document(File.Contents(""C:\Users\test""),[Delimiter="","", Columns=11, Encoding=932, QuoteStyle=QuoteStyle.None])," & Chr(13) & "" & Chr(10) & "    ���i���ꂽ�w�b�_�[�� = Table.PromoteHeaders(�\�[�X, [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10) & "    �ύX���ꂽ�^ = Table.TransformColumnTypes(���i���ꂽ�w�b�_�[��,{{""���O"", type text}, {""�ӂ肪��"", type text}, {""�A�h���X"", type text}, {""����"", type" & _
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
End Sub
