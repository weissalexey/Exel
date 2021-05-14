VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

    Dim RepWB As Workbook

 Set RepWB = ThisWorkbook
    Dim RepSht As Excel.Worksheet
    Set RepSht = RepWB.Sheets("Перечень все проводок")
    RepSht.Range("date_pol") = Date - 1

End Sub


Sub Макрос1()
'
' Макрос1 Макрос
' Макрос записан 15/05/2006 (iossel-ma)
'
Dim FirstDate As Date
Dim contdate As Long
Dim str As String
Dim RepWB As Workbook

 Set RepWB = ThisWorkbook
 
  Dim RepSht As Excel.Worksheet
 Set RepSht = RepWB.Sheets("Перечень все проводок")
 
 RepSht.Range("A5:AG60000").Select

    Selection.ClearContents
    RepSht.Range("A5").Select
'    Selection.QueryTable.Delete
FirstDate = CDate(Replace(Replace(RepSht.Range("date_pol"), ".", "-"), "/", "-"))
contdate = Year(FirstDate) * 10000 + Month(FirstDate) * 100 + Day(FirstDate)
str = CStr(contdate)

'
    With RepSht.QueryTables.Add(Connection:=Array(Array( _
        "ODBC;DSN=con;APP=Легал Отбор По словам;DATABASE=DBA_1;Network=DBMSSOCN;A" _
        ), Array("utoTranslate=No;Trusted_Connection=Yes;QuotedId=No;AnsiNPW=No")), _
        Destination:=Range("A5"))
        .CommandText = Array( _
        "SELECT * from dba_1..[Cart_von_mench]" _
        )
        .Name = "Мой запрос"
        .FieldNames = True
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
        .Refresh BackgroundQuery:=False
    End With
    
    Selection.AutoFilter
End Sub


