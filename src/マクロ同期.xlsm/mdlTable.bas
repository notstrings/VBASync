Attribute VB_Name = "mdlTable"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // テーブル操作
' // ・ListObject関連オブジェクト操作のヘルパモジュール
' // ・ワークシートとテーブルは１対１に対応する事を原則とする
' // 20210901:初版
' // 20211006:テーブル行列の追加削除処理を追加
' // 202208xx:GetTableRows追加等
' // 20230523:Resize修正
' // 20240523:大きめのテコ入れ

' // テーブル操作 /////////////////////////////////////////

' テーブル作成
Function AddTable(ByVal oSheet As Worksheet, ByVal oRange As Range, Optional ByVal bHasHeader As Boolean = True) As ListObject
    Dim oRet As ListObject
    Set oRet = oSheet.ListObjects.Add(SourceType:=xlSrcRange, source:=oRange, XlListObjectHasHeaders:=IIf(bHasHeader, xlYes, xlNo))
    Set AddTable = oRet
End Function
Function AddTableWithRecordset(ByVal oSheet As Worksheet, ByVal oRange As Range, ByVal oSrcRecordset As Object) As ListObject
    Dim oRet As ListObject
    Set oRet = oSheet.ListObjects.Add(SourceType:=xlSrcQuery, source:=oSrcRecordset, LinkSource:=False, XlListObjectHasHeaders:=xlYes, Destination:=oRange)
    Call oRet.QueryTable.Refresh
    Set AddTableWithRecordset = oRet
End Function

' テーブル削除
Sub DelTable(ByVal oSheet As Worksheet)
    GetTable(oSheet).Unlist
End Sub
Sub DelTableByName(ByVal oSheet As Worksheet, Optional ByVal sTableName As String = "")
    GetTableByName(oSheet, sTableName).Unlist
End Sub

' テーブル取得
Function GetTable(ByVal oSheet As Worksheet) As ListObject
    Dim oRet As ListObject
    If oSheet.ListObjects.Count > 0 Then
        Set oRet = oSheet.ListObjects(1)
    End If
    Set GetTable = oRet
End Function
Function GetTableByName(ByVal oSheet As Worksheet, Optional ByVal sName As String = "") As ListObject
    Dim oRet As ListObject
    Dim elm As ListObject
    For Each elm In oSheet.ListObjects
        If elm.Name Like sName Then
            Set oRet = elm
            Exit For
        End If
    Next
    Set GetTableByName = oRet
End Function

' // テーブル全体

' テーブル全体取得      ■構造化参照：Range("テーブル名[#All]")
Function GetTableRng(ByVal oListObject As ListObject) As Range
    Set GetTableRng = oListObject.Range
End Function

' テーブルリサイズ
Sub ExtendTable(ByVal oListObject As ListObject, ByVal lTop As Long, ByVal lLeft As Long, ByVal lBottom As Long, lRight)
    Call oListObject.Resize(ExtendRange(oListObject, lTop, lLeft, lBottom, lRight))
End Sub

' // テーブル全体の列操作
' // ・全テーブル範囲制御(概念的にヘッダフッタを含む)

' テーブル列数取得
Function GetTableColCnt(ByVal oListObject As ListObject) As Long
    GetTableColCnt = oListObject.ListColumns.Count
End Function

' テーブル列番号取得
Function GetTableColIdx(ByVal oListObject As ListObject, ByVal sColumn As String) As Long
    GetTableColIdx = oListObject.ListColumns(sColumn).Index
End Function

' テーブル全列取得
Function GetTableCols(ByVal oListObject As ListObject) As ListObject
    Set GetTableCols = oListObject.ListColumns
End Function

' テーブル列取得(ListColumn)
Function GetTableCol(ByVal oListObject As ListObject, ByVal sColumn As String) As ListColumn
    Set GetTableCol = oListObject.ListColumns(sColumn)
End Function

' テーブル列追加
Function AddTableCol(ByVal oListObject As ListObject) As ListColumn
    Set AddTableCol = oListObject.ListColumns.Add()
End Function

' テーブル列挿入
Function InsTableCol(ByVal oListObject As ListObject, ByVal lPosition As Long, Optional ByVal sColumn As String = "") As ListColumn
    Dim oRet As ListColumn
    Set oRet = oListObject.ListColumns.Add(lPosition)
    If sColumn <> "" Then
        oRet.Name = sColumn
    End If
    Set InsTableCol = oRet
End Function

' テーブル列削除(列名称)
Sub DelTableCol(ByVal oListObject As ListObject, ByVal sColumn As String)
    oListObject.ListColumns(sColumn).Delete
End Sub

' テーブル列削除(列番号)
Sub DelTableColbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long)
    oListObject.ListColumns(lColIdx).Delete
End Sub

' // テーブル全体の行操作
' // ・全テーブル範囲制御(概念的にヘッダフッタを含む)

' テーブル行数取得
Function GetTableRowCnt(ByVal oListObject As ListObject) As Long
    GetTableRowCnt = oListObject.ListRows.Count
End Function

' テーブル全行取得
Function GetTableRows(ByVal oListObject As ListObject) As ListRows
    Set GetTableRows = oListObject.ListRows
End Function

' テーブル行取得(ListRow)
Function GetTableRow(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As ListRow
    Set GetTableRow = oListObject.ListRows(lRowIdx)
End Function

' テーブル行追加
Function AddTableRow(ByVal oListObject As ListObject) As ListRow
    Set AddTableRow = oListObject.ListRows.Add()
End Function

' テーブル行挿入
Function InsTableRow(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As ListRow
    Set InsTableRow = oListObject.ListRows.Add(lRowIdx, True)
End Function

' テーブル行削除
Sub DelTableRow(ByVal oListObject As ListObject, ByVal lPosition As Long)
    oListObject.ListRows(lPosition).Delete
End Sub

' テーブル全行消去
Sub ClrTable(ByVal oListObject As ListObject)
    If Not oListObject.DataBodyRange Is Nothing Then
        oListObject.DataBodyRange.Delete
    End If
End Sub

' // テーブルヘッダー /////////////////////////////////
' // ・列タイトルのコト

' ヘッダ列数取得
Function GetTableHeaderCnt(ByVal oListObject As ListObject) As Long
    GetTableHeaderCnt = oListObject.HeaderRowRange.Count
End Function

' ヘッダ部全体取得      ■構造化参照：Range("テーブル名[#Header]")
Function GetTableHeaderRng(ByVal oListObject As ListObject) As Range
    Set GetTableHeaderRng = oListObject.HeaderRowRange
End Function

' ヘッダ部の指定セル取得(列名称)
Function GetTableHeaderCell(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableHeaderCell = oListObject.ListColumns(sColumn).Range(1, 1)
End Function

' ヘッダ部の指定セル取得(列番号)
Function GetTableHeaderCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Set GetTableHeaderCellbyIdx = oListObject.HeaderRowRange(lColIdx)
End Function

' ヘッダ部の指定セル取得(テーブル内任意セル)
Function GetTableHeaderCellbyElm(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Set GetTableHeaderCellbyElm = Intersect(oListObject.HeaderRowRange, oCell.EntireColumn) ' 指定セル列とテーブルヘッダ領域の積集合を取った領域
End Function

' // テーブルフッタ－ /////////////////////////////////
' // ・集計行のコト

' フッタ列数取得
Function GetTableFooterCnt(ByVal oListObject As ListObject) As Long
    GetTableFooterCnt = oListObject.TotalsRowRange.Count
End Function

' フッタ部全体取得      ■構造化参照：Range("テーブル名[#Totals]")
Function GetTableFooterRng(ByVal oListObject As ListObject) As Range
    Set GetTableFooterRng = oListObject.TotalsRowRange
End Function

' フッタ部の指定セル取得(列名称)
Function GetTableFooterCell(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableFooterCell = Intersect(oListObject.TotalsRowRange, oListObject.ListColumns(sColumn).Range)
End Function

' フッタ部の指定セル取得(列番号)
Function GetTableFooterCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Set GetTableFooterCellbyIdx = oListObject.TotalsRowRange(lColIdx)
End Function

' フッタ部の指定セル取得(テーブル内任意セル)
Function GetTableFooterCellbyElm(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Set GetTableFooterCellbyElm = Intersect(oListObject.TotalsRowRange, oCell.EntireColumn) ' 指定セル列とテーブルヘッダ領域の積集合を取った領域
End Function

' // テーブルデータ部 /////////////////////////////////

' データ行数
Function GetTableBodyCnt(ByVal oListObject As ListObject) As Long
    GetTableBodyCnt = oListObject.ListRows.Count
End Function

' データ部全体          ■構造化参照：Range("テーブル名[#Data]")
Function GetTableBodyRng(ByVal oListObject As ListObject) As Range
    Set GetTableBodyRng = oListObject.DataBodyRange
End Function

' データ部の指定列
Function GetTableBodyColRng(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableBodyColRng = oListObject.ListColumns(sColumn).DataBodyRange
End Function

' データ部の指定行
Function GetTableBodyRowRng(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As Range
    Set GetTableBodyRowRng = oListObject.ListRows(lRowIdx).Range
End Function

' データ部の指定セル(行番号/列名称)
Function GetTableBodyCell(ByVal oListObject As ListObject, ByVal lRowIdx As Long, ByVal sColumn As String) As Range
    Set GetTableBodyCell = oListObject.ListColumns(sColumn).DataBodyRange(lRowIdx)
End Function

' データ部の指定セル(行番号/列番号)
Function GetTableBodyCellByIdx(ByVal oListObject As ListObject, ByVal lRowIdx As Long, ByVal lColIdx As Long) As Range
    Set GetTableBodyCellByIdx = oListObject.DataBodyRange.Cells(lRowIdx, lColIdx)
End Function

' データ部の探索カラムに対応する結果カラムを返却

' ルックアップ版
Function LookupTableBody(ByVal oListObject As ListObject, ByVal sSearchCol As String, ByVal sResultCol As String, ByVal vSearchVal As Variant) As Variant
    Dim vRet As Variant
    Dim lRow As Long
    For lRow = 1 To oListObject.ListRows.Count
        If oListObject.ListColumns(sSearchCol).DataBodyRange(lRow) = vSearchVal Then
            vRet = oListObject.ListColumns(sResultCol).DataBodyRange(lRow)
            Exit For
        End If
    Next
    LookupTableBody = vRet
End Function

' ディクショナリ版
Function MakeTableBodyDict(ByVal oListObject As ListObject, ByVal sKeyCol As String) As Object
    Dim oRet As Object
    Set oRet = CreateObject("Scripting.Dictionary")
    Dim elm As Variant
    Dim lKey As Long
    Dim sKey As String
    lKey = oListObject.ListColumns(sKeyCol).Index
    For Each elm In oListObject.ListRows
        sKey = elm.Range(lKey).Text
        If oRet.Exists(sKey) = False Then
            Call oRet.Add(sKey, elm.Range.Value)
        End If
    Next
    Set MakeTableBodyDict = oRet
End Function
Function ResolveTableBodyDict(ByVal oListObject As ListObject, ByVal oDict As Object, ByVal sKey As String, ByVal sRef As String) As Variant
    Dim vRet As Variant
    If oDict.Exists(sKey) Then
        vRet = oDict.Item(sKey)(1, GetTableColIdx(oListObject, sRef))
    End If
    ResolveTableBodyDict = vRet
End Function

' // テーブルデータ部フィルタ /////////////////////////////
' // ・フィルタ設定はExcel機能そのものであるため、かなり処理が重いことに注意
' // ・手動/コードでフィルタ(非表示状態に)したデータもテーブルデータ部取得関連処理では普通に参照できる
' // ・手動/コードでフィルタ(非表示状態に)したデータを無視したい場合はSpecialCells(xlCellTypeVisible)で切り分ける

' テーブルにフィルタを適用
Sub AddTableFilter(ByVal oListObject As ListObject, ByVal sColumn As String, ByVal sCriteria As String, Optional ByVal lOperator As XlAutoFilterOperator = xlAnd)
    Call oListObject.Range.AutoFilter(oListObject.ListColumns(sColumn).Index, sCriteria, lOperator)
End Sub

' テーブルのフィルタをクリア
' Excel機能のフィルタは重いので注意
Sub ClrTableRowFilter(ByVal oListObject As ListObject)
    oListObject.AutoFilter.ShowAllData
End Sub

' テーブルに設定するフィルタ文字列をサニタイズ
Function SanitizeFilterText(ByVal sText) As String
    sText = Replace(sText, "*", "~*")
    sText = Replace(sText, "?", "~?")
    sText = Replace(sText, "=", "==")
    sText = Replace(sText, "<", "=<")
    sText = Replace(sText, ">", "=>")
    SanitizeFilterText = sText
End Function

' // テーブルデータ部ソート ///////////////////////////////

' ソート実施
Sub ApplyTableSort(ByVal oListObject As ListObject, ByVal sColumn As String, Optional ByVal bMatchCase As Boolean = True, Optional ByVal lSortMethod As XlSortMethod = xlPinYin)
    With oListObject.Sort
        .MatchCase = bMatchCase
        .SortMethod = lSortMethod
        .Apply
    End With
End Sub

' ソート条件追加
Sub AddTableSort(ByVal oListObject As ListObject, ByVal sColumn As String, Optional ByVal lOrder As XlSortOrder = xlAscending, Optional ByVal lSortOn As XlSortOn = xlSortOnValues, Optional lDataOption As XlSortDataOption = xlSortTextAsNumbers)
    Call oListObject.Sort.SortFields.Add(key:=oListObject.ListColumns(sColumn).Range, SortOn:=lSortOn, Order:=lOrder, DataOption:=lDataOption)
End Sub

' ソート条件消去
Sub ClrTableSort(ByVal oListObject As ListObject, ByVal sColumn As String)
    Call oListObject.Sort.SortFields.Clear
End Sub

' // テーブルの右クリックメニュー /////////////////////////
' // ・普通のセル上での右クリックには反応しない

' 右クリックメニュー追加
Sub AddTableRClickMenu(ByVal sTitle As String, ByVal sMacro As String)
    Dim oCmdBar As CommandBarButton
    Set oCmdBar = Application.CommandBars("List Range Popup").Controls.Add(Temporary:=True)
    With oCmdBar
        .Caption = sTitle
        .OnAction = sMacro
    End With
End Sub

' 右クリックメニュークリア
Sub ClrTableRClickMenu()
    Call Application.CommandBars("List Range Popup").Reset
End Sub
