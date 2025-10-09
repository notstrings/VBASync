Attribute VB_Name = "mdlTable"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // テーブル操作
' // ・ListObject関連オブジェクト操作のヘルパモジュール
' // 20210901:初版
' // 20211006:テーブル行列の追加削除処理を追加
' // 202208xx:GetTableRows追加等
' // 20230523:Resize修正
' // 20240523:大きめのテコ入れ
' // 202509XX:実用版

' // テーブル全体 /////////////////////////////////////////////////////////////

' テーブル作成
Function AddTable(ByVal oSheet As Worksheet, ByVal oRange As Range, Optional ByVal sTable As String = "Table", Optional ByVal bHasHeader As Boolean = True) As ListObject
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
Sub DelTable(ByVal oSheet As Worksheet, Optional ByVal lIndex As Long = 1)
    GetTable(oSheet, lIndex).Unlist
End Sub
Sub DelTableByName(ByVal oSheet As Worksheet, Optional ByVal sTableName As String = "")
    GetTableByName(oSheet, sTableName).Unlist
End Sub

' テーブル取得(全)
Function GetTables(ByVal oSheet As Worksheet, Optional ByVal lIndex As Long = 1) As ListObjects
    Set GetTables = oSheet.ListObjects
End Function

' テーブル取得(単)
Function GetTable(ByVal oSheet As Worksheet, Optional ByVal lIndex As Long = 1) As ListObject
    Dim oRet As ListObject
    If lIndex >= 1 And oSheet.ListObjects.Count >= lIndex Then
        Set oRet = oSheet.ListObjects(lIndex)
    End If
    Set GetTable = oRet
End Function
Function GetTableByName(ByVal oSheet As Worksheet, Optional ByVal sName As String = "") As ListObject
    Dim oRet As ListObject
    Dim elm As ListObject
    For Each elm In oSheet.ListObjects
        If elm.Name = sName Then
            Set oRet = elm
            Exit For
        End If
    Next
    Set GetTableByName = oRet
End Function

' テーブル範囲全体＝構造化参照：Range("テーブル名[#All]")
Function GetTableRng(ByVal oListObject As ListObject) As Range
    Set GetTableRng = oListObject.Range
End Function

' テーブル範囲リサイズ
Sub ExtendTableRng(ByVal oListObject As ListObject, ByVal lTop As Long, ByVal lLeft As Long, ByVal lBottom As Long, lRight)
    Call oListObject.Resize(ExtendRange(oListObject.Range, lTop, lLeft, lBottom, lRight))
End Sub

' // テーブル列操作 ///////////////////////////////////////////////////////////

' テーブル列数
Function GetTableColCnt(ByVal oListObject As ListObject) As Long
    GetTableColCnt = oListObject.ListColumns.Count
End Function

' テーブル列番号
Function GetTableColIdx(ByVal oListObject As ListObject, ByVal sColumn As String) As Long
    GetTableColIdx = oListObject.ListColumns(sColumn).Index
End Function

' テーブル列取得(全)
Function GetTableCols(ByVal oListObject As ListObject) As ListColumns
    Set GetTableCols = oListObject.ListColumns
End Function

' テーブル列取得(単)
Function GetTableCol(ByVal oListObject As ListObject, ByVal sColumn As String) As ListColumn
    Dim oRet As ListColumn
    Dim elm As ListColumn
    For Each elm In oListObject.ListColumns
        If elm.Name = sColumn Then
            Set oRet = elm
            Exit For
        End If
    Next
    Set GetTableCol = oRet
End Function
Function GetTableColbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As ListColumn
    Dim oRet As ListColumn
    If lColIdx >= 1 And lColIdx <= oListObject.ListColumns.Count Then
        Set oRet = oListObject.ListColumns(lColIdx)
    End If
    Set GetTableColbyIdx = oRet
End Function

' テーブル列追加
Function AddTableCol(ByVal oListObject As ListObject, Optional ByVal sColumn As String = "", Optional ByVal lColIdx As Long = 0) As ListColumn
    Dim oRet As ListColumn
    If lColIdx >= 1 And lColIdx <= oListObject.ListColumns.Count Then
        Set oRet = oListObject.ListColumns.Add(lColIdx)
    Else
        Set oRet = oListObject.ListColumns.Add()
    End If
    If sColumn <> "" Then
        oRet.Name = GenUniqRowName(oListObject, sColumn, oRet.Name)
    End If
    Set AddTableCol = oRet
End Function

' テーブル列削除(列名称)
Sub DelTableCol(ByVal oListObject As ListObject, ByVal sColumn As String)
    oListObject.ListColumns(sColumn).Delete
End Sub

' テーブル列削除(列番号)
Sub DelTableColbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long)
    oListObject.ListColumns(lColIdx).Delete
End Sub

' テーブル内でユニークな列名を作成する
Private Function GenUniqRowName(ByVal oListObject As ListObject, ByVal sName As String, ByVal sIgnore As String) As String
    Dim txt As String
    Dim idx As Long
    txt = sName
    idx = 2
    Do
        Dim fnd As Boolean: fnd = False
        Dim elm As ListColumn
        For Each elm In oListObject.ListColumns
            If elm.Name = txt And elm.Name <> sIgnore Then
                fnd = True
                Exit For
            End If
        Next
        If fnd Then
            txt = sName & Format(idx, "00")
            idx = idx + 1
        End If
    Loop While fnd = True
    GenUniqRowName = txt
End Function

' // テーブル行操作 ///////////////////////////////////////////////////////////

' テーブル行数取得
Function GetTableRowCnt(ByVal oListObject As ListObject) As Long
    GetTableRowCnt = oListObject.ListRows.Count
End Function

' テーブル全行取得(ListRows)
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

' // テーブルヘッダー /////////////////////////////////////////////////////////
' // ・列タイトル行範囲※注意：ヘッダー非表示の場合は取得できない事

' ヘッダ部全体取得＝構造化参照：Range("テーブル名[#Header]")
Function GetTableHeaderRng(ByVal oListObject As ListObject) As Range
    Set GetTableHeaderRng = oListObject.HeaderRowRange
End Function

' ヘッダ部の指定セル取得(列名称)
Function GetTableHeaderCell(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Dim oRet As Range
    If Not oListObject.HeaderRowRange Is Nothing Then
        Dim elm As Range
        For Each elm In oListObject.HeaderRowRange
            If elm.Text = sColumn Then
                Set oRet = elm
                Exit For
            End If
        Next
    End If
    Set GetTableHeaderCell = oRet
End Function

' ヘッダ部の指定セル取得(列番号)
Function GetTableHeaderCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Dim oRet As Range
    If Not oListObject.HeaderRowRange Is Nothing Then
        If lColIdx >= 1 And lColIdx <= oListObject.HeaderRowRange.Count Then
            Set oRet = oListObject.HeaderRowRange(lColIdx)
        End If
    End If
    Set GetTableHeaderCellbyIdx = oRet
End Function

' ヘッダ部の指定セル取得(テーブル内任意セル)
Function GetTableHeaderCellbyCell(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Dim oRet As Range
    If Not oListObject.HeaderRowRange Is Nothing Then
        Set oRet = IntersectRange(oListObject.HeaderRowRange, oCell.EntireColumn)
    End If
    Set GetTableHeaderCellbyCell = oRet
End Function

' // テーブルフッタ－ /////////////////////////////////////////////////////////
' // ・集計行範囲※注意；フッタ－非表示の場合は取得できない

' フッタ部全体取得＝構造化参照：Range("テーブル名[#Totals]")
Function GetTableFooterRng(ByVal oListObject As ListObject) As Range
    Set GetTableFooterRng = oListObject.TotalsRowRange
End Function

' フッタ部の指定セル取得(列名称)
Function GetTableFooterCell(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Dim oRet As Range
    If Not oListObject.TotalsRowRange Is Nothing Then
        Dim elm As Range
        For Each elm In oListObject.TotalsRowRange
            If elm.Text = sColumn Then
                Set oRet = elm
                Exit For
            End If
        Next
    End If
    Set GetTableFooterCell = oRet
End Function

' フッタ部の指定セル取得(列番号)
Function GetTableFooterCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Dim oRet As Range
    If Not oListObject.TotalsRowRange Is Nothing Then
        If lColIdx >= 1 And lColIdx <= oListObject.TotalsRowRange.Count Then
            Set oRet = oListObject.TotalsRowRange(lColIdx)
        End If
    End If
    Set GetTableFooterCellbyIdx = oRet
End Function

' フッタ部の指定セル取得(テーブル内任意セル)
Function GetTableFooterCellbyElm(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Dim oRet As Range
    If Not oListObject.TotalsRowRange Is Nothing Then
        Set oRet = IntersectRange(oListObject.TotalsRowRange, oCell.EntireColumn)
    End If
    Set GetTableFooterCellbyElm = oRet
End Function

' // テーブルデータ部 /////////////////////////////////////////////////////////

' データ行列数
Function GetTableBodyColCnt(ByVal oListObject As ListObject) As Long
    GetTableBodyColCnt = oListObject.ListColumns.Count
End Function
Function GetTableBodyRowCnt(ByVal oListObject As ListObject) As Long
    GetTableBodyRowCnt = oListObject.ListRows.Count
End Function

' テーブル座標系変換
Function GetTableBodyRowIdx(ByVal oListObject As ListObject, ByVal oRange As Range) As Long
    GetTableBodyRowIdx = oRange.Row - oListObject.DataBodyRange.Row + 1
End Function
Function GetTableBodyColIdx(ByVal oListObject As ListObject, ByVal oRange As Range) As Long
    GetTableBodyColIdx = oRange.Column - oListObject.DataBodyRange.Column + 1
End Function

' データ範囲クリップ
Function ClipTableBodyRange(ByVal oListObject As ListObject, ByVal oRange As Range) As Range
    Set ClipTableBodyRange = IntersectRange(oListObject.DataBodyRange, oRange)
End Function

' データ部消去
Sub ClrTableBody(ByVal oListObject As ListObject)
    If Not oListObject.DataBodyRange Is Nothing Then
        oListObject.DataBodyRange.Delete
    End If
End Sub

' データ部全体＝構造化参照：Range("テーブル名[#Data]")
Function GetTableBodyRng(ByVal oListObject As ListObject) As Range
    Set GetTableBodyRng = oListObject.DataBodyRange
End Function
'Sub sample()
'    Dim oDataTable As ListObject
'    Set oDataTable = GetTable(ThisWorkbook.ActiveSheet)
'    If oDataTable Is Nothing Then Exit Sub                          ' テーブルがないので終了
'    Dim oTarget As Range
'    Set oTarget = Selection                                         ' テーブルに関係ないセル範囲
'    Set oTarget = ClipTableBodyRange(oDataTable, oTarget)           ' テーブルに関係ないセル範囲をテーブル範囲内にクリップ
'    If oTarget Is Nothing Then Exit Sub                             ' テーブル範囲外なので終了
'    Dim lc As Long: lc = GetTableBodyColIdx(oDataTable, oTarget)    ' テーブル内部でのインデックス番号に変換
'    Dim lr As Long: lr = GetTableBodyRowIdx(oDataTable, oTarget)    ' テーブル内部でのインデックス番号に変換
'    GetTableBodyRng(oDataTable)(lr, lc) = "aaa"                     ' 位置指定アクセス
'End Sub

' データ部の指定列
Function GetTableBodyColRng(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableBodyColRng = oListObject.ListColumns(sColumn).DataBodyRange
End Function
Function GetTableBodyColRngSafe(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Dim oRet As Range
    Dim elm As ListColumn
    For Each elm In oListObject.ListColumns
        If elm.Name = sColumn Then
            Set oRet = elm.DataBodyRange
            Exit For
        End If
    Next
    Set GetTableBodyColRngSafe = oRet
End Function

' データ部の指定行
Function GetTableBodyRowRng(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As Range
    Set GetTableBodyRowRng = oListObject.ListRows(lRowIdx).Range
End Function

' データ部の指定セル(行番号/列名称)
Function GetTableBodyCell(ByVal oListObject As ListObject, ByVal lRowIdx As Long, ByVal sColumn As String) As Range
    Set GetTableBodyCell = oListObject.ListColumns(sColumn).DataBodyRange(lRowIdx)
End Function
Function GetTableBodyCellSafe(ByVal oListObject As ListObject, ByVal lRowIdx As Long, ByVal sColumn As String) As Range
    Dim oRet As Range
    Dim elm As ListColumn
    For Each elm In oListObject.ListColumns
        If elm.Name = sColumn Then
            Set oRet = elm.DataBodyRange(lRowIdx)
            Exit For
        End If
    Next
    Set GetTableBodyCellSafe = oRet
End Function

' 探索

' ルックアップ(１次元)
Function LookupTableBody(ByVal oListObject As ListObject, ByVal sSearchCol As String, ByVal sResultCol As String, ByVal vSearchVal As Variant) As Variant
    Dim vRet As Variant
    Dim vValues As Variant
    vValues = oListObject.DataBodyRange.Value
    Dim lRow As Long
    For lRow = LBound(vValues) To UBound(vValues)
        If vValues(lRow, oListObject.ListColumns(sSearchCol).Index) = vSearchVal Then
            vRet = vValues(lRow, oListObject.ListColumns(sResultCol).Index)
            Exit For
        End If
    Next
    LookupTableBody = vRet
End Function

' ルックアップ(２次元)
Function LookupTableBodyDictGen(ByVal oListObject As ListObject, ByVal sMainColKey As String) As Object
    Dim oRet As Object
    Set oRet = CreateObject("Scripting.Dictionary")
    Dim vValues As Variant
    Dim lCol As Long
    Dim lRow As Long
    vValues = oListObject.DataBodyRange.Value
    For lRow = LBound(vValues, 1) To UBound(vValues, 1)
        For lCol = LBound(vValues, 2) To UBound(vValues, 2)
            Dim sKey As String
            Dim vVal As Variant
            sKey = vValues(lRow, oListObject.ListColumns(sMainColKey).Index) & "@" & oListObject.ListColumns(lCol).Name
            vVal = vValues(lRow, lCol)
            Call oRet.Add(sKey, vVal)
        Next
    Next
    Set LookupTableBodyDictGen = oRet
End Function
Function LookupTableBodyDict(ByVal oDict As Object, ByVal sRowKey As String, ByVal sColKey As String) As Variant
    Dim vRet As Variant
    Dim sKey As String
    sKey = sRowKey & "@" & sColKey
    If oDict.Exists(sKey) Then
        vRet = oDict(sKey)
    End If
    LookupTableBodyDict = vRet
End Function

' // テーブルデータ部フィルタ /////////////////////////////////////////////////
' // ・手動/コードでフィルタしたデータもテーブルデータ部取得関連処理では普通に参照できる
' // ・手動/コードでフィルタしたデータを無視する場合はSpecialCells(xlCellTypeVisible)で切り分け

' テーブルにフィルタを適用
Sub AddTableFilter(ByVal oListObject As ListObject, ByVal sColumn As String, ByVal sCriteria As String, Optional ByVal lOperator As XlAutoFilterOperator = xlAnd)
    Call oListObject.Range.AutoFilter(oListObject.ListColumns(sColumn).Index, sCriteria, lOperator)
End Sub

' テーブルのフィルタをクリア
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

' // テーブルデータ部ソート ///////////////////////////////////////////////////

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

' // テーブルの右クリックメニュー /////////////////////////////////////////////
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

