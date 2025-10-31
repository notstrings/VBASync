Attribute VB_Name = "mdlCommon"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // どこにでも使いそうな雑多な処理
' // 20210901:初版
' // 202208xx:InputNum不具合修正
' // 20221101:レーベンシュタイン距離追加
' // 20230221:機能追加(SetStatusBar追加/SearchBook追加/CleansingText削除/etc)
' // 20230308:Office2010用のWebService関数を暫定作成
' // 20230308:１セル１文字書式用処理追加
' // 20230308:Excel2021では微妙なのでウィンドウ整列削除
' // 20230412:シート表示状態設定追加
' // 20230720:シート追加削除複写追加
' // 20240401:ブック関連操作を拡充
' //          シート・ワークシート・チャートシートの扱いを明確に分離
' //          ワークシート追加/複製時の挙動を調整
' //          通常検索追加&簡易検索廃止
' //          Rangeの補集合/差集合算出機能追加
' //          １セル１文字形式で"'"と"="についての制御を修正
' //          レーベンシュタイン比率算出追加
' //          コメント設定関連処理追加
' //          ハイパーリンク関連処理追加
' //          WebService用改めWebAPI用処理各種改修
' //          ByVal/ByRefや戻り値の型指定を徹底
' // 20251009:正規表現ライブラリ差し替え＆ライセンス的に微妙なので分離

Public Const csELPtrn As String = "[" & vbCr & vbLf & "]"   ' TrimEx用

Public Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer

' // 共用 /////////////////////////////////////////////////

' FileSystemObject生成
Private Function CreateFSO() As Object
    Static oFSO As Object
    If oFSO Is Nothing Then
        Set oFSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set CreateFSO = oFSO
End Function

' // Excel一般操作 ////////////////////////////////////////

' 情報表示用メッセージボックス
Function InfBox(ByVal sTitle As String, ByVal sMessage As String) As Long
    InfBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKOnly Or vbInformation)
End Function

' 意思確認用メッセージボックス
Function AskBox(ByVal sTitle As String, ByVal sMessage As String, Optional ByVal bDefaultOK As Boolean = False) As Long
    If bDefaultOK Then
        AskBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKCancel Or vbInformation)
    Else
        AskBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKCancel Or vbDefaultButton2 Or vbExclamation)
    End If
End Function

' エラー表示用メッセージボックス
Function ErrBox(ByVal sTitle As String, ByVal sMessage As String) As Long
    ErrBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKOnly Or vbCritical Or vbSystemModal)
End Function

' 入力ボックス(数値)
Function InputNum(ByVal sTitle As String, ByVal sMessage As String, ByVal vDefault As Variant, ByRef vResult As Variant) As Boolean
    Dim bRet As Boolean
    Dim vVal As Variant
    Dim sDef As String
    sDef = Val(vDefault)
    vVal = Application.InputBox(Prompt:=sMessage, Title:=sTitle, Default:=sDef, Type:=1)
    If VarType(vVal) <> vbBoolean Then
        bRet = True
        vResult = vVal
    Else
        bRet = False
        vResult = sDef
    End If
    InputNum = bRet
End Function

' 入力ボックス(テキスト)
Function InputTxt(ByVal sTitle As String, ByVal sMessage As String, ByVal sDefault As String, ByRef sResult As String) As Boolean
    Dim bRet As Boolean
    Dim vVal As Variant
    vVal = Application.InputBox(Prompt:=sMessage, Title:=sTitle, Default:=sDefault, Type:=2)
    If VarType(vVal) <> vbBoolean Then
        bRet = True
        sResult = vVal
    Else
        bRet = False
        sResult = sDefault
    End If
    InputTxt = bRet
End Function

' 入力ボックス(範囲)
Function InputRng(ByVal sTitle As String, ByVal sMessage As String, ByVal oDefault As Range, ByRef sResult As Range) As Boolean
    Dim bRet As Boolean
    Dim oVal As Range
On Error GoTo ErrExit
    Dim sDef As String
    If Not oDefault Is Nothing Then
        sDef = oDefault.Address
    End If
    ' ユーザキャンセル操作時に馬鹿なInputBoxがBooleanを設定しようとしてSetが例外を発生する...orz
    ' どうもこの例外を回避するのは結構な手間なようなので、強行してエラー状態を確認して対処が吉っぽい
    Set sResult = Application.InputBox(Prompt:=sMessage, Title:=sTitle, Default:=sDef, Type:=8)
    bRet = True
NrmExit:
    InputRng = bRet
    Exit Function
ErrExit:
    Set sResult = Nothing
    Resume NrmExit
End Function

' ステータスバー設定
Public Sub SetStatusBar(ByVal sText As String)
    Application.StatusBar = sText
End Sub

' ステータスバー進捗
Public Sub SetStatusBarProgress(ByVal sText As String, ByVal lIdx As Long, ByVal lMax As Long)
    Application.StatusBar = sText & " " & Left(String(Int(lIdx / lMax * 10), "■") & String(10, "□"), 10) & "(" & lIdx & "/" & lMax & ")"
End Sub

' ステータスバー消去
Public Sub ClrStatusBar()
    Application.StatusBar = False
End Sub

' 右クリックメニュー追加
Sub AddRClickMenu(ByVal sTitle As String, ByVal sMacro As String)
    Dim oCmdBar As CommandBarButton
    Set oCmdBar = Application.CommandBars("Cell").Controls.Add(Temporary:=True)
    With oCmdBar
        .Caption = sTitle
        .OnAction = sMacro
    End With
End Sub

' 右クリックメニュークリア
Sub ClrRClickMenu()
    Call Application.CommandBars("Cell").Reset
End Sub

' // ブック操作 ///////////////////////////////////////////

' ブック作成
Function MakeBook( _
    ByVal bDisableEvents As Boolean, _
    ByVal bDisableAlarts As Boolean, _
    ParamArray oSheets() As Variant _
) As Workbook
    Dim bOrgEvt As Boolean: bOrgEvt = Application.EnableEvents
    Dim bOrgAlt As Boolean: bOrgAlt = Application.DisplayAlerts
On Error GoTo ErrExit
    If bDisableEvents Then Application.EnableEvents = False
    If bDisableAlarts Then Application.DisplayAlerts = False
    ' デフォで単一シートを含むブックを作成
    Dim oBook As Workbook
    Set oBook = Application.Workbooks.Add(xlWBATWorksheet)
    ' テンプレートシートをコピー
    Dim elm
    For Each elm In oSheets
        Call elm.Copy(After:=oBook.Sheets(oBook.Sheets.Count))
    Next
    If IsMissing(oSheets) = False Then
        Call oBook.Sheets(1).Delete
    End If
NrmExit:
    Application.DisplayAlerts = bOrgAlt
    Application.EnableEvents = bOrgEvt
    Set MakeBook = oBook
    Exit Function
ErrExit:
    If Not oBook Is Nothing Then
        Call oBook.Close(False)
        Set oBook = Nothing
    End If
    Resume NrmExit
End Function

' ブックオープン
Function OpenBook( _
    ByVal sPath As String, _
    ByVal bReadOnly As Boolean, _
    ByVal bDisableEvents As Boolean, _
    ByVal bDisableAlarts As Boolean, _
    ByVal bDisableUdtLnk As Boolean _
) As Workbook
    Dim oBook As Workbook
    Dim bOrgEvt As Boolean: bOrgEvt = Application.EnableEvents
    Dim bOrgAlt As Boolean: bOrgAlt = Application.DisplayAlerts
On Error GoTo ErrExit
    If bDisableEvents Then Application.EnableEvents = False
    If bDisableAlarts Then Application.DisplayAlerts = False
    ' ブックオープン
    If bDisableUdtLnk Then
        Set oBook = Application.Workbooks.Open(fileName:=sPath, ReadOnly:=bReadOnly, IgnoreReadOnlyRecommended:=True, UpdateLinks:=0)
    Else
        Set oBook = Application.Workbooks.Open(fileName:=sPath, ReadOnly:=bReadOnly, IgnoreReadOnlyRecommended:=True)
    End If
NrmExit:
    Application.DisplayAlerts = bOrgAlt
    Application.EnableEvents = bOrgEvt
    Set OpenBook = oBook
    Exit Function
ErrExit:
    If Not oBook Is Nothing Then
        Call oBook.Close(False)
        Set oBook = Nothing
    End If
    Resume NrmExit
End Function

' テンプレートブックオープン
' ・テンプレート形式ではない任意のファイルをテンプレートとしてメモリ上に開く
Function OpenTemplateBook( _
    ByVal sPath As String, _
    ByVal bDisableEvents As Boolean, _
    ByVal bDisableAlarts As Boolean _
) As Workbook
    Dim oBook As Workbook
    Dim bOrgEvt As Boolean: bOrgEvt = Application.EnableEvents
    Dim bOrgAlt As Boolean: bOrgAlt = Application.DisplayAlerts
On Error GoTo ErrExit
    If bDisableEvents Then Application.EnableEvents = False
    If bDisableAlarts Then Application.DisplayAlerts = False
    ' テンプレートブックオープン
    Set oBook = Application.Workbooks.Add(sPath)
NrmExit:
    Application.DisplayAlerts = bOrgAlt
    Application.EnableEvents = bOrgEvt
    Set OpenTemplateBook = oBook
    Exit Function
ErrExit:
    If Not oBook Is Nothing Then
        Call oBook.Close(False)
        Set oBook = Nothing
    End If
    Resume NrmExit
End Function

' ブック検索
Function SearchBook(ByVal sBook As String) As Workbook
    Dim elm As Workbook
    For Each elm In Application.Workbooks
        If elm.Name Like sBook Then
            Exit For
        End If
    Next
    Set SearchBook = elm
End Function

' ブック名称
Function BookName(ByRef oBook As Workbook, Optional ByVal bExtension As Boolean = True) As String
    If bExtension = True Then
        BookName = oBook.Name
    Else
        BookName = CreateFSO().GetBaseName(oBook.Name)
    End If
End Function

' ブックの表示状態を設定
Sub SetBookVisibleState(ByRef oBook As Workbook, ByVal bVisible As Boolean)
    Application.Windows(oBook.Name).Visible = bVisible
End Sub

' ブックの最終版状態を設定
Sub SetBookFinalState(ByRef oBook As Workbook, Optional ByVal bMode As Boolean = True)
    Dim bOrgAlt As Boolean: bOrgAlt = Application.DisplayAlerts
On Error GoTo ErrExit
    Application.DisplayAlerts = False
    oBook.Final = bMode
NrmExit:
    Application.DisplayAlerts = bOrgAlt
    Exit Sub
ErrExit:
    Resume NrmExit
End Sub

' ブック属性設定
' ・Title           : タイトル
' ・Subject         : サブタイトル
' ・Company         : 会社名
' ・Author          : 作成者
' ・Last Author     : 更新者
' ・Keywords        : キーワード
' ・Comments        : コメント
' ・Revision Number : 改訂番号
' ・Security        : セキュリティ
' ・Hyperlink Base  : ハイパーリンクの基点
Function SetBookProp(ByRef oBook As Workbook, ByVal sProp As String, ByVal sText As String)
    oBook.BuiltinDocumentProperties(sProp).Value = sText
End Function

' ブック属性クリア
Function ClrBookProp(ByRef oBook As Workbook) As Boolean
On Error Resume Next
    Dim oProp As DocumentProperty
    For Each oProp In oBook.BuiltinDocumentProperties
        If oProp.Name <> "Hyperlink base" And _
           oProp.Name <> "Creation date" And _
           oProp.Name <> "Last save time" And _
           oProp.Name <> "Last print date" _
        Then
            oProp.Value = CStr(Empty)
        End If
    Next
End Function

' // シート操作 ///////////////////////////////////////////
' // ・シートにはワークシートとチャートシートの二種類がある

' シート確認
Function IsExistSheet(ByVal oBook As Workbook, ByVal sNamePtrn As String) As Boolean
    Dim bRet As Boolean
    Dim elm
    For Each elm In oBook.Sheets
        If elm.Name Like sNamePtrn Then
            bRet = True
            Exit For
        End If
    Next
    IsExistSheet = bRet
End Function

' ワークシート検索
Function SearchWorkSheet(ByVal oBook As Workbook, ByVal sName As String) As Worksheet
    Dim elm As Worksheet
    For Each elm In oBook.Worksheets
        If elm.Name = sName Then
            Exit For
        End If
    Next
    Set SearchWorkSheet = elm
End Function

' チャートシート検索
Function SearchChartSheet(ByVal oBook As Workbook, ByVal sName As String) As Chart
    Dim elm As Chart
    For Each elm In oBook.Charts
        If elm.Name = sName Then
            Exit For
        End If
    Next
    Set SearchChartSheet = elm
End Function

' ワークシート追加
Function AddWorkSheet(ByRef oBook As Workbook, ByVal sName As String) As Worksheet
    Dim oRet As Worksheet
    Set oRet = oBook.Sheets.Add(After:=oBook.Sheets(oBook.Sheets.Count))
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' 重複しない名称を自動設定
    Set AddWorkSheet = oRet
End Function

' チャートシート追加
Function AddChartSheet(ByRef oBook As Workbook, ByVal sName As String) As Chart
    Dim oRet As Chart
    Set oRet = oBook.Sheets.Add(After:=oBook.Sheets(oBook.Sheets.Count))
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' 重複しない名称を自動設定
    Set AddChartSheet = oRet
End Function

' ワークシート複写
Function CopyWorkSheet(ByRef oBook As Workbook, ByRef oSheet As Worksheet, ByVal sName As String) As Worksheet
    Dim oRet As Worksheet
    Call oSheet.Copy(After:=oBook.Sheets(oBook.Sheets.Count))   ' 戻り値が無いのでブック末尾に固定配置
    Set oRet = oBook.Sheets(oBook.Sheets.Count)                 ' 戻り値が無いのでブック末尾に固定配置
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' 重複しない名称を自動設定
    oRet.Visible = xlSheetVisible                               ' コピー後に不可視では困るので可視強制
    Set CopyWorkSheet = oRet
End Function

' チャートシート複写
Function CopyChartSheet(ByRef oBook As Workbook, ByRef oSheet As Chart, ByVal sName As String) As Chart
    Dim oRet As Chart
    Call oSheet.Copy(After:=oBook.Sheets(oBook.Sheets.Count))   ' 戻り値が無いのでブック末尾に固定配置
    Set oRet = oBook.Sheets(oBook.Sheets.Count)                 ' 戻り値が無いのでブック末尾に固定配置
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' 重複しない名称を自動設定
    oRet.Visible = xlSheetVisible                               ' コピー後に不可視では困るので可視強制
    Set CopyChartSheet = oRet
End Function

' ブック内でユニークなシート名を作成する
Private Function GenUniqSheetName(ByVal oBook As Workbook, ByVal sName As String) As String
    Dim txt As String
    Dim idx As Long
    txt = sName
    idx = 2
    Do
        Dim fnd As Boolean: fnd = False
        Dim elm As Variant
        For Each elm In oBook.Sheets
            If elm.Name = txt Then
                fnd = True
                Exit For
            End If
        Next
        If fnd Then
            txt = sName & "(" & idx & ")"
            idx = idx + 1
        End If
    Loop While fnd = True
    GenUniqSheetName = txt
End Function

' シート削除
Function DelSheet(ByRef oBook As Workbook, ByVal oSheet As Worksheet, Optional ByVal bDisableAlarts As Boolean = False) As Boolean
    Dim bRet As Boolean
    Dim bOrgAlt As Boolean: bOrgAlt = Application.DisplayAlerts
On Error GoTo ErrExit:
    If bDisableAlarts Then
        Application.DisplayAlerts = False
    End If
    oSheet.Delete
    bRet = True
NrmExit:
    If bDisableAlarts Then
        Application.DisplayAlerts = bOrgAlt
    End If
    DelSheet = bRet
    Exit Function
ErrExit:
    Resume NrmExit
End Function

' シート移動
Sub MoveSheet(ByRef oBook As Workbook, ByVal oSheet As Worksheet, Optional ByVal bTop As Boolean = True)
    If bTop Then
        Call oSheet.Move(Before:=oBook.Sheets(1))
    Else
        Call oSheet.Move(After:=oBook.Sheets(oBook.Sheets.Count))
    End If
End Sub

' シート並べ替え
Sub SortSheet(ByRef oBook As Workbook, ByVal bReverse As Boolean)
    Dim li As Long
    Dim lj As Long
    Dim aBuff() As String
    Dim sSwap As String
    ReDim aBuff(oBook.Sheets.Count)
    For li = 1 To oBook.Sheets.Count
        aBuff(li) = oBook.Sheets(li).Name
    Next li
    For li = 1 To oBook.Sheets.Count
        For lj = oBook.Sheets.Count To li Step -1
            If bReverse = False Then
                If aBuff(li) > aBuff(lj) Then
                    sSwap = aBuff(li)
                    aBuff(li) = aBuff(lj)
                    aBuff(lj) = sSwap
                End If
            Else
                If aBuff(li) < aBuff(lj) Then
                    sSwap = aBuff(li)
                    aBuff(li) = aBuff(lj)
                    aBuff(lj) = sSwap
                End If
            End If
        Next
    Next
    oBook.Sheets(aBuff(1)).Move Before:=oBook.Sheets(1)
    For li = 2 To oBook.Sheets.Count
        oBook.Sheets(aBuff(li)).Move After:=oBook.Sheets(li - 1)
    Next
End Sub

' ブック内の指定シートのみを表示する
Sub SetVisibleSheet(ByRef oBook As Workbook, ParamArray sSheetPtrns())
    ' 非表示→表示
    Dim elm As Worksheet
    Dim sSheetPtrn As Variant
    For Each elm In oBook.Sheets
        For Each sSheetPtrn In sSheetPtrns
            If elm.Name Like sSheetPtrn Then
                elm.Visible = xlSheetVisible
            End If
        Next
    Next
    ' 表示→非表示
    Dim eVisible As XlSheetVisibility
    For Each elm In oBook.Sheets
        eVisible = xlSheetHidden
        For Each sSheetPtrn In sSheetPtrns
            If elm.Name Like sSheetPtrn Then
                eVisible = xlSheetVisible
            End If
        Next
        elm.Visible = eVisible
    Next
End Sub

' // 範囲操作 /////////////////////////////////////////////

' 文字列をRangeに変換
Function CRng(sRange As String) As Range
    CRng = Application.Range(sRange)
End Function

' 列番号→A1参照形式の列名部分(例:10列目→J列)
Function ColumnIdx2Name(ByVal lColIdx As Long) As String
    Dim sRet As String
    Dim lVal As Long
    lVal = lColIdx
    Do While lVal > 0
        sRet = Chr(65 + (lVal - 1) Mod 26) & sRet
        lVal = (lVal - 1) \ 26
    Loop
    ColumnIdx2Name = sRet
End Function

' A1参照形式の列名部分→列番号(例:J列→10列目)
Function ColumnName2Idx(ByVal sColName As String) As Long
    Dim lRet As Long
    Dim sText As String
    Dim lChar As Long
    sText = Trim(UCase(sColName))
    If sText = "" Then
        lRet = -1
        GoTo NrmExit
    End If
    Dim idx As Long
    For idx = 1 To Len(sText)
        lChar = Asc(Mid(sText, idx, 1)) - 64
        If lChar < 1 Or lChar > 26 Then
            lRet = -1
            GoTo NrmExit
        End If
        lRet = (lRet * 26) + lChar
    Next
NrmExit:
    ColumnName2Idx = lRet
End Function

' 簡易開始行番号取得
Function MinRow(ByVal oSheet As Worksheet, lCol As Long) As Long
    If oSheet.Cells(1, lCol).Value = "" Then
        MinRow = oSheet.Cells(1, lCol).End(xlDown).Row
    Else
        MinRow = 1
    End If
End Function

' 簡易最終行番号取得
Function MaxRow(ByVal oSheet As Worksheet, lCol As Long) As Long
    If oSheet.Cells(oSheet.Rows.Count, lCol).Value = "" Then
        MaxRow = oSheet.Cells(oSheet.Rows.Count, lCol).End(xlUp).Row
    Else
        MaxRow = oSheet.Rows.Count
    End If
End Function

' 簡易開始列番号取得
Function MinCol(ByVal oSheet As Worksheet, lRow As Long) As Long
    If oSheet.Cells(lRow, 1).Value = "" Then
        MinCol = oSheet.Cells(lRow, 1).End(xlToRight).Column
    Else
        MinCol = 1
    End If
End Function

' 簡易最終列番号取得
Function MaxCol(ByVal oSheet As Worksheet, lRow As Long) As Long
    If oSheet.Cells(lRow, oSheet.Columns.Count).Value = "" Then
        MaxCol = oSheet.Cells(lRow, oSheet.Columns.Count).End(xlToLeft).Column
    Else
        MaxCol = oSheet.Columns.Count
    End If
End Function

' 範囲拡張
Function ExtendRange(ByVal oRange As Range, ByVal lTop As Long, ByVal lLeft As Long, ByVal lBottom As Long, ByVal lRight) As Range
    Set ExtendRange = oRange.offset(-lTop, -lLeft).Resize(oRange.Rows.Count + lTop + lBottom, oRange.Columns.Count + lLeft + lRight)
End Function

' セル結合を考慮した行拡張
Function EntireRowEx(ByVal oSheet As Worksheet, ByVal oRange As Range) As Range
    Dim oResult As Range
    Set oResult = oRange.EntireRow
    If Not Intersect(oResult, oSheet.UsedRange) Is Nothing Then
        Do
            Dim sPrev As String
            Dim sCrnt As String
            sPrev = oResult.Address
            Dim elm As Range
            For Each elm In Intersect(oResult.EntireRow, oSheet.UsedRange)
                Set oResult = Union(oResult, elm.MergeArea.EntireRow)
            Next
            sCrnt = oResult.Address
        Loop While sPrev <> sCrnt
    End If
    Set EntireRowEx = oResult
End Function

' セル結合を考慮した列拡張
Function EntireColumnEx(ByVal oSheet As Worksheet, ByVal oRange As Range) As Range
    Dim oResult As Range
    Set oResult = oRange.EntireColumn
    If Not Intersect(oResult, oSheet.UsedRange) Is Nothing Then
        Do
            Dim sPrev As String
            Dim sCrnt As String
            sPrev = oResult.Address
            Dim elm As Range
            For Each elm In Intersect(oResult.EntireColumn, oSheet.UsedRange)
                Set oResult = Union(oResult, elm.MergeArea.EntireColumn)
            Next
            sCrnt = oResult.Address
        Loop While sPrev <> sCrnt
    End If
    Set EntireColumnEx = oResult
End Function

' 範囲補集合
Function NotRange(ByVal oRange As Range, Optional ByVal oSheet As Worksheet = Nothing) As Range
    Dim oResult As Range
    If Not oRange Is Nothing Then
        Set oResult = oRange.Worksheet.Cells
    Else
        Set oResult = oSheet.Cells ' oRangeもoSheetもNothingなら死ぬのでTPOに合わせてどうぞ。
    End If
    ' 全体集合と個々の判定領域の積を取る
    Dim rng As Range
    For Each rng In oRange.Areas
        Set oResult = IntersectRange(oResult, coNotRange(rng))
    Next
    Set NotRange = oResult
End Function
Private Function coNotRange(ByVal oRange As Range) As Range
    Dim oResult As Range
    Set oResult = Nothing
    Dim oSheet As Worksheet
    Set oSheet = oRange.Worksheet
    Dim idx As Long
    Dim rng As Range
    ' 指定範囲の上側(■の部分)
    '  ■■■
    '  □×□
    '  □□□
    idx = oRange.Item(1).Row - 1
    If idx > 0 Then
        Set oResult = UnionRange(oResult, oSheet.Range(oSheet.Rows(1), oSheet.Rows(idx)))
    End If
    '指定範囲の下側(■の部分)
    '  □□□
    '  □×□
    '  ■■■
    idx = oRange.Item(oRange.Rows.Count, oRange.Columns.Count).Row + 1
    If idx < oSheet.Rows.Count Then
        Set oResult = UnionRange(oResult, oSheet.Range(oSheet.Rows(idx), oSheet.Rows(oSheet.Rows.Count)))
    End If
    '指定範囲の左側(■の部分)
    '  □□□
    '  ■×□
    '  □□□
    idx = oRange.Item(1).Column - 1
    If idx > 0 Then
        Set rng = Intersect(oSheet.Range(oSheet.Columns(1), oSheet.Columns(idx)), oRange.EntireRow)
        Set oResult = UnionRange(oResult, rng)
    End If
    '指定範囲の右側(■の部分)
    '  □□□
    '  □×■
    '  □□□
    idx = oRange.Item(oRange.Rows.Count, oRange.Columns.Count).Column + 1
    If idx < oSheet.Columns.Count Then
        Set rng = Intersect(oSheet.Range(oSheet.Columns(idx), oSheet.Columns(oSheet.Columns.Count)), oRange.EntireRow)
        Set oResult = UnionRange(oResult, rng)
    End If
    Set coNotRange = oResult
End Function

' 範囲和集合
Function UnionRange(ParamArray oRanges() As Variant) As Range
    Dim oResult As Range
    Set oResult = Nothing
    If UBound(oRanges) - LBound(oRanges) + 1 > 0 Then
        Set oResult = oRanges(0)
        Dim rng As Variant
        For Each rng In oRanges
            If oResult Is Nothing Then
                If rng Is Nothing Then
                    Set oResult = Nothing
                Else
                    Set oResult = rng
                End If
            Else
                If rng Is Nothing Then
                    Set oResult = oResult
                Else
                    Set oResult = Union(oResult, rng)
                End If
            End If
        Next
    End If
    Set UnionRange = oResult
End Function

' 範囲積集合
Function IntersectRange(ParamArray oRanges() As Variant) As Range
    Dim oResult As Range
    Set oResult = Nothing
    If UBound(oRanges) - LBound(oRanges) + 1 > 0 Then
        Set oResult = oRanges(0)
        Dim rng As Variant
        For Each rng In oRanges
            If oResult Is Nothing Then
                If rng Is Nothing Then
                    Set oResult = Nothing
                Else
                    Set oResult = Nothing
                End If
            Else
                If rng Is Nothing Then
                    Set oResult = Nothing
                Else
                    Set oResult = Intersect(oResult, rng)
                End If
            End If
        Next
    End If
    Set IntersectRange = oResult
End Function

' 範囲差集合
Function ExceptRange(ByVal oLHS As Range, ByVal oRHS As Range) As Range
    Dim oResult As Range
    If oLHS Is Nothing Then
        Set oResult = Nothing
        GoTo NrmExit
    End If
    If oRHS Is Nothing Then
        Set oResult = oLHS
        GoTo NrmExit
    End If
    ' 左手領域から右手領域を削り落とす
    Dim lhs As Range: Set lhs = Nothing
    For Each lhs In oLHS.Areas
        Dim tmp As Range: Set tmp = lhs
        Dim rhs As Range: Set rhs = Nothing
        For Each rhs In oRHS.Areas
            Set tmp = IntersectRange(tmp, Intersect(lhs, coNotRange(rhs)))
        Next
        Set oResult = UnionRange(oResult, tmp)
    Next
NrmExit:
    Set ExceptRange = oResult
End Function

' 範囲に１次元配列を一括出力
' ・配列サイズから出力範囲を自動決定
Sub WriteRange1D(ByRef oRange As Range, ByVal vArray As Variant)
    oRange.Resize(UBound(vArray) - LBound(vArray) + 1).Value = vArray
End Sub

' 範囲に２次元配列を一括出力
' ・配列サイズから出力範囲を自動決定
Sub WriteRange2D(ByRef oRange As Range, ByVal vArray As Variant)
    oRange.Resize(UBound(vArray, 1) - LBound(vArray, 1) + 1, UBound(vArray, 2) - LBound(vArray, 2) + 1).Value = vArray
End Sub

' 範囲検索
Public Function FindRange( _
    ByVal oRange As Range, _
    ByVal sText As String, _
    Optional ByVal LookIn As XlFindLookIn = xlValues, _
    Optional ByVal LookAt As XlLookAt = xlPart, _
    Optional ByVal MatchCase As Boolean = False, _
    Optional ByVal MatchByte As Boolean = False, _
    Optional ByVal SearchFormat As Boolean = False _
) As Range
    ' Excelのバグ？で結合セルを検出できなくなるため
    ' SearchOrder:=xlByColumnsは指定できなくしている
    Dim oSttRng As Range
    Dim oCrtRng As Range
    Dim oRetRng As Range
    Set oSttRng = oRange.Find(What:=sText, After:=oRange(1), _
                              LookIn:=LookIn, LookAt:=LookAt, _
                              SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                              MatchCase:=MatchCase, MatchByte:=MatchByte, SearchFormat:=SearchFormat)
    Set oCrtRng = oSttRng
    Set oRetRng = oSttRng
    If Not oCrtRng Is Nothing Then
        Do
            Set oRetRng = Union(oRetRng, oCrtRng)
            Set oCrtRng = oRange.FindNext(oCrtRng)
            If oCrtRng Is Nothing Then
                Exit Do
            End If
        Loop While oSttRng.Address <> oCrtRng.Address
    End If
    Set FindRange = oRetRng
End Function

'' 簡易検索
'Function SearchRange(oRange As Range, ByVal sText As String) As Range
'    ' ※FindとかMatchは日付絡みの扱いが妙に難しいので単純検索
'    Dim oRet As Range
'    Dim elm As Range
'    For Each elm In oRange.Cells
'        If elm.Text Like sText Then
'            Set oRet = elm
'            Exit For
'        End If
'    Next
'    Set SearchRange = oRet
'End Function

'' 簡易検索
'Function SearchRanges(ByVal oRange As Range, ByVal sText As String) As Collection
'    ' ※FindとかMatchは日付絡みの扱いが妙に難しいので単純検索
'    Dim oRet As New Collection
'    Dim elm As Range
'    For Each elm In oRange.Cells
'        If elm.Text Like sText Then
'            Call oRet.Add(elm)
'        End If
'    Next
'    Set SearchRanges = oRet
'End Function

' // 方眼Excel用 //////////////////////////////////////////

' １セル１文字形式読出
' ・１セルに１文字以上の文字が入力してある場合は例外を吐きます
' ・未入力の空セルは空白文字が入力してあるものと見做します
Function ReadCells(ByVal oRange As Range, Optional ByVal bLTrim As Boolean = False, Optional ByVal bRTrim As Boolean = True, Optional bCheck As Boolean = True) As String
    Dim sChar As String
    Dim sBuff As String
    Dim elm As Range
    For Each elm In oRange.Cells
        Select Case elm.Text
        Case ""
            sChar = " "
        Case "''"
            sChar = "'"
        Case "'="
            sChar = "="
        Case Else
            sChar = elm.Text
        End Select
        If bCheck Then
            If Len(sChar) > 1 Then
                Call Err.Raise(9999, "", "１セル１文字としてください:" & elm.Address)
            End If
        End If
        sBuff = sBuff & sChar
    Next
    If bLTrim Then
        sBuff = LTrim(sBuff)
    End If
    If bRTrim Then
        sBuff = RTrim(sBuff)
    End If
    ReadCells = sBuff
End Function

' １セル１文字形式書出
' ・書込範囲は始点セル１点＋長さによって指定します
' ・折り返しは複数回の呼び出しで対処してください
Sub WriteCells(ByVal oDst As Range, ByVal sText As String, ByVal lLen As Long, Optional ByVal lAlign As Long = 0)
    ' アライメント
    Dim sBuff As String
    Select Case lAlign
        Case 0: sBuff = Left(sText & String(lLen, " "), lLen)   ' 左詰め
        Case 1: sBuff = Right(String(lLen, " ") & sText, lLen)  ' 右詰め
    End Select
    ' 一括転送
    Dim lIdx As Integer
    Dim aBuff() As String
    Dim sChar As String
    ReDim aBuff(lLen - 1) As String
    For lIdx = 1 To Len(sBuff)
        sChar = Mid(sBuff, lIdx, 1)
        Select Case sChar
        Case "="
            aBuff(lIdx - 1) = "'="
        Case "'"
            aBuff(lIdx - 1) = "''"
        Case Else
            aBuff(lIdx - 1) = Mid(sBuff, lIdx, 1)
        End Select
    Next
    oDst.Resize(1, UBound(aBuff) + 1).Value = aBuff
End Sub

' １セル１文字形式消去
' ・消去範囲は始点セル１点＋長さによって指定します
' ・折り返しは複数回の呼び出しで対処してください
Sub ClearCells(ByRef oDst As Range, ByVal lLen As Long)
    oDst.Resize(1, lLen).ClearContents
End Sub

' １セル１文字形式複写
' ・転送元/転送範囲は始点セル１点＋長さによって指定します
' ・折り返しは複数回の呼び出しで対処してください
Sub CopyCells(ByRef oDst As Range, ByVal oSrc As Range, ByVal lLen As Long)
    oDst.Resize(1, lLen).Value = oSrc.Resize(1, lLen).Value
End Sub

' // 便利系 ///////////////////////////////////////////////
' // ・あんまし汎用じゃないけど、どうせコピペするしなぁ...という代物

' シート装飾
Sub DecolateSheet( _
    ByRef oSheet As Worksheet, _
    ByRef oRange As Range, _
    Optional ByVal bHead As Boolean = True, _
    Optional ByVal bFilter As Boolean = True, _
    Optional ByVal lSort As Long = 0 _
)
    ' 罫線
    oRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    oRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    oRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    oRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    oRange.Borders(xlInsideHorizontal).LineStyle = xlDash
    oRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    If bHead Then
        oRange.Rows(1).Borders(xlEdgeTop).LineStyle = xlContinuous
        oRange.Rows(1).Borders(xlEdgeBottom).LineStyle = xlDouble
        oRange.Rows(1).Interior.ThemeColor = xlThemeColorLight2
        oRange.Rows(1).Font.ThemeColor = xlThemeColorDark2
    End If
    
    ' 行列巾自動調整
    oRange.Rows.AutoFit
    oRange.Columns.AutoFit
    
    ' オートフィルタ設定
    If bHead And bFilter And oSheet.AutoFilterMode = False Then
        oRange.Columns.AutoFilter
    End If
    
    ' ソート
    ' ・ソートインデックスを0ならソートしない
    ' ・ソートインデックスを0以上なら昇順ソート
    ' ・ソートインデックスを0以下なら降順ソート
    If lSort <> 0 Then
        oRange.Columns.Sort Header:=IIf(bHead, xlYes, xlNo), Key1:=oRange.Columns(Abs(lSort)), Order1:=IIf(lSort > 0, xlAscending, xlDescending)
    End If
    
    ' コメント位置自動調整
    ' ・ソートや行列巾自動調整でズレるので、半ば必須
    Call AutoFitComment(oSheet)
End Sub

' 指定行以降をクリア
' ・最初の行は入力規則等を残すため削除しない
' ・それ以降はUsedRangeを削減するために削除する
Sub ClearRows(ByRef oSheet As Worksheet, ByVal lRow As Long)
    Dim oRng1st As Range
    Dim oRngOth As Range
    Set oRng1st = oSheet.Cells(lRow + 0, 1).EntireRow
    Set oRngOth = oSheet.Range(oSheet.Cells(lRow + 1, 1), oSheet.Cells(oSheet.Rows.Count, 1)).EntireRow
    Call oRng1st.ClearContents
    Call oRngOth.Delete
End Sub

' 指定列以降をクリア
' ・最初の列は入力規則等を残すため削除しない
' ・それ以降はUsedRangeを削減するために削除する
Sub ClearCols(ByRef oSheet As Worksheet, ByVal lCol As Long)
    Dim oRng1st As Range
    Dim oRngOth As Range
    Set oRng1st = oSheet.Cells(1, lCol + 0).EntireColumn
    Set oRngOth = oSheet.Range(oSheet.Cells(1, lCol + 1), oSheet.Cells(1, oSheet.Columns.Count)).EntireColumn
    Call oRng1st.ClearContents
    Call oRngOth.Delete
End Sub

' // 数値操作 /////////////////////////////////////////////

' 最小値
Function Min(ParamArray oVals() As Variant) As Variant
    Min = WorksheetFunction.Min(oVals)
End Function

' 最大値
Function Max(ParamArray oVals() As Variant) As Variant
    Max = WorksheetFunction.Max(oVals)
End Function

' 平均値
Function Ave(ParamArray oVals() As Variant) As Variant
    Ave = WorksheetFunction.Average(oVals)
End Function

' 標準偏差
Function StDev(ParamArray oVals() As Variant) As Variant
    StDev = WorksheetFunction.StDev(oVals)
End Function

' 四捨五入
Function Round(ByVal vVal As Variant, ByVal lDigit As Long) As Variant
    Round = WorksheetFunction.Round(vVal, lDigit) ' VBAな素のRoundメソッドは銀行型丸め。一般的な四捨五入はこっち。
End Function

' 切り捨て
Function RoundDown(ByVal vVal As Variant, ByVal lDigit As Long) As Variant
    RoundDown = WorksheetFunction.RoundDown(vVal, lDigit)
End Function

' 切り上げ
Function RoundUp(ByVal vVal As Variant, ByVal lDigit As Long) As Variant
    RoundUp = WorksheetFunction.RoundUp(vVal, lDigit)
End Function

' 数値範囲制限
Function RestrictNum(ByVal vVal As Variant, Optional ByVal vMin As Variant, Optional ByVal vMax As Variant) As Variant
    If Not IsMissing(vMin) Then vVal = IIf(vVal > vMin, vVal, vMin)
    If Not IsMissing(vMax) Then vVal = IIf(vVal < vMax, vVal, vMax)
    RestrictNum = vVal
End Function

' // テキスト操作 /////////////////////////////////////////

' シングルクォート
Function SQuote(ByVal sText As String)
    SQuote = "'" & sText & "'"
End Function

' ダブルクォート
Function DQuote(ByVal sText As String)
    DQuote = """" & sText & """"
End Function

' 文字埋込
Function EmbText(ByVal sText As String, ParamArray vParam() As Variant) As String
    Dim idx As Long
    Dim elm As Variant
    For Each elm In vParam
        sText = Replace(sText, "{" & idx & "}", elm)
        idx = idx + 1
    Next
    EmbText = sText
End Function

' NULL or 空文字判定
Function IsNullOrEmpty(ByVal sText As Variant) As Boolean
    If IsEmpty(sText) = True Then
        IsNullOrEmpty = True
    ElseIf IsNull(sText) = True Then
        IsNullOrEmpty = True
    ElseIf sText = "" Then
        IsNullOrEmpty = True
    Else
        IsNullOrEmpty = False
    End If
End Function

' 文字結合(配列)
Function ConcatArray(ByVal sSep As String, sText() As Variant) As String
    ConcatArray = Join(sText, sSep)
End Function

' 文字結合(可変長配列)
Function ConcatArgs(ByVal sSep As String, ParamArray sText() As Variant) As String
    ConcatArgs = Join(sText, sSep)
End Function

' 文字結合(コレクション)
Function ConcatCollection(ByVal sSep As String, ByVal sText As Collection) As String
    Dim sRet As String
    Dim sChr As Variant
    For Each sChr In sText
        sRet = sRet & sChr & sSep
    Next
    If Right(sRet, 1) = sSep Then
        sRet = Left(sRet, Len(sRet) - 1)
    End If
    ConcatCollection = sRet
End Function

' 空文字以外を結合(配列)
Function PackArray(ByVal sSep As String, ByVal sText As Variant) As String
    Dim sRet As String
    Dim sChr As Variant
    For Each sChr In sText
        sRet = sRet & IIf(IsNullOrEmpty(sChr) = False, sChr & sSep, "")
    Next
    If Right(sRet, 1) = sSep Then
        sRet = Left(sRet, Len(sRet) - 1)
    End If
    PackArray = sRet
End Function

' 空文字以外を結合(可変長配列)
Function PackArgs(ByVal sSep As String, ParamArray sText() As Variant) As String
    Dim sRet As String
    Dim sChr As Variant
    For Each sChr In sText
        sRet = sRet & IIf(IsNullOrEmpty(sChr) = False, sChr & sSep, "")
    Next
    If Right(sRet, 1) = sSep Then
        sRet = Left(sRet, Len(sRet) - 1)
    End If
    PackArgs = sRet
End Function

' 空文字以外を結合(コレクション)
Function PackCollection(ByVal sSep As String, ByVal sText As Collection) As String
    Dim sRet As String
    Dim sChr As Variant
    For Each sChr In sText
        sRet = sRet & IIf(IsNullOrEmpty(sChr) = False, sChr & sSep, "")
    Next
    If Right(sRet, 1) = sSep Then
        sRet = Left(sRet, Len(sRet) - 1)
    End If
    PackCollection = sRet
End Function

' 左端指定文字消去
Function LTrimEx(ByVal sText As String, ByVal sChar As String) As String
    ' NOTE:TrimEx(sText, csELPtrn) ←こうすれば先頭末尾のどっちかよくわからんCRLFを全部削除できる
    While Left(sText, 1) Like sChar
        sText = Right(sText, Len(sText) - 1)
    Wend
    LTrimEx = sText
End Function

' 右端指定文字消去
Function RTrimEx(ByVal sText As String, ByVal sChar As String) As String
    ' NOTE:TrimEx(sText, csELPtrn) ←こうすれば先頭末尾のどっちかよくわからんCRLFを全部削除できる
    While Right(sText, 1) Like sChar
        sText = Left(sText, Len(sText) - 1)
    Wend
    RTrimEx = sText
End Function

' 両端指定文字消去
Function TrimEx(ByVal sText As String, ByVal sChar As String) As String
    ' NOTE:TrimEx(sText, csELPtrn) ←こうすれば先頭末尾のどっちかよくわからんCRLFを全部削除できる
    sText = LTrimEx(sText, sChar)
    sText = RTrimEx(sText, sChar)
    TrimEx = sText
End Function

' YYYYMMDD/YYMMDD形式の文字を日付に変換
Function YYYYMMDD2Date(ByVal sText As String, Optional ByVal sFormat As String = "####/##/##") As Date
    YYYYMMDD2Date = CDate(Format(sText, sFormat))
End Function

' 和式西暦フォーマット
Function FormatDate(ByVal oDate As Date) As String
    FormatDate = Format(oDate, "YYYY/MM/DD")
End Function

' 和式時刻フォーマット
Function FormatTime(ByVal oDate As Date) As String
    FormatTime = Format(oDate, "hh:mm")
End Function

' 西暦4桁+月2桁+日2桁
Function YYYYMMDD(ByVal oDate As Date) As String
    YYYYMMDD = Format(oDate, "YYYYMMDD")
End Function

' 西暦4桁+月2桁
Function YYYYMM(ByVal oDate As Date) As String
    YYYYMM = Format(oDate, "YYYYMM")
End Function

' 月2桁+日2桁
Function MMDD(ByVal oDate As Date) As String
    MMDD = Format(oDate, "MMDD")
End Function

' 西暦4桁
Function YYYY(ByVal oDate As Date) As String
    YYYY = Format(oDate, "YYYY")
End Function

' 西暦2桁
Function YY(ByVal oDate As Date) As String
    YY = Format(oDate, "YY")
End Function

' 月2桁
Function MM(ByVal oDate As Date) As String
    MM = Format(oDate, "MM")
End Function

' 日2桁
Function DD(ByVal oDate As Date) As String
    DD = Format(oDate, "DD")
End Function

' 年度
Function FinancialYear(ByVal oDate As Date) As String
    FinancialYear = Format(DateAdd("m", -3, oDate), "YYYY")
End Function

' 半期(ピリオド)
Function FinancialPeriod(ByVal oDate As Date, Optional ByVal s1H As String = "T1", Optional ByVal s2H As String = "T2") As String
    FinancialPeriod = IIf(Month(oDate) >= 4 And Month(oDate) <= 9, s1H, s2H)
End Function

' 四半期(クォータ)
Function FinancialQuarter(ByVal oDate As Date, Optional ByVal s1Q As String = "Q1", Optional ByVal s2Q As String = "Q2", Optional ByVal s3Q As String = "Q3", Optional ByVal s4Q As String = "Q4") As String
    Dim sRet As String
    Select Case Format(DateAdd("m", -3, oDate), "Q")
        Case 1: sRet = s1Q
        Case 2: sRet = s2Q
        Case 3: sRet = s3Q
        Case 4: sRet = s4Q
    End Select
    FinancialQuarter = sRet
End Function

' 期間内判定
Function IsDuring(oTime As Date, oSTime As Date, oETime As Date)
    IsDuring = (oSTime <= oTime And oTime <= oETime)
End Function

' NULLを読替える
Public Function NZ(ByVal Value As Variant, Optional ByVal IsNullValue As Variant = Empty) As Variant
    If IsNull(Value) Then
        NZ = IsNullValue
    Else
        NZ = Value
    End If
End Function

' // 配列操作 /////////////////////////////////////////////

' 指定値で埋められた任意長の配列を生成
Function Arrays(ByVal lNum As Long, ByVal vVal As Variant) As Variant
    Dim vRet As Variant
    ReDim vRet(lNum - 1) As Variant
    Dim idx As Long
    For idx = LBound(vRet) To UBound(vRet)
        vRet(idx) = vVal
    Next
    Arrays = vRet
End Function

' ２次元配列から１次元配列をスライス(行方向)
Function SliceArrayRow(ByVal vVal As Variant, ByVal lRow As Long) As Variant
    Dim vRet() As Variant
    ReDim vRet(LBound(vVal, 2) To UBound(vVal, 2))
    Dim idx As Long
    For idx = LBound(vVal, 2) To UBound(vVal, 2)
        vRet(idx) = vVal(lRow, idx)
    Next
    SliceArrayRow = vRet
End Function

' ２次元配列から１次元配列をスライス(列方向)
Function SliceArrayCol(ByVal vVal As Variant, ByVal lCol As Long) As Variant
    Dim vRet() As Variant
    ReDim vRet(LBound(vVal, 1) To UBound(vVal, 1))
    Dim idx As Long
    For idx = LBound(vVal, 1) To UBound(vVal, 1)
        vRet(idx) = vVal(idx, lCol)
    Next
    SliceArrayCol = vRet
End Function

' // ハイパーリンク操作 ///////////////////////////////////

' 汎用リンク設定
Sub SetHLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sAddress As String, Optional ByVal sSubAddress As String = "")
    If sAddress <> "" Or sSubAddress <> "" Then
        If sTextToDisplay <> "" Then
            Call oRange.Hyperlinks.Add(Anchor:=oRange, Address:=sAddress, SubAddress:=sSubAddress, TextToDisplay:=sTextToDisplay)
        ElseIf oRange.Text <> "" Then
            Call oRange.Hyperlinks.Add(Anchor:=oRange, Address:=sAddress, SubAddress:=sSubAddress, TextToDisplay:=oRange.Text)
        Else
            Call oRange.Hyperlinks.Add(Anchor:=oRange, Address:=sAddress, SubAddress:=sSubAddress, TextToDisplay:="x")
        End If
    Else
        Call oRange.Hyperlinks.Delete
    End If
End Sub

' セル参照リンク設定
Sub SetCellLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal oRef As Range)
    If sTextToDisplay = "" Then
        If oRange.Worksheet.Name <> oRef.Worksheet.Name Then
            Call SetHLink(oRange, "''" & oRef.Worksheet.Name & "'!" & oRef.Address, "", "'" & oRef.Worksheet.Name & "'!" & oRef.Address)
        Else
            Call SetHLink(oRange, oRef.Address, "", "'" & oRef.Worksheet.Name & "'!" & oRef.Address)
        End If
    Else
        Call SetHLink(oRange, sTextToDisplay, "", "'" & oRef.Worksheet.Name & "'!" & oRef.Address)
    End If
End Sub

' セル参照リンク判定
Function isCellLink(ByVal oLink As Hyperlink) As Boolean
    isCellLink = (oLink.Address = "" And oLink.SubAddress <> "")
End Function

' セル参照リンクシート
Function GetCellLinkSheet(ByVal oLink As Hyperlink) As Worksheet
    If InStr(oLink.SubAddress, "!") > 0 Then
        Set GetCellLinkSheet = Application.Range(oLink.SubAddress).Worksheet
    Else
        Set GetCellLinkSheet = Application.Range("'" & oLink.parent.Worksheet.Name & "'!" & oLink.SubAddress).Worksheet
    End If
End Function

' セル参照リンク範囲
Function GetCellLinkRange(ByVal oLink As Hyperlink) As Range
    If InStr(oLink.SubAddress, "!") > 0 Then
        Set GetCellLinkRange = Application.Range(oLink.SubAddress)
    Else
        Set GetCellLinkRange = Application.Range("'" & oLink.parent.Worksheet.Name & "'!" & oLink.SubAddress)
    End If
End Function

' ファイルリンク設定
Sub SetFileLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sAddress As String, Optional ByVal bForce As Boolean = False)
    If (bForce = True) Or CreateFSO().FolderExists(sAddress) Or CreateFSO().FileExists(sAddress) Then
        ' 強制orファイル/フォルダが有るならリンクを設定
        Call SetHLink(oRange, sTextToDisplay, sAddress, "")
    Else
        Call oRange.Hyperlinks.Delete
    End If
End Sub

' ファイルリンク判定
Function isFileLink(ByVal oLink As Hyperlink) As Boolean
    isFileLink = (oLink.Address <> "" And oLink.SubAddress = "" And isURLLink(oLink) = False)
End Function

' URLリンク設定
Sub SetURLLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sAddress As String)
    Call SetHLink(oRange, sTextToDisplay, sAddress, "")
End Sub

' URLリンク判定
Function isURLLink(oLink As Hyperlink) As Boolean
    ' システム中で有効なURLスキームを全て確認できるわけではないので
    ' 一般的なURLスキームらしい":"の存在をチェックするだけにしている
    isURLLink = InStr(oLink.Address, ":") > 0
End Function

' ファイル検索リンク設定
Sub SetSearchLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sLocation As String, ByVal sSearch As String, Optional ByVal bForce As Boolean = False)
    If (bForce = True) Or CreateFSO().FolderExists(sLocation) Or CreateFSO().FileExists(sLocation) Then
        ' 強制orファイル/フォルダが有るならリンクを設定
        Call SetHLink(oRange, sTextToDisplay, "search-ms:query=" & sSearch & "&" & "crumb=location:" & sLocation, "")
    Else
        Call oRange.Hyperlinks.Delete
    End If
End Sub

' // コメント操作 /////////////////////////////////////////

' コメント設定
Sub SetComment(ByRef oRange As Range, ByVal sText As String, Optional ByVal bShow As Boolean = True)
    Call ClrComment(oRange)
    If sText <> "" Then
        With oRange.AddComment
            .Visible = bShow
            .Shape.TextFrame.AutoSize = True
            .Shape.TextFrame.Characters.Text = sText
        End With
    End If
End Sub

' コメント消去
Sub ClrComment(ByRef oRange As Range)
    If Not oRange.Comment Is Nothing Then
        oRange.ClearComments
    End If
End Sub

' コメントの自動調整
Sub AutoFitComment(ByRef oSheet As Worksheet)
    Dim cmt As Comment
    For Each cmt In oSheet.Comments
        cmt.Shape.Top = cmt.parent.offset(0, 1).Top
        cmt.Shape.Left = cmt.parent.offset(0, 1).Left
    Next
End Sub

' // ピボットテーブル操作 /////////////////////////////////
' ・とりあえず最低限

' ピボットテーブル検索
Function SearchPivotTable(ByVal oSheet As Worksheet, ByVal sName As String) As PivotTable
    Dim elm As PivotTable
    For Each elm In oSheet.PivotTables
        If elm.Name Like sName Then
            Exit For
        End If
    Next
    Set SearchPivotTable = elm
End Function

' // クエリ操作 ///////////////////////////////////////////
' ・とりあえず最低限

' クエリ検索
Function SearchWorkbookQuery(ByVal oBook As Workbook, ByVal sName As String) As WorkbookQuery
    Dim elm As WorkbookQuery
    For Each elm In oBook.Queries
        If elm.Name Like sName Then
            Exit For
        End If
    Next
    Set SearchWorkbookQuery = elm
End Function

' // WebService /////////////////////////////////////////////
' // ・WebAPI用処理各種
' // ・VBA-JSON(https://github.com/VBA-tools/VBA-JSON)辺りを併用する前提

'Sub Usage()
'    Dim sURL As String
'    Dim oHead As Collection: Set oHead = New Collection
'    Dim sBody As String: sBody = ""
'    sURL = EmbText("https://api.mouser.com/api/v2/search/keyword?apiKey={0}", sMSRKey)
'    oHead.Add Array("Content-Type", "application/json")
'    oHead.Add Array("Accept", "application/json")
'    sBody = EmbText("{ ""SearchByKeywordRequest"": { ""keyword"": ""{0}"" } }", sSrch)
'    sTrns = SendWebAPIRequest(sURL, "POST", oHead, sBody, slTimeOut)
'End Sub

' HTTP生成
Private Function CreateHTTP(Optional ByVal lTimeout As Long = 5000) As Object
    Static oHTTP As Object
    If oHTTP Is Nothing Then
        Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If oHTTP.readyState <> 0 Then
        oHTTP.abort
    End If
    Call oHTTP.SetTimeOuts(lTimeout, lTimeout, lTimeout, lTimeout)
    Set CreateHTTP = oHTTP
End Function

' WebAPI汎用要求送信
Function SendWebAPIRequest(ByVal sURL As String, ByVal sType As String, ByVal oHead As Collection, ByVal sBody As String, ByVal lTimeout As Long) As String
    ' HTTPリクエストの設定
    Dim oHTTP As Object
    Set oHTTP = CreateHTTP(lTimeout)

    ' HTTPリクエスト
    oHTTP.Open sType, sURL, False
    If Not oHead Is Nothing Then
        Dim elm
        For Each elm In oHead
            oHTTP.setRequestHeader elm(0), elm(1)
        Next
    End If
    If sBody <> "" Then
        oHTTP.send sBody
    Else
        oHTTP.send
    End If
    
    ' 結果取得
    Dim sRet As String
    If oHTTP.Status >= 200 And oHTTP.Status <= 299 Then
        sRet = oHTTP.responseText
    Else
        Debug.Print "SendWebAPIRequest ERR:" & sURL
        Debug.Print "SendWebAPIRequest ERR:" & oHTTP.Status
        Debug.Print "SendWebAPIRequest ERR:" & oHTTP.responseText
        sRet = "HTTP ERROR:" & oHTTP.Status
    End If
    SendWebAPIRequest = sRet
End Function

Function IsWebAPISuccess(ByVal sText As String) As Boolean
    Dim bRet As Boolean
    bRet = True
    If sText <> "" Then
        If Split(sText, ":")(0) = "HTTP ERROR" Then
            bRet = False
        End If
    End If
    IsWebAPISuccess = bRet
End Function

'Function EncodeURL(ByVal sText As String) As String
'    If Val(Application.Version) >= 15 Then
'        EncodeURL = WorksheetFunction.EncodeURL(sText)
'    Else
'        ' ※64bit版Excelでは動かない
'        With CreateObject("ScriptControl")
'            .Language = "JScript"
'            EncodeURL = .CodeObject.encodeURIComponent(sText)
'        End With
'    End If
'End Function

'------------------------------------------------------------------------------
' Function: EncodeURL
' Author: Jeremy Varnham
' GistURL: https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17
' Version: 1.1.0
' Date: 22 August 2024
' Description: Encodes a string into a URL-encoded format, supporting ASCII, Unicode, and UTF-8 encoding.
' Parameters:
'   - txt: The string to encode.
' Returns:
'   - The URL-encoded string.
'------------------------------------------------------------------------------
Function EncodeURL(ByRef txt As String) As String
    ' Declare and initialize variables
    Dim buffer As String
    Dim i As Long, c As Long, n As Long
    
    ' Initialize the buffer with enough space for the encoded string
    buffer = String$(Len(txt) * 12, "%")
    
    ' Loop through each character in the input string
    For i = 1 To Len(txt)
        ' Get the character code for the current character
        c = AscW(Mid$(txt, i, 1)) And 65535
        
        ' Determine if the character needs to be encoded or can be left as is
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped characters: 0-9, A-Z, a-z, - . _ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c) ' Add the character to the buffer
            Case Is <= 127            ' Escaped UTF-8 1 byte (U+0000 to U+007F) '
                n = n + 3
                Mid$(buffer, n - 2) = "%" ' Add the percent symbol
                Mid$(buffer, n - 1) = Right$("0" & Hex$(c), 2) ' Add the hex representation
            Case Is <= 2047           ' Escaped UTF-8 2 bytes (U+0080 to U+07FF) '
                n = n + 6
                Mid$(buffer, n - 5) = "%" ' Add the percent symbol
                Mid$(buffer, n - 4) = Right$("0" & Hex$(192 + (c \ 64)), 2) ' Add the first byte of the encoded character
                Mid$(buffer, n - 2) = "%" ' Add the percent symbol
                Mid$(buffer, n - 1) = Right$("0" & Hex$(128 + (c Mod 64)), 2) ' Add the second byte of the encoded character
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes (U+010000 to U+10FFFF) '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 11) = "%" ' Add the percent symbol
                Mid$(buffer, n - 10) = Right$("0" & Hex$(240 + (c \ 262144)), 2) ' Add the first byte
                Mid$(buffer, n - 8) = "%" ' Add the percent symbol
                Mid$(buffer, n - 7) = Right$("0" & Hex$(128 + ((c \ 4096) Mod 64)), 2) ' Add the second byte
                Mid$(buffer, n - 5) = "%" ' Add the percent symbol
                Mid$(buffer, n - 4) = Right$("0" & Hex$(128 + ((c \ 64) Mod 64)), 2) ' Add the third byte
                Mid$(buffer, n - 2) = "%" ' Add the percent symbol
                Mid$(buffer, n - 1) = Right$("0" & Hex$(128 + (c Mod 64)), 2) ' Add the fourth byte
            Case Else                 ' Escaped UTF-8 3 bytes (U+0800 to U+FFFF) '
                n = n + 9
                Mid$(buffer, n - 8) = "%" ' Add the percent symbol
                Mid$(buffer, n - 7) = Right$("0" & Hex$(224 + (c \ 4096)), 2) ' Add the first byte
                Mid$(buffer, n - 5) = "%" ' Add the percent symbol
                Mid$(buffer, n - 4) = Right$("0" & Hex$(128 + ((c \ 64) Mod 64)), 2) ' Add the second byte
                Mid$(buffer, n - 2) = "%" ' Add the percent symbol
                Mid$(buffer, n - 1) = Right$("0" & Hex$(128 + (c Mod 64)), 2) ' Add the third byte
        End Select
    Next
    
    ' Trim the buffer to the actual length of the encoded string
    EncodeURL = Left$(buffer, n)
End Function

'Function DecodeURL(ByVal sText As String) As String
'    ' ※64bit版Excelでは動かない
'    With CreateObject("ScriptControl")
'        .Language = "JScript"
'        DecodeURL = .CodeObject.decodeURIComponent(sText)
'    End With
'End Function

'------------------------------------------------------------------------------
' Function: DecodeURL
' Author: Jeremy Varnham
' GistURL: https://gist.github.com/jvarn/5e11b1fd741b5f79d8a516c9c2368f17
' Version: 1.1.0
' Date: 22 August 2024
' Description: Decodes a URL-encoded string, supporting ASCII, Unicode, and UTF-8 encoding.
' Parameters:
'   - strIn: The URL-encoded string to decode.
' Returns:
'   - The decoded string.
'------------------------------------------------------------------------------
Function DecodeURL(ByVal strIn As String) As String
    ' Declare and initialize variables
    Dim sl As Long, tl As Long
    Dim key As String, kl As Long
    Dim hh As String, hi As String, hl As String
    Dim a As Long
    
    ' Set the key to look for the percent symbol used in URL encoding
    key = "%"
    kl = Len(key)
    sl = 1: tl = 1

    ' Find the first occurrence of the key (percent symbol) in the input string
    sl = InStr(sl, strIn, key, vbTextCompare)
    
    ' Loop through the input string until no more percent symbols are found
    Do While sl > 0
        ' Add unprocessed characters to the result
        If (tl = 1 And sl <> 1) Or tl < sl Then
            DecodeURL = DecodeURL & Mid(strIn, tl, sl - tl)
        End If
        
        ' Determine the type of encoding (Unicode, UTF-8, or ASCII) and decode accordingly
        Select Case UCase(Mid(strIn, sl + kl, 1))
            Case "U"    ' Unicode URL encoding (e.g., %uXXXX)
                a = Val("&H" & Mid(strIn, sl + kl + 1, 4)) ' Convert hex to decimal
                DecodeURL = DecodeURL & ChrW(a) ' Convert decimal to character
                sl = sl + 6 ' Move to the next character after the encoded sequence
            Case "E"    ' UTF-8 URL encoding (e.g., %EXXX)
                hh = Mid(strIn, sl + kl, 2) ' Get the first two hex digits
                a = Val("&H" & hh) ' Convert hex to decimal
                If a < 128 Then
                    sl = sl + 3 ' Move to the next character
                    DecodeURL = DecodeURL & Chr(a) ' Convert to ASCII character
                Else
                    ' For multibyte UTF-8 characters
                    hi = Mid(strIn, sl + 3 + kl, 2) ' Get the next two hex digits
                    hl = Mid(strIn, sl + 6 + kl, 2) ' Get the final two hex digits
                    a = ((Val("&H" & hh) And &HF) * 2 ^ 12) Or ((Val("&H" & hi) And &H3F) * 2 ^ 6) Or (Val("&H" & hl) And &H3F)
                    DecodeURL = DecodeURL & ChrW(a) ' Convert to a wide character
                    sl = sl + 9 ' Move to the next character after the encoded sequence
                End If
            Case Else    ' Standard ASCII URL encoding (e.g., %XX)
                hh = Mid(strIn, sl + kl, 2) ' Get the two hex digits
                a = Val("&H" & hh) ' Convert hex to decimal
                If a < 128 Then
                    sl = sl + 3 ' Move to the next character
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2) ' Get the next two hex digits
                    a = ((Val("&H" & hh) - 194) * 64) + Val("&H" & hi) ' Convert to a character code
                    sl = sl + 6 ' Move to the next character after the encoded sequence
                End If
                DecodeURL = DecodeURL & ChrW(a) ' Convert to a wide character
        End Select
        
        ' Update the position of the last processed character
        tl = sl
        ' Find the next occurrence of the percent symbol
        sl = InStr(sl, strIn, key, vbTextCompare)
    Loop
    
    ' Append any remaining characters after the last percent symbol
    DecodeURL = DecodeURL & Mid(strIn, tl)
End Function
