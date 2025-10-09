Attribute VB_Name = "mdlTable"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // �e�[�u������
' // �EListObject�֘A�I�u�W�F�N�g����̃w���p���W���[��
' // 20210901:����
' // 20211006:�e�[�u���s��̒ǉ��폜������ǉ�
' // 202208xx:GetTableRows�ǉ���
' // 20230523:Resize�C��
' // 20240523:�傫�߂̃e�R����
' // 202509XX:���p��

' // �e�[�u���S�� /////////////////////////////////////////////////////////////

' �e�[�u���쐬
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

' �e�[�u���폜
Sub DelTable(ByVal oSheet As Worksheet, Optional ByVal lIndex As Long = 1)
    GetTable(oSheet, lIndex).Unlist
End Sub
Sub DelTableByName(ByVal oSheet As Worksheet, Optional ByVal sTableName As String = "")
    GetTableByName(oSheet, sTableName).Unlist
End Sub

' �e�[�u���擾(�S)
Function GetTables(ByVal oSheet As Worksheet, Optional ByVal lIndex As Long = 1) As ListObjects
    Set GetTables = oSheet.ListObjects
End Function

' �e�[�u���擾(�P)
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

' �e�[�u���͈͑S�́��\�����Q�ƁFRange("�e�[�u����[#All]")
Function GetTableRng(ByVal oListObject As ListObject) As Range
    Set GetTableRng = oListObject.Range
End Function

' �e�[�u���͈̓��T�C�Y
Sub ExtendTableRng(ByVal oListObject As ListObject, ByVal lTop As Long, ByVal lLeft As Long, ByVal lBottom As Long, lRight)
    Call oListObject.Resize(ExtendRange(oListObject.Range, lTop, lLeft, lBottom, lRight))
End Sub

' // �e�[�u���񑀍� ///////////////////////////////////////////////////////////

' �e�[�u����
Function GetTableColCnt(ByVal oListObject As ListObject) As Long
    GetTableColCnt = oListObject.ListColumns.Count
End Function

' �e�[�u����ԍ�
Function GetTableColIdx(ByVal oListObject As ListObject, ByVal sColumn As String) As Long
    GetTableColIdx = oListObject.ListColumns(sColumn).Index
End Function

' �e�[�u����擾(�S)
Function GetTableCols(ByVal oListObject As ListObject) As ListColumns
    Set GetTableCols = oListObject.ListColumns
End Function

' �e�[�u����擾(�P)
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

' �e�[�u����ǉ�
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

' �e�[�u����폜(�񖼏�)
Sub DelTableCol(ByVal oListObject As ListObject, ByVal sColumn As String)
    oListObject.ListColumns(sColumn).Delete
End Sub

' �e�[�u����폜(��ԍ�)
Sub DelTableColbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long)
    oListObject.ListColumns(lColIdx).Delete
End Sub

' �e�[�u�����Ń��j�[�N�ȗ񖼂��쐬����
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

' // �e�[�u���s���� ///////////////////////////////////////////////////////////

' �e�[�u���s���擾
Function GetTableRowCnt(ByVal oListObject As ListObject) As Long
    GetTableRowCnt = oListObject.ListRows.Count
End Function

' �e�[�u���S�s�擾(ListRows)
Function GetTableRows(ByVal oListObject As ListObject) As ListRows
    Set GetTableRows = oListObject.ListRows
End Function

' �e�[�u���s�擾(ListRow)
Function GetTableRow(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As ListRow
    Set GetTableRow = oListObject.ListRows(lRowIdx)
End Function

' �e�[�u���s�ǉ�
Function AddTableRow(ByVal oListObject As ListObject) As ListRow
    Set AddTableRow = oListObject.ListRows.Add()
End Function

' �e�[�u���s�}��
Function InsTableRow(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As ListRow
    Set InsTableRow = oListObject.ListRows.Add(lRowIdx, True)
End Function

' �e�[�u���s�폜
Sub DelTableRow(ByVal oListObject As ListObject, ByVal lPosition As Long)
    oListObject.ListRows(lPosition).Delete
End Sub

' // �e�[�u���w�b�_�[ /////////////////////////////////////////////////////////
' // �E��^�C�g���s�͈́����ӁF�w�b�_�[��\���̏ꍇ�͎擾�ł��Ȃ���

' �w�b�_���S�̎擾���\�����Q�ƁFRange("�e�[�u����[#Header]")
Function GetTableHeaderRng(ByVal oListObject As ListObject) As Range
    Set GetTableHeaderRng = oListObject.HeaderRowRange
End Function

' �w�b�_���̎w��Z���擾(�񖼏�)
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

' �w�b�_���̎w��Z���擾(��ԍ�)
Function GetTableHeaderCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Dim oRet As Range
    If Not oListObject.HeaderRowRange Is Nothing Then
        If lColIdx >= 1 And lColIdx <= oListObject.HeaderRowRange.Count Then
            Set oRet = oListObject.HeaderRowRange(lColIdx)
        End If
    End If
    Set GetTableHeaderCellbyIdx = oRet
End Function

' �w�b�_���̎w��Z���擾(�e�[�u�����C�ӃZ��)
Function GetTableHeaderCellbyCell(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Dim oRet As Range
    If Not oListObject.HeaderRowRange Is Nothing Then
        Set oRet = IntersectRange(oListObject.HeaderRowRange, oCell.EntireColumn)
    End If
    Set GetTableHeaderCellbyCell = oRet
End Function

' // �e�[�u���t�b�^�| /////////////////////////////////////////////////////////
' // �E�W�v�s�͈́����ӁG�t�b�^�|��\���̏ꍇ�͎擾�ł��Ȃ�

' �t�b�^���S�̎擾���\�����Q�ƁFRange("�e�[�u����[#Totals]")
Function GetTableFooterRng(ByVal oListObject As ListObject) As Range
    Set GetTableFooterRng = oListObject.TotalsRowRange
End Function

' �t�b�^���̎w��Z���擾(�񖼏�)
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

' �t�b�^���̎w��Z���擾(��ԍ�)
Function GetTableFooterCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Dim oRet As Range
    If Not oListObject.TotalsRowRange Is Nothing Then
        If lColIdx >= 1 And lColIdx <= oListObject.TotalsRowRange.Count Then
            Set oRet = oListObject.TotalsRowRange(lColIdx)
        End If
    End If
    Set GetTableFooterCellbyIdx = oRet
End Function

' �t�b�^���̎w��Z���擾(�e�[�u�����C�ӃZ��)
Function GetTableFooterCellbyElm(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Dim oRet As Range
    If Not oListObject.TotalsRowRange Is Nothing Then
        Set oRet = IntersectRange(oListObject.TotalsRowRange, oCell.EntireColumn)
    End If
    Set GetTableFooterCellbyElm = oRet
End Function

' // �e�[�u���f�[�^�� /////////////////////////////////////////////////////////

' �f�[�^�s��
Function GetTableBodyColCnt(ByVal oListObject As ListObject) As Long
    GetTableBodyColCnt = oListObject.ListColumns.Count
End Function
Function GetTableBodyRowCnt(ByVal oListObject As ListObject) As Long
    GetTableBodyRowCnt = oListObject.ListRows.Count
End Function

' �e�[�u�����W�n�ϊ�
Function GetTableBodyRowIdx(ByVal oListObject As ListObject, ByVal oRange As Range) As Long
    GetTableBodyRowIdx = oRange.Row - oListObject.DataBodyRange.Row + 1
End Function
Function GetTableBodyColIdx(ByVal oListObject As ListObject, ByVal oRange As Range) As Long
    GetTableBodyColIdx = oRange.Column - oListObject.DataBodyRange.Column + 1
End Function

' �f�[�^�͈̓N���b�v
Function ClipTableBodyRange(ByVal oListObject As ListObject, ByVal oRange As Range) As Range
    Set ClipTableBodyRange = IntersectRange(oListObject.DataBodyRange, oRange)
End Function

' �f�[�^������
Sub ClrTableBody(ByVal oListObject As ListObject)
    If Not oListObject.DataBodyRange Is Nothing Then
        oListObject.DataBodyRange.Delete
    End If
End Sub

' �f�[�^���S�́��\�����Q�ƁFRange("�e�[�u����[#Data]")
Function GetTableBodyRng(ByVal oListObject As ListObject) As Range
    Set GetTableBodyRng = oListObject.DataBodyRange
End Function
'Sub sample()
'    Dim oDataTable As ListObject
'    Set oDataTable = GetTable(ThisWorkbook.ActiveSheet)
'    If oDataTable Is Nothing Then Exit Sub                          ' �e�[�u�����Ȃ��̂ŏI��
'    Dim oTarget As Range
'    Set oTarget = Selection                                         ' �e�[�u���Ɋ֌W�Ȃ��Z���͈�
'    Set oTarget = ClipTableBodyRange(oDataTable, oTarget)           ' �e�[�u���Ɋ֌W�Ȃ��Z���͈͂��e�[�u���͈͓��ɃN���b�v
'    If oTarget Is Nothing Then Exit Sub                             ' �e�[�u���͈͊O�Ȃ̂ŏI��
'    Dim lc As Long: lc = GetTableBodyColIdx(oDataTable, oTarget)    ' �e�[�u�������ł̃C���f�b�N�X�ԍ��ɕϊ�
'    Dim lr As Long: lr = GetTableBodyRowIdx(oDataTable, oTarget)    ' �e�[�u�������ł̃C���f�b�N�X�ԍ��ɕϊ�
'    GetTableBodyRng(oDataTable)(lr, lc) = "aaa"                     ' �ʒu�w��A�N�Z�X
'End Sub

' �f�[�^���̎w���
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

' �f�[�^���̎w��s
Function GetTableBodyRowRng(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As Range
    Set GetTableBodyRowRng = oListObject.ListRows(lRowIdx).Range
End Function

' �f�[�^���̎w��Z��(�s�ԍ�/�񖼏�)
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

' �T��

' ���b�N�A�b�v(�P����)
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

' ���b�N�A�b�v(�Q����)
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

' // �e�[�u���f�[�^���t�B���^ /////////////////////////////////////////////////
' // �E�蓮/�R�[�h�Ńt�B���^�����f�[�^���e�[�u���f�[�^���擾�֘A�����ł͕��ʂɎQ�Ƃł���
' // �E�蓮/�R�[�h�Ńt�B���^�����f�[�^�𖳎�����ꍇ��SpecialCells(xlCellTypeVisible)�Ő؂蕪��

' �e�[�u���Ƀt�B���^��K�p
Sub AddTableFilter(ByVal oListObject As ListObject, ByVal sColumn As String, ByVal sCriteria As String, Optional ByVal lOperator As XlAutoFilterOperator = xlAnd)
    Call oListObject.Range.AutoFilter(oListObject.ListColumns(sColumn).Index, sCriteria, lOperator)
End Sub

' �e�[�u���̃t�B���^���N���A
Sub ClrTableRowFilter(ByVal oListObject As ListObject)
    oListObject.AutoFilter.ShowAllData
End Sub

' �e�[�u���ɐݒ肷��t�B���^��������T�j�^�C�Y
Function SanitizeFilterText(ByVal sText) As String
    sText = Replace(sText, "*", "~*")
    sText = Replace(sText, "?", "~?")
    sText = Replace(sText, "=", "==")
    sText = Replace(sText, "<", "=<")
    sText = Replace(sText, ">", "=>")
    SanitizeFilterText = sText
End Function

' // �e�[�u���f�[�^���\�[�g ///////////////////////////////////////////////////

' �\�[�g���{
Sub ApplyTableSort(ByVal oListObject As ListObject, ByVal sColumn As String, Optional ByVal bMatchCase As Boolean = True, Optional ByVal lSortMethod As XlSortMethod = xlPinYin)
    With oListObject.Sort
        .MatchCase = bMatchCase
        .SortMethod = lSortMethod
        .Apply
    End With
End Sub

' �\�[�g�����ǉ�
Sub AddTableSort(ByVal oListObject As ListObject, ByVal sColumn As String, Optional ByVal lOrder As XlSortOrder = xlAscending, Optional ByVal lSortOn As XlSortOn = xlSortOnValues, Optional lDataOption As XlSortDataOption = xlSortTextAsNumbers)
    Call oListObject.Sort.SortFields.Add(key:=oListObject.ListColumns(sColumn).Range, SortOn:=lSortOn, Order:=lOrder, DataOption:=lDataOption)
End Sub

' �\�[�g��������
Sub ClrTableSort(ByVal oListObject As ListObject, ByVal sColumn As String)
    Call oListObject.Sort.SortFields.Clear
End Sub

' // �e�[�u���̉E�N���b�N���j���[ /////////////////////////////////////////////
' // �E���ʂ̃Z����ł̉E�N���b�N�ɂ͔������Ȃ�

' �E�N���b�N���j���[�ǉ�
Sub AddTableRClickMenu(ByVal sTitle As String, ByVal sMacro As String)
    Dim oCmdBar As CommandBarButton
    Set oCmdBar = Application.CommandBars("List Range Popup").Controls.Add(Temporary:=True)
    With oCmdBar
        .Caption = sTitle
        .OnAction = sMacro
    End With
End Sub

' �E�N���b�N���j���[�N���A
Sub ClrTableRClickMenu()
    Call Application.CommandBars("List Range Popup").Reset
End Sub

