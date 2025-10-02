Attribute VB_Name = "mdlTable"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // �e�[�u������
' // �EListObject�֘A�I�u�W�F�N�g����̃w���p���W���[��
' // �E���[�N�V�[�g�ƃe�[�u���͂P�΂P�ɑΉ����鎖�������Ƃ���
' // 20210901:����
' // 20211006:�e�[�u���s��̒ǉ��폜������ǉ�
' // 202208xx:GetTableRows�ǉ���
' // 20230523:Resize�C��
' // 20240523:�傫�߂̃e�R����

' // �e�[�u������ /////////////////////////////////////////

' �e�[�u���쐬
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

' �e�[�u���폜
Sub DelTable(ByVal oSheet As Worksheet)
    GetTable(oSheet).Unlist
End Sub
Sub DelTableByName(ByVal oSheet As Worksheet, Optional ByVal sTableName As String = "")
    GetTableByName(oSheet, sTableName).Unlist
End Sub

' �e�[�u���擾
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

' // �e�[�u���S��

' �e�[�u���S�̎擾      ���\�����Q�ƁFRange("�e�[�u����[#All]")
Function GetTableRng(ByVal oListObject As ListObject) As Range
    Set GetTableRng = oListObject.Range
End Function

' �e�[�u�����T�C�Y
Sub ExtendTable(ByVal oListObject As ListObject, ByVal lTop As Long, ByVal lLeft As Long, ByVal lBottom As Long, lRight)
    Call oListObject.Resize(ExtendRange(oListObject, lTop, lLeft, lBottom, lRight))
End Sub

' // �e�[�u���S�̗̂񑀍�
' // �E�S�e�[�u���͈͐���(�T�O�I�Ƀw�b�_�t�b�^���܂�)

' �e�[�u���񐔎擾
Function GetTableColCnt(ByVal oListObject As ListObject) As Long
    GetTableColCnt = oListObject.ListColumns.Count
End Function

' �e�[�u����ԍ��擾
Function GetTableColIdx(ByVal oListObject As ListObject, ByVal sColumn As String) As Long
    GetTableColIdx = oListObject.ListColumns(sColumn).Index
End Function

' �e�[�u���S��擾
Function GetTableCols(ByVal oListObject As ListObject) As ListObject
    Set GetTableCols = oListObject.ListColumns
End Function

' �e�[�u����擾(ListColumn)
Function GetTableCol(ByVal oListObject As ListObject, ByVal sColumn As String) As ListColumn
    Set GetTableCol = oListObject.ListColumns(sColumn)
End Function

' �e�[�u����ǉ�
Function AddTableCol(ByVal oListObject As ListObject) As ListColumn
    Set AddTableCol = oListObject.ListColumns.Add()
End Function

' �e�[�u����}��
Function InsTableCol(ByVal oListObject As ListObject, ByVal lPosition As Long, Optional ByVal sColumn As String = "") As ListColumn
    Dim oRet As ListColumn
    Set oRet = oListObject.ListColumns.Add(lPosition)
    If sColumn <> "" Then
        oRet.Name = sColumn
    End If
    Set InsTableCol = oRet
End Function

' �e�[�u����폜(�񖼏�)
Sub DelTableCol(ByVal oListObject As ListObject, ByVal sColumn As String)
    oListObject.ListColumns(sColumn).Delete
End Sub

' �e�[�u����폜(��ԍ�)
Sub DelTableColbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long)
    oListObject.ListColumns(lColIdx).Delete
End Sub

' // �e�[�u���S�̂̍s����
' // �E�S�e�[�u���͈͐���(�T�O�I�Ƀw�b�_�t�b�^���܂�)

' �e�[�u���s���擾
Function GetTableRowCnt(ByVal oListObject As ListObject) As Long
    GetTableRowCnt = oListObject.ListRows.Count
End Function

' �e�[�u���S�s�擾
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

' �e�[�u���S�s����
Sub ClrTable(ByVal oListObject As ListObject)
    If Not oListObject.DataBodyRange Is Nothing Then
        oListObject.DataBodyRange.Delete
    End If
End Sub

' // �e�[�u���w�b�_�[ /////////////////////////////////
' // �E��^�C�g���̃R�g

' �w�b�_�񐔎擾
Function GetTableHeaderCnt(ByVal oListObject As ListObject) As Long
    GetTableHeaderCnt = oListObject.HeaderRowRange.Count
End Function

' �w�b�_���S�̎擾      ���\�����Q�ƁFRange("�e�[�u����[#Header]")
Function GetTableHeaderRng(ByVal oListObject As ListObject) As Range
    Set GetTableHeaderRng = oListObject.HeaderRowRange
End Function

' �w�b�_���̎w��Z���擾(�񖼏�)
Function GetTableHeaderCell(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableHeaderCell = oListObject.ListColumns(sColumn).Range(1, 1)
End Function

' �w�b�_���̎w��Z���擾(��ԍ�)
Function GetTableHeaderCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Set GetTableHeaderCellbyIdx = oListObject.HeaderRowRange(lColIdx)
End Function

' �w�b�_���̎w��Z���擾(�e�[�u�����C�ӃZ��)
Function GetTableHeaderCellbyElm(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Set GetTableHeaderCellbyElm = Intersect(oListObject.HeaderRowRange, oCell.EntireColumn) ' �w��Z����ƃe�[�u���w�b�_�̈�̐ϏW����������̈�
End Function

' // �e�[�u���t�b�^�| /////////////////////////////////
' // �E�W�v�s�̃R�g

' �t�b�^�񐔎擾
Function GetTableFooterCnt(ByVal oListObject As ListObject) As Long
    GetTableFooterCnt = oListObject.TotalsRowRange.Count
End Function

' �t�b�^���S�̎擾      ���\�����Q�ƁFRange("�e�[�u����[#Totals]")
Function GetTableFooterRng(ByVal oListObject As ListObject) As Range
    Set GetTableFooterRng = oListObject.TotalsRowRange
End Function

' �t�b�^���̎w��Z���擾(�񖼏�)
Function GetTableFooterCell(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableFooterCell = Intersect(oListObject.TotalsRowRange, oListObject.ListColumns(sColumn).Range)
End Function

' �t�b�^���̎w��Z���擾(��ԍ�)
Function GetTableFooterCellbyIdx(ByVal oListObject As ListObject, ByVal lColIdx As Long) As Range
    Set GetTableFooterCellbyIdx = oListObject.TotalsRowRange(lColIdx)
End Function

' �t�b�^���̎w��Z���擾(�e�[�u�����C�ӃZ��)
Function GetTableFooterCellbyElm(ByVal oListObject As ListObject, ByVal oCell As Range) As Range
    Set GetTableFooterCellbyElm = Intersect(oListObject.TotalsRowRange, oCell.EntireColumn) ' �w��Z����ƃe�[�u���w�b�_�̈�̐ϏW����������̈�
End Function

' // �e�[�u���f�[�^�� /////////////////////////////////

' �f�[�^�s��
Function GetTableBodyCnt(ByVal oListObject As ListObject) As Long
    GetTableBodyCnt = oListObject.ListRows.Count
End Function

' �f�[�^���S��          ���\�����Q�ƁFRange("�e�[�u����[#Data]")
Function GetTableBodyRng(ByVal oListObject As ListObject) As Range
    Set GetTableBodyRng = oListObject.DataBodyRange
End Function

' �f�[�^���̎w���
Function GetTableBodyColRng(ByVal oListObject As ListObject, ByVal sColumn As String) As Range
    Set GetTableBodyColRng = oListObject.ListColumns(sColumn).DataBodyRange
End Function

' �f�[�^���̎w��s
Function GetTableBodyRowRng(ByVal oListObject As ListObject, ByVal lRowIdx As Long) As Range
    Set GetTableBodyRowRng = oListObject.ListRows(lRowIdx).Range
End Function

' �f�[�^���̎w��Z��(�s�ԍ�/�񖼏�)
Function GetTableBodyCell(ByVal oListObject As ListObject, ByVal lRowIdx As Long, ByVal sColumn As String) As Range
    Set GetTableBodyCell = oListObject.ListColumns(sColumn).DataBodyRange(lRowIdx)
End Function

' �f�[�^���̎w��Z��(�s�ԍ�/��ԍ�)
Function GetTableBodyCellByIdx(ByVal oListObject As ListObject, ByVal lRowIdx As Long, ByVal lColIdx As Long) As Range
    Set GetTableBodyCellByIdx = oListObject.DataBodyRange.Cells(lRowIdx, lColIdx)
End Function

' �f�[�^���̒T���J�����ɑΉ����錋�ʃJ������ԋp

' ���b�N�A�b�v��
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

' �f�B�N�V���i����
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

' // �e�[�u���f�[�^���t�B���^ /////////////////////////////
' // �E�t�B���^�ݒ��Excel�@�\���̂��̂ł��邽�߁A���Ȃ菈�����d�����Ƃɒ���
' // �E�蓮/�R�[�h�Ńt�B���^(��\����Ԃ�)�����f�[�^���e�[�u���f�[�^���擾�֘A�����ł͕��ʂɎQ�Ƃł���
' // �E�蓮/�R�[�h�Ńt�B���^(��\����Ԃ�)�����f�[�^�𖳎��������ꍇ��SpecialCells(xlCellTypeVisible)�Ő؂蕪����

' �e�[�u���Ƀt�B���^��K�p
Sub AddTableFilter(ByVal oListObject As ListObject, ByVal sColumn As String, ByVal sCriteria As String, Optional ByVal lOperator As XlAutoFilterOperator = xlAnd)
    Call oListObject.Range.AutoFilter(oListObject.ListColumns(sColumn).Index, sCriteria, lOperator)
End Sub

' �e�[�u���̃t�B���^���N���A
' Excel�@�\�̃t�B���^�͏d���̂Œ���
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

' // �e�[�u���f�[�^���\�[�g ///////////////////////////////

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

' // �e�[�u���̉E�N���b�N���j���[ /////////////////////////
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
