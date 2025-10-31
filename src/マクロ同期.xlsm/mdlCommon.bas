Attribute VB_Name = "mdlCommon"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // �ǂ��ɂł��g�������ȎG���ȏ���
' // 20210901:����
' // 202208xx:InputNum�s��C��
' // 20221101:���[�x���V���^�C�������ǉ�
' // 20230221:�@�\�ǉ�(SetStatusBar�ǉ�/SearchBook�ǉ�/CleansingText�폜/etc)
' // 20230308:Office2010�p��WebService�֐����b��쐬
' // 20230308:�P�Z���P���������p�����ǉ�
' // 20230308:Excel2021�ł͔����Ȃ̂ŃE�B���h�E����폜
' // 20230412:�V�[�g�\����Ԑݒ�ǉ�
' // 20230720:�V�[�g�ǉ��폜���ʒǉ�
' // 20240401:�u�b�N�֘A������g�[
' //          �V�[�g�E���[�N�V�[�g�E�`���[�g�V�[�g�̈����𖾊m�ɕ���
' //          ���[�N�V�[�g�ǉ�/�������̋����𒲐�
' //          �ʏ팟���ǉ�&�ȈՌ����p�~
' //          Range�̕�W��/���W���Z�o�@�\�ǉ�
' //          �P�Z���P�����`����"'"��"="�ɂ��Ă̐�����C��
' //          ���[�x���V���^�C���䗦�Z�o�ǉ�
' //          �R�����g�ݒ�֘A�����ǉ�
' //          �n�C�p�[�����N�֘A�����ǉ�
' //          WebService�p����WebAPI�p�����e����C
' //          ByVal/ByRef��߂�l�̌^�w���O��
' // 20251009:���K�\�����C�u���������ւ������C�Z���X�I�ɔ����Ȃ̂ŕ���

Public Const csELPtrn As String = "[" & vbCr & vbLf & "]"   ' TrimEx�p

Public Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer

' // ���p /////////////////////////////////////////////////

' FileSystemObject����
Private Function CreateFSO() As Object
    Static oFSO As Object
    If oFSO Is Nothing Then
        Set oFSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set CreateFSO = oFSO
End Function

' // Excel��ʑ��� ////////////////////////////////////////

' ���\���p���b�Z�[�W�{�b�N�X
Function InfBox(ByVal sTitle As String, ByVal sMessage As String) As Long
    InfBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKOnly Or vbInformation)
End Function

' �ӎv�m�F�p���b�Z�[�W�{�b�N�X
Function AskBox(ByVal sTitle As String, ByVal sMessage As String, Optional ByVal bDefaultOK As Boolean = False) As Long
    If bDefaultOK Then
        AskBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKCancel Or vbInformation)
    Else
        AskBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKCancel Or vbDefaultButton2 Or vbExclamation)
    End If
End Function

' �G���[�\���p���b�Z�[�W�{�b�N�X
Function ErrBox(ByVal sTitle As String, ByVal sMessage As String) As Long
    ErrBox = MsgBox(Title:=sTitle, Prompt:=sMessage, Buttons:=vbOKOnly Or vbCritical Or vbSystemModal)
End Function

' ���̓{�b�N�X(���l)
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

' ���̓{�b�N�X(�e�L�X�g)
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

' ���̓{�b�N�X(�͈�)
Function InputRng(ByVal sTitle As String, ByVal sMessage As String, ByVal oDefault As Range, ByRef sResult As Range) As Boolean
    Dim bRet As Boolean
    Dim oVal As Range
On Error GoTo ErrExit
    Dim sDef As String
    If Not oDefault Is Nothing Then
        sDef = oDefault.Address
    End If
    ' ���[�U�L�����Z�����쎞�ɔn����InputBox��Boolean��ݒ肵�悤�Ƃ���Set����O�𔭐�����...orz
    ' �ǂ������̗�O���������̂͌��\�Ȏ�ԂȂ悤�Ȃ̂ŁA���s���ăG���[��Ԃ��m�F���đΏ����g���ۂ�
    Set sResult = Application.InputBox(Prompt:=sMessage, Title:=sTitle, Default:=sDef, Type:=8)
    bRet = True
NrmExit:
    InputRng = bRet
    Exit Function
ErrExit:
    Set sResult = Nothing
    Resume NrmExit
End Function

' �X�e�[�^�X�o�[�ݒ�
Public Sub SetStatusBar(ByVal sText As String)
    Application.StatusBar = sText
End Sub

' �X�e�[�^�X�o�[�i��
Public Sub SetStatusBarProgress(ByVal sText As String, ByVal lIdx As Long, ByVal lMax As Long)
    Application.StatusBar = sText & " " & Left(String(Int(lIdx / lMax * 10), "��") & String(10, "��"), 10) & "(" & lIdx & "/" & lMax & ")"
End Sub

' �X�e�[�^�X�o�[����
Public Sub ClrStatusBar()
    Application.StatusBar = False
End Sub

' �E�N���b�N���j���[�ǉ�
Sub AddRClickMenu(ByVal sTitle As String, ByVal sMacro As String)
    Dim oCmdBar As CommandBarButton
    Set oCmdBar = Application.CommandBars("Cell").Controls.Add(Temporary:=True)
    With oCmdBar
        .Caption = sTitle
        .OnAction = sMacro
    End With
End Sub

' �E�N���b�N���j���[�N���A
Sub ClrRClickMenu()
    Call Application.CommandBars("Cell").Reset
End Sub

' // �u�b�N���� ///////////////////////////////////////////

' �u�b�N�쐬
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
    ' �f�t�H�ŒP��V�[�g���܂ރu�b�N���쐬
    Dim oBook As Workbook
    Set oBook = Application.Workbooks.Add(xlWBATWorksheet)
    ' �e���v���[�g�V�[�g���R�s�[
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

' �u�b�N�I�[�v��
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
    ' �u�b�N�I�[�v��
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

' �e���v���[�g�u�b�N�I�[�v��
' �E�e���v���[�g�`���ł͂Ȃ��C�ӂ̃t�@�C�����e���v���[�g�Ƃ��ă�������ɊJ��
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
    ' �e���v���[�g�u�b�N�I�[�v��
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

' �u�b�N����
Function SearchBook(ByVal sBook As String) As Workbook
    Dim elm As Workbook
    For Each elm In Application.Workbooks
        If elm.Name Like sBook Then
            Exit For
        End If
    Next
    Set SearchBook = elm
End Function

' �u�b�N����
Function BookName(ByRef oBook As Workbook, Optional ByVal bExtension As Boolean = True) As String
    If bExtension = True Then
        BookName = oBook.Name
    Else
        BookName = CreateFSO().GetBaseName(oBook.Name)
    End If
End Function

' �u�b�N�̕\����Ԃ�ݒ�
Sub SetBookVisibleState(ByRef oBook As Workbook, ByVal bVisible As Boolean)
    Application.Windows(oBook.Name).Visible = bVisible
End Sub

' �u�b�N�̍ŏI�ŏ�Ԃ�ݒ�
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

' �u�b�N�����ݒ�
' �ETitle           : �^�C�g��
' �ESubject         : �T�u�^�C�g��
' �ECompany         : ��Ж�
' �EAuthor          : �쐬��
' �ELast Author     : �X�V��
' �EKeywords        : �L�[���[�h
' �EComments        : �R�����g
' �ERevision Number : �����ԍ�
' �ESecurity        : �Z�L�����e�B
' �EHyperlink Base  : �n�C�p�[�����N�̊�_
Function SetBookProp(ByRef oBook As Workbook, ByVal sProp As String, ByVal sText As String)
    oBook.BuiltinDocumentProperties(sProp).Value = sText
End Function

' �u�b�N�����N���A
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

' // �V�[�g���� ///////////////////////////////////////////
' // �E�V�[�g�ɂ̓��[�N�V�[�g�ƃ`���[�g�V�[�g�̓��ނ�����

' �V�[�g�m�F
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

' ���[�N�V�[�g����
Function SearchWorkSheet(ByVal oBook As Workbook, ByVal sName As String) As Worksheet
    Dim elm As Worksheet
    For Each elm In oBook.Worksheets
        If elm.Name = sName Then
            Exit For
        End If
    Next
    Set SearchWorkSheet = elm
End Function

' �`���[�g�V�[�g����
Function SearchChartSheet(ByVal oBook As Workbook, ByVal sName As String) As Chart
    Dim elm As Chart
    For Each elm In oBook.Charts
        If elm.Name = sName Then
            Exit For
        End If
    Next
    Set SearchChartSheet = elm
End Function

' ���[�N�V�[�g�ǉ�
Function AddWorkSheet(ByRef oBook As Workbook, ByVal sName As String) As Worksheet
    Dim oRet As Worksheet
    Set oRet = oBook.Sheets.Add(After:=oBook.Sheets(oBook.Sheets.Count))
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' �d�����Ȃ����̂������ݒ�
    Set AddWorkSheet = oRet
End Function

' �`���[�g�V�[�g�ǉ�
Function AddChartSheet(ByRef oBook As Workbook, ByVal sName As String) As Chart
    Dim oRet As Chart
    Set oRet = oBook.Sheets.Add(After:=oBook.Sheets(oBook.Sheets.Count))
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' �d�����Ȃ����̂������ݒ�
    Set AddChartSheet = oRet
End Function

' ���[�N�V�[�g����
Function CopyWorkSheet(ByRef oBook As Workbook, ByRef oSheet As Worksheet, ByVal sName As String) As Worksheet
    Dim oRet As Worksheet
    Call oSheet.Copy(After:=oBook.Sheets(oBook.Sheets.Count))   ' �߂�l�������̂Ńu�b�N�����ɌŒ�z�u
    Set oRet = oBook.Sheets(oBook.Sheets.Count)                 ' �߂�l�������̂Ńu�b�N�����ɌŒ�z�u
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' �d�����Ȃ����̂������ݒ�
    oRet.Visible = xlSheetVisible                               ' �R�s�[��ɕs���ł͍���̂ŉ�����
    Set CopyWorkSheet = oRet
End Function

' �`���[�g�V�[�g����
Function CopyChartSheet(ByRef oBook As Workbook, ByRef oSheet As Chart, ByVal sName As String) As Chart
    Dim oRet As Chart
    Call oSheet.Copy(After:=oBook.Sheets(oBook.Sheets.Count))   ' �߂�l�������̂Ńu�b�N�����ɌŒ�z�u
    Set oRet = oBook.Sheets(oBook.Sheets.Count)                 ' �߂�l�������̂Ńu�b�N�����ɌŒ�z�u
    oRet.Name = GenUniqSheetName(oBook, sName)                  ' �d�����Ȃ����̂������ݒ�
    oRet.Visible = xlSheetVisible                               ' �R�s�[��ɕs���ł͍���̂ŉ�����
    Set CopyChartSheet = oRet
End Function

' �u�b�N���Ń��j�[�N�ȃV�[�g�����쐬����
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

' �V�[�g�폜
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

' �V�[�g�ړ�
Sub MoveSheet(ByRef oBook As Workbook, ByVal oSheet As Worksheet, Optional ByVal bTop As Boolean = True)
    If bTop Then
        Call oSheet.Move(Before:=oBook.Sheets(1))
    Else
        Call oSheet.Move(After:=oBook.Sheets(oBook.Sheets.Count))
    End If
End Sub

' �V�[�g���בւ�
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

' �u�b�N���̎w��V�[�g�݂̂�\������
Sub SetVisibleSheet(ByRef oBook As Workbook, ParamArray sSheetPtrns())
    ' ��\�����\��
    Dim elm As Worksheet
    Dim sSheetPtrn As Variant
    For Each elm In oBook.Sheets
        For Each sSheetPtrn In sSheetPtrns
            If elm.Name Like sSheetPtrn Then
                elm.Visible = xlSheetVisible
            End If
        Next
    Next
    ' �\������\��
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

' // �͈͑��� /////////////////////////////////////////////

' �������Range�ɕϊ�
Function CRng(sRange As String) As Range
    CRng = Application.Range(sRange)
End Function

' ��ԍ���A1�Q�ƌ`���̗񖼕���(��:10��ځ�J��)
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

' A1�Q�ƌ`���̗񖼕�������ԍ�(��:J��10���)
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

' �ȈՊJ�n�s�ԍ��擾
Function MinRow(ByVal oSheet As Worksheet, lCol As Long) As Long
    If oSheet.Cells(1, lCol).Value = "" Then
        MinRow = oSheet.Cells(1, lCol).End(xlDown).Row
    Else
        MinRow = 1
    End If
End Function

' �ȈՍŏI�s�ԍ��擾
Function MaxRow(ByVal oSheet As Worksheet, lCol As Long) As Long
    If oSheet.Cells(oSheet.Rows.Count, lCol).Value = "" Then
        MaxRow = oSheet.Cells(oSheet.Rows.Count, lCol).End(xlUp).Row
    Else
        MaxRow = oSheet.Rows.Count
    End If
End Function

' �ȈՊJ�n��ԍ��擾
Function MinCol(ByVal oSheet As Worksheet, lRow As Long) As Long
    If oSheet.Cells(lRow, 1).Value = "" Then
        MinCol = oSheet.Cells(lRow, 1).End(xlToRight).Column
    Else
        MinCol = 1
    End If
End Function

' �ȈՍŏI��ԍ��擾
Function MaxCol(ByVal oSheet As Worksheet, lRow As Long) As Long
    If oSheet.Cells(lRow, oSheet.Columns.Count).Value = "" Then
        MaxCol = oSheet.Cells(lRow, oSheet.Columns.Count).End(xlToLeft).Column
    Else
        MaxCol = oSheet.Columns.Count
    End If
End Function

' �͈͊g��
Function ExtendRange(ByVal oRange As Range, ByVal lTop As Long, ByVal lLeft As Long, ByVal lBottom As Long, ByVal lRight) As Range
    Set ExtendRange = oRange.offset(-lTop, -lLeft).Resize(oRange.Rows.Count + lTop + lBottom, oRange.Columns.Count + lLeft + lRight)
End Function

' �Z���������l�������s�g��
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

' �Z���������l��������g��
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

' �͈͕�W��
Function NotRange(ByVal oRange As Range, Optional ByVal oSheet As Worksheet = Nothing) As Range
    Dim oResult As Range
    If Not oRange Is Nothing Then
        Set oResult = oRange.Worksheet.Cells
    Else
        Set oResult = oSheet.Cells ' oRange��oSheet��Nothing�Ȃ玀�ʂ̂�TPO�ɍ��킹�Ăǂ����B
    End If
    ' �S�̏W���ƌX�̔���̈�̐ς����
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
    ' �w��͈͂̏㑤(���̕���)
    '  ������
    '  ���~��
    '  ������
    idx = oRange.Item(1).Row - 1
    If idx > 0 Then
        Set oResult = UnionRange(oResult, oSheet.Range(oSheet.Rows(1), oSheet.Rows(idx)))
    End If
    '�w��͈͂̉���(���̕���)
    '  ������
    '  ���~��
    '  ������
    idx = oRange.Item(oRange.Rows.Count, oRange.Columns.Count).Row + 1
    If idx < oSheet.Rows.Count Then
        Set oResult = UnionRange(oResult, oSheet.Range(oSheet.Rows(idx), oSheet.Rows(oSheet.Rows.Count)))
    End If
    '�w��͈͂̍���(���̕���)
    '  ������
    '  ���~��
    '  ������
    idx = oRange.Item(1).Column - 1
    If idx > 0 Then
        Set rng = Intersect(oSheet.Range(oSheet.Columns(1), oSheet.Columns(idx)), oRange.EntireRow)
        Set oResult = UnionRange(oResult, rng)
    End If
    '�w��͈͂̉E��(���̕���)
    '  ������
    '  ���~��
    '  ������
    idx = oRange.Item(oRange.Rows.Count, oRange.Columns.Count).Column + 1
    If idx < oSheet.Columns.Count Then
        Set rng = Intersect(oSheet.Range(oSheet.Columns(idx), oSheet.Columns(oSheet.Columns.Count)), oRange.EntireRow)
        Set oResult = UnionRange(oResult, rng)
    End If
    Set coNotRange = oResult
End Function

' �͈͘a�W��
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

' �͈͐ϏW��
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

' �͈͍��W��
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
    ' ����̈悩��E��̈����藎�Ƃ�
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

' �͈͂ɂP�����z����ꊇ�o��
' �E�z��T�C�Y����o�͔͈͂���������
Sub WriteRange1D(ByRef oRange As Range, ByVal vArray As Variant)
    oRange.Resize(UBound(vArray) - LBound(vArray) + 1).Value = vArray
End Sub

' �͈͂ɂQ�����z����ꊇ�o��
' �E�z��T�C�Y����o�͔͈͂���������
Sub WriteRange2D(ByRef oRange As Range, ByVal vArray As Variant)
    oRange.Resize(UBound(vArray, 1) - LBound(vArray, 1) + 1, UBound(vArray, 2) - LBound(vArray, 2) + 1).Value = vArray
End Sub

' �͈͌���
Public Function FindRange( _
    ByVal oRange As Range, _
    ByVal sText As String, _
    Optional ByVal LookIn As XlFindLookIn = xlValues, _
    Optional ByVal LookAt As XlLookAt = xlPart, _
    Optional ByVal MatchCase As Boolean = False, _
    Optional ByVal MatchByte As Boolean = False, _
    Optional ByVal SearchFormat As Boolean = False _
) As Range
    ' Excel�̃o�O�H�Ō����Z�������o�ł��Ȃ��Ȃ邽��
    ' SearchOrder:=xlByColumns�͎w��ł��Ȃ����Ă���
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

'' �ȈՌ���
'Function SearchRange(oRange As Range, ByVal sText As String) As Range
'    ' ��Find�Ƃ�Match�͓��t���݂̈��������ɓ���̂ŒP������
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

'' �ȈՌ���
'Function SearchRanges(ByVal oRange As Range, ByVal sText As String) As Collection
'    ' ��Find�Ƃ�Match�͓��t���݂̈��������ɓ���̂ŒP������
'    Dim oRet As New Collection
'    Dim elm As Range
'    For Each elm In oRange.Cells
'        If elm.Text Like sText Then
'            Call oRet.Add(elm)
'        End If
'    Next
'    Set SearchRanges = oRet
'End Function

' // ����Excel�p //////////////////////////////////////////

' �P�Z���P�����`���Ǐo
' �E�P�Z���ɂP�����ȏ�̕��������͂��Ă���ꍇ�͗�O��f���܂�
' �E�����͂̋�Z���͋󔒕��������͂��Ă�����̂ƌ��􂵂܂�
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
                Call Err.Raise(9999, "", "�P�Z���P�����Ƃ��Ă�������:" & elm.Address)
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

' �P�Z���P�����`�����o
' �E�����͈͎͂n�_�Z���P�_�{�����ɂ���Ďw�肵�܂�
' �E�܂�Ԃ��͕�����̌Ăяo���őΏ����Ă�������
Sub WriteCells(ByVal oDst As Range, ByVal sText As String, ByVal lLen As Long, Optional ByVal lAlign As Long = 0)
    ' �A���C�����g
    Dim sBuff As String
    Select Case lAlign
        Case 0: sBuff = Left(sText & String(lLen, " "), lLen)   ' ���l��
        Case 1: sBuff = Right(String(lLen, " ") & sText, lLen)  ' �E�l��
    End Select
    ' �ꊇ�]��
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

' �P�Z���P�����`������
' �E�����͈͎͂n�_�Z���P�_�{�����ɂ���Ďw�肵�܂�
' �E�܂�Ԃ��͕�����̌Ăяo���őΏ����Ă�������
Sub ClearCells(ByRef oDst As Range, ByVal lLen As Long)
    oDst.Resize(1, lLen).ClearContents
End Sub

' �P�Z���P�����`������
' �E�]����/�]���͈͎͂n�_�Z���P�_�{�����ɂ���Ďw�肵�܂�
' �E�܂�Ԃ��͕�����̌Ăяo���őΏ����Ă�������
Sub CopyCells(ByRef oDst As Range, ByVal oSrc As Range, ByVal lLen As Long)
    oDst.Resize(1, lLen).Value = oSrc.Resize(1, lLen).Value
End Sub

' // �֗��n ///////////////////////////////////////////////
' // �E����܂��ėp����Ȃ����ǁA�ǂ����R�s�y���邵�Ȃ�...�Ƃ����㕨

' �V�[�g����
Sub DecolateSheet( _
    ByRef oSheet As Worksheet, _
    ByRef oRange As Range, _
    Optional ByVal bHead As Boolean = True, _
    Optional ByVal bFilter As Boolean = True, _
    Optional ByVal lSort As Long = 0 _
)
    ' �r��
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
    
    ' �s��Ў�������
    oRange.Rows.AutoFit
    oRange.Columns.AutoFit
    
    ' �I�[�g�t�B���^�ݒ�
    If bHead And bFilter And oSheet.AutoFilterMode = False Then
        oRange.Columns.AutoFilter
    End If
    
    ' �\�[�g
    ' �E�\�[�g�C���f�b�N�X��0�Ȃ�\�[�g���Ȃ�
    ' �E�\�[�g�C���f�b�N�X��0�ȏ�Ȃ珸���\�[�g
    ' �E�\�[�g�C���f�b�N�X��0�ȉ��Ȃ�~���\�[�g
    If lSort <> 0 Then
        oRange.Columns.Sort Header:=IIf(bHead, xlYes, xlNo), Key1:=oRange.Columns(Abs(lSort)), Order1:=IIf(lSort > 0, xlAscending, xlDescending)
    End If
    
    ' �R�����g�ʒu��������
    ' �E�\�[�g��s��Ў��������ŃY����̂ŁA���ΕK�{
    Call AutoFitComment(oSheet)
End Sub

' �w��s�ȍ~���N���A
' �E�ŏ��̍s�͓��͋K�������c�����ߍ폜���Ȃ�
' �E����ȍ~��UsedRange���팸���邽�߂ɍ폜����
Sub ClearRows(ByRef oSheet As Worksheet, ByVal lRow As Long)
    Dim oRng1st As Range
    Dim oRngOth As Range
    Set oRng1st = oSheet.Cells(lRow + 0, 1).EntireRow
    Set oRngOth = oSheet.Range(oSheet.Cells(lRow + 1, 1), oSheet.Cells(oSheet.Rows.Count, 1)).EntireRow
    Call oRng1st.ClearContents
    Call oRngOth.Delete
End Sub

' �w���ȍ~���N���A
' �E�ŏ��̗�͓��͋K�������c�����ߍ폜���Ȃ�
' �E����ȍ~��UsedRange���팸���邽�߂ɍ폜����
Sub ClearCols(ByRef oSheet As Worksheet, ByVal lCol As Long)
    Dim oRng1st As Range
    Dim oRngOth As Range
    Set oRng1st = oSheet.Cells(1, lCol + 0).EntireColumn
    Set oRngOth = oSheet.Range(oSheet.Cells(1, lCol + 1), oSheet.Cells(1, oSheet.Columns.Count)).EntireColumn
    Call oRng1st.ClearContents
    Call oRngOth.Delete
End Sub

' // ���l���� /////////////////////////////////////////////

' �ŏ��l
Function Min(ParamArray oVals() As Variant) As Variant
    Min = WorksheetFunction.Min(oVals)
End Function

' �ő�l
Function Max(ParamArray oVals() As Variant) As Variant
    Max = WorksheetFunction.Max(oVals)
End Function

' ���ϒl
Function Ave(ParamArray oVals() As Variant) As Variant
    Ave = WorksheetFunction.Average(oVals)
End Function

' �W���΍�
Function StDev(ParamArray oVals() As Variant) As Variant
    StDev = WorksheetFunction.StDev(oVals)
End Function

' �l�̌ܓ�
Function Round(ByVal vVal As Variant, ByVal lDigit As Long) As Variant
    Round = WorksheetFunction.Round(vVal, lDigit) ' VBA�ȑf��Round���\�b�h�͋�s�^�ۂ߁B��ʓI�Ȏl�̌ܓ��͂������B
End Function

' �؂�̂�
Function RoundDown(ByVal vVal As Variant, ByVal lDigit As Long) As Variant
    RoundDown = WorksheetFunction.RoundDown(vVal, lDigit)
End Function

' �؂�グ
Function RoundUp(ByVal vVal As Variant, ByVal lDigit As Long) As Variant
    RoundUp = WorksheetFunction.RoundUp(vVal, lDigit)
End Function

' ���l�͈͐���
Function RestrictNum(ByVal vVal As Variant, Optional ByVal vMin As Variant, Optional ByVal vMax As Variant) As Variant
    If Not IsMissing(vMin) Then vVal = IIf(vVal > vMin, vVal, vMin)
    If Not IsMissing(vMax) Then vVal = IIf(vVal < vMax, vVal, vMax)
    RestrictNum = vVal
End Function

' // �e�L�X�g���� /////////////////////////////////////////

' �V���O���N�H�[�g
Function SQuote(ByVal sText As String)
    SQuote = "'" & sText & "'"
End Function

' �_�u���N�H�[�g
Function DQuote(ByVal sText As String)
    DQuote = """" & sText & """"
End Function

' ��������
Function EmbText(ByVal sText As String, ParamArray vParam() As Variant) As String
    Dim idx As Long
    Dim elm As Variant
    For Each elm In vParam
        sText = Replace(sText, "{" & idx & "}", elm)
        idx = idx + 1
    Next
    EmbText = sText
End Function

' NULL or �󕶎�����
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

' ��������(�z��)
Function ConcatArray(ByVal sSep As String, sText() As Variant) As String
    ConcatArray = Join(sText, sSep)
End Function

' ��������(�ϒ��z��)
Function ConcatArgs(ByVal sSep As String, ParamArray sText() As Variant) As String
    ConcatArgs = Join(sText, sSep)
End Function

' ��������(�R���N�V����)
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

' �󕶎��ȊO������(�z��)
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

' �󕶎��ȊO������(�ϒ��z��)
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

' �󕶎��ȊO������(�R���N�V����)
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

' ���[�w�蕶������
Function LTrimEx(ByVal sText As String, ByVal sChar As String) As String
    ' NOTE:TrimEx(sText, csELPtrn) ����������ΐ擪�����̂ǂ������悭�킩���CRLF��S���폜�ł���
    While Left(sText, 1) Like sChar
        sText = Right(sText, Len(sText) - 1)
    Wend
    LTrimEx = sText
End Function

' �E�[�w�蕶������
Function RTrimEx(ByVal sText As String, ByVal sChar As String) As String
    ' NOTE:TrimEx(sText, csELPtrn) ����������ΐ擪�����̂ǂ������悭�킩���CRLF��S���폜�ł���
    While Right(sText, 1) Like sChar
        sText = Left(sText, Len(sText) - 1)
    Wend
    RTrimEx = sText
End Function

' ���[�w�蕶������
Function TrimEx(ByVal sText As String, ByVal sChar As String) As String
    ' NOTE:TrimEx(sText, csELPtrn) ����������ΐ擪�����̂ǂ������悭�킩���CRLF��S���폜�ł���
    sText = LTrimEx(sText, sChar)
    sText = RTrimEx(sText, sChar)
    TrimEx = sText
End Function

' YYYYMMDD/YYMMDD�`���̕�������t�ɕϊ�
Function YYYYMMDD2Date(ByVal sText As String, Optional ByVal sFormat As String = "####/##/##") As Date
    YYYYMMDD2Date = CDate(Format(sText, sFormat))
End Function

' �a������t�H�[�}�b�g
Function FormatDate(ByVal oDate As Date) As String
    FormatDate = Format(oDate, "YYYY/MM/DD")
End Function

' �a�������t�H�[�}�b�g
Function FormatTime(ByVal oDate As Date) As String
    FormatTime = Format(oDate, "hh:mm")
End Function

' ����4��+��2��+��2��
Function YYYYMMDD(ByVal oDate As Date) As String
    YYYYMMDD = Format(oDate, "YYYYMMDD")
End Function

' ����4��+��2��
Function YYYYMM(ByVal oDate As Date) As String
    YYYYMM = Format(oDate, "YYYYMM")
End Function

' ��2��+��2��
Function MMDD(ByVal oDate As Date) As String
    MMDD = Format(oDate, "MMDD")
End Function

' ����4��
Function YYYY(ByVal oDate As Date) As String
    YYYY = Format(oDate, "YYYY")
End Function

' ����2��
Function YY(ByVal oDate As Date) As String
    YY = Format(oDate, "YY")
End Function

' ��2��
Function MM(ByVal oDate As Date) As String
    MM = Format(oDate, "MM")
End Function

' ��2��
Function DD(ByVal oDate As Date) As String
    DD = Format(oDate, "DD")
End Function

' �N�x
Function FinancialYear(ByVal oDate As Date) As String
    FinancialYear = Format(DateAdd("m", -3, oDate), "YYYY")
End Function

' ����(�s���I�h)
Function FinancialPeriod(ByVal oDate As Date, Optional ByVal s1H As String = "T1", Optional ByVal s2H As String = "T2") As String
    FinancialPeriod = IIf(Month(oDate) >= 4 And Month(oDate) <= 9, s1H, s2H)
End Function

' �l����(�N�H�[�^)
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

' ���ԓ�����
Function IsDuring(oTime As Date, oSTime As Date, oETime As Date)
    IsDuring = (oSTime <= oTime And oTime <= oETime)
End Function

' NULL��Ǒւ���
Public Function NZ(ByVal Value As Variant, Optional ByVal IsNullValue As Variant = Empty) As Variant
    If IsNull(Value) Then
        NZ = IsNullValue
    Else
        NZ = Value
    End If
End Function

' // �z�񑀍� /////////////////////////////////////////////

' �w��l�Ŗ��߂�ꂽ�C�Ӓ��̔z��𐶐�
Function Arrays(ByVal lNum As Long, ByVal vVal As Variant) As Variant
    Dim vRet As Variant
    ReDim vRet(lNum - 1) As Variant
    Dim idx As Long
    For idx = LBound(vRet) To UBound(vRet)
        vRet(idx) = vVal
    Next
    Arrays = vRet
End Function

' �Q�����z�񂩂�P�����z����X���C�X(�s����)
Function SliceArrayRow(ByVal vVal As Variant, ByVal lRow As Long) As Variant
    Dim vRet() As Variant
    ReDim vRet(LBound(vVal, 2) To UBound(vVal, 2))
    Dim idx As Long
    For idx = LBound(vVal, 2) To UBound(vVal, 2)
        vRet(idx) = vVal(lRow, idx)
    Next
    SliceArrayRow = vRet
End Function

' �Q�����z�񂩂�P�����z����X���C�X(�����)
Function SliceArrayCol(ByVal vVal As Variant, ByVal lCol As Long) As Variant
    Dim vRet() As Variant
    ReDim vRet(LBound(vVal, 1) To UBound(vVal, 1))
    Dim idx As Long
    For idx = LBound(vVal, 1) To UBound(vVal, 1)
        vRet(idx) = vVal(idx, lCol)
    Next
    SliceArrayCol = vRet
End Function

' // �n�C�p�[�����N���� ///////////////////////////////////

' �ėp�����N�ݒ�
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

' �Z���Q�ƃ����N�ݒ�
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

' �Z���Q�ƃ����N����
Function isCellLink(ByVal oLink As Hyperlink) As Boolean
    isCellLink = (oLink.Address = "" And oLink.SubAddress <> "")
End Function

' �Z���Q�ƃ����N�V�[�g
Function GetCellLinkSheet(ByVal oLink As Hyperlink) As Worksheet
    If InStr(oLink.SubAddress, "!") > 0 Then
        Set GetCellLinkSheet = Application.Range(oLink.SubAddress).Worksheet
    Else
        Set GetCellLinkSheet = Application.Range("'" & oLink.parent.Worksheet.Name & "'!" & oLink.SubAddress).Worksheet
    End If
End Function

' �Z���Q�ƃ����N�͈�
Function GetCellLinkRange(ByVal oLink As Hyperlink) As Range
    If InStr(oLink.SubAddress, "!") > 0 Then
        Set GetCellLinkRange = Application.Range(oLink.SubAddress)
    Else
        Set GetCellLinkRange = Application.Range("'" & oLink.parent.Worksheet.Name & "'!" & oLink.SubAddress)
    End If
End Function

' �t�@�C�������N�ݒ�
Sub SetFileLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sAddress As String, Optional ByVal bForce As Boolean = False)
    If (bForce = True) Or CreateFSO().FolderExists(sAddress) Or CreateFSO().FileExists(sAddress) Then
        ' ����or�t�@�C��/�t�H���_���L��Ȃ烊���N��ݒ�
        Call SetHLink(oRange, sTextToDisplay, sAddress, "")
    Else
        Call oRange.Hyperlinks.Delete
    End If
End Sub

' �t�@�C�������N����
Function isFileLink(ByVal oLink As Hyperlink) As Boolean
    isFileLink = (oLink.Address <> "" And oLink.SubAddress = "" And isURLLink(oLink) = False)
End Function

' URL�����N�ݒ�
Sub SetURLLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sAddress As String)
    Call SetHLink(oRange, sTextToDisplay, sAddress, "")
End Sub

' URL�����N����
Function isURLLink(oLink As Hyperlink) As Boolean
    ' �V�X�e�����ŗL����URL�X�L�[����S�Ċm�F�ł���킯�ł͂Ȃ��̂�
    ' ��ʓI��URL�X�L�[���炵��":"�̑��݂��`�F�b�N���邾���ɂ��Ă���
    isURLLink = InStr(oLink.Address, ":") > 0
End Function

' �t�@�C�����������N�ݒ�
Sub SetSearchLink(ByRef oRange As Range, ByVal sTextToDisplay As String, ByVal sLocation As String, ByVal sSearch As String, Optional ByVal bForce As Boolean = False)
    If (bForce = True) Or CreateFSO().FolderExists(sLocation) Or CreateFSO().FileExists(sLocation) Then
        ' ����or�t�@�C��/�t�H���_���L��Ȃ烊���N��ݒ�
        Call SetHLink(oRange, sTextToDisplay, "search-ms:query=" & sSearch & "&" & "crumb=location:" & sLocation, "")
    Else
        Call oRange.Hyperlinks.Delete
    End If
End Sub

' // �R�����g���� /////////////////////////////////////////

' �R�����g�ݒ�
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

' �R�����g����
Sub ClrComment(ByRef oRange As Range)
    If Not oRange.Comment Is Nothing Then
        oRange.ClearComments
    End If
End Sub

' �R�����g�̎�������
Sub AutoFitComment(ByRef oSheet As Worksheet)
    Dim cmt As Comment
    For Each cmt In oSheet.Comments
        cmt.Shape.Top = cmt.parent.offset(0, 1).Top
        cmt.Shape.Left = cmt.parent.offset(0, 1).Left
    Next
End Sub

' // �s�{�b�g�e�[�u������ /////////////////////////////////
' �E�Ƃ肠�����Œ��

' �s�{�b�g�e�[�u������
Function SearchPivotTable(ByVal oSheet As Worksheet, ByVal sName As String) As PivotTable
    Dim elm As PivotTable
    For Each elm In oSheet.PivotTables
        If elm.Name Like sName Then
            Exit For
        End If
    Next
    Set SearchPivotTable = elm
End Function

' // �N�G������ ///////////////////////////////////////////
' �E�Ƃ肠�����Œ��

' �N�G������
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
' // �EWebAPI�p�����e��
' // �EVBA-JSON(https://github.com/VBA-tools/VBA-JSON)�ӂ�𕹗p����O��

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

' HTTP����
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

' WebAPI�ėp�v�����M
Function SendWebAPIRequest(ByVal sURL As String, ByVal sType As String, ByVal oHead As Collection, ByVal sBody As String, ByVal lTimeout As Long) As String
    ' HTTP���N�G�X�g�̐ݒ�
    Dim oHTTP As Object
    Set oHTTP = CreateHTTP(lTimeout)

    ' HTTP���N�G�X�g
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
    
    ' ���ʎ擾
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
'        ' ��64bit��Excel�ł͓����Ȃ�
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
'    ' ��64bit��Excel�ł͓����Ȃ�
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
