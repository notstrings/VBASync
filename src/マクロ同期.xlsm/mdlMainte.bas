Attribute VB_Name = "mdlMainte"
Option Explicit

' // //////////////////////////////////////////////////////////////////////////
' // �}�N�������e���W���[��
' // 20251001:���ō쐬

' �}�N�������e���{
Public Sub Mainte()
    Application.EnableEvents = False
On Error GoTo ErrExit
    Dim oDataTable As ListObject
    Set oDataTable = GetTable(ThisWorkbook.Worksheets("Main"))
    If oDataTable Is Nothing Then
        GoTo NrmExit
    End If
    Dim lRow As Long
    For lRow = 1 To GetTableBodyRowCnt(oDataTable)
        Dim sSrcPath As String
        Dim sDstPath As String
        Dim oSrcBook As Workbook: Set oSrcBook = Nothing
        Dim oDstBook As Workbook: Set oDstBook = Nothing
        sSrcPath = Trim(GetTableBodyCell(oDataTable, lRow, "�]����"))
        sDstPath = Trim(GetTableBodyCell(oDataTable, lRow, "�]����"))
        Set oSrcBook = OpenBook(sSrcPath, True, True, True, True)
        Set oDstBook = OpenBook(sDstPath, False, True, True, True)
        Call SyncMacro(oDstBook, oSrcBook)
        If Not oSrcBook Is Nothing Then Call oSrcBook.Close(False)
        If Not oDstBook Is Nothing Then Call oDstBook.Close(True)
    Next
NrmExit:
    Application.EnableEvents = True
    Exit Sub
ErrExit:
    If Err.Number <> 0 Then Call MsgBox(Err.Description)
    GoTo NrmExit
End Sub

' �}�N�����W���[���𓯊�
Private Sub SyncMacro(ByRef oDstWorkBook As Workbook, ByVal oSrcWorkBook As Workbook)
    Dim oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Dim elm
    Dim mdl

    ' �e���|�����t�H���_�p��
    Dim sBasePath As String
    sBasePath = CombinePath(BookPath(), "Temp")
    Call RemoveDir(sBasePath, True)
    Call MakeDir(sBasePath)

    ' �@�]�����u�b�N�̃��W���[�����G�N�X�|�[�g
    Debug.Print "EXPORT START"
    For Each mdl In oSrcWorkBook.VBProject.VBComponents
        Debug.Print "EXPORT:" & mdl.Name
        Select Case mdl.Type
            Case 1:     Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".bas")) ' �W�����W���[��
            Case 2:     Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".cls")) ' �N���X
            Case 3:     Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".frm")) ' �t�H�[��
            Case 100:   Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".dcm")) ' �V�[�g
            Case Else:  Debug.Assert False
        End Select
    Next
    Debug.Print "EXPORT END"

    ' �A�]����}�N���u�b�N�ɑ��݂��邷�ׂẴ}�N�����폜����
    Debug.Print "DELETE START"
    Dim bCont As Boolean
    Do
        bCont = False
        For Each mdl In oDstWorkBook.VBProject.VBComponents
            Debug.Print "DELETE:" & mdl.Name
            Select Case mdl.Type
                Case 1:     Call oDstWorkBook.VBProject.VBComponents.Remove(mdl): bCont = True: Exit For    ' �W�����W���[��
                Case 2:     Call oDstWorkBook.VBProject.VBComponents.Remove(mdl): bCont = True: Exit For    ' �N���X
                Case 3:     Call oDstWorkBook.VBProject.VBComponents.Remove(mdl): bCont = True: Exit For    ' �t�H�[��
                Case 100:   Call mdl.CodeModule.DeleteLines(1, mdl.CodeModule.CountOfLines)                 ' �V�[�g
            End Select
        Next
    Loop While bCont
    Debug.Print "DELETE END"

    ' �B�]����փ��W���[�����C���|�[�g
    Debug.Print "IMPORRT START"
    For Each elm In EnumFile(sBasePath, "*.*")
        Dim sPath As String
        Dim sName As String
        Dim sExt As String
        sPath = elm
        sName = oFSO.GetBaseName(elm)
        sExt = LCase(oFSO.GetExtensionName(elm))
        If sName <> "mdlMainte" Then
            Debug.Print "ADD:" & sName
            Select Case sExt
                Case "bas": Call oDstWorkBook.VBProject.VBComponents.Import(sPath)  ' �W�����W���[��
                Case "cls": Call oDstWorkBook.VBProject.VBComponents.Import(sPath)  ' �N���X
                Case "frm": Call oDstWorkBook.VBProject.VBComponents.Import(sPath)  ' �t�H�[��
                Case "dcm"                                                          ' �V�[�g
                    For Each mdl In oDstWorkBook.VBProject.VBComponents
                        If mdl.Name = sName Then
                            Call mdl.CodeModule.AddFromFile(sPath)
                            Dim sHead As String
                            sHead = ""
                            sHead = sHead & "VERSION 1.0 CLASS" & vbCrLf
                            sHead = sHead & "BEGIN" & vbCrLf
                            sHead = sHead & "  MultiUse = -1  'True" & vbCrLf
                            sHead = sHead & "End"
                            If mdl.CodeModule.Lines(1, 4) = sHead Then
                                Call mdl.CodeModule.DeleteLines(1, 4)
                            End If
                        End If
                    Next
            End Select
        End If
    Next
    Debug.Print "IMPORRT END"
End Sub
