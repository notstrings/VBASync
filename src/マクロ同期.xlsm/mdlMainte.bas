Attribute VB_Name = "mdlMainte"
Option Explicit

' // //////////////////////////////////////////////////////////////////////////
' // マクロメンテモジュール
' // 20251001:初版作成

' マクロメンテ実施
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
        sSrcPath = Trim(GetTableBodyCell(oDataTable, lRow, "転送元"))
        sDstPath = Trim(GetTableBodyCell(oDataTable, lRow, "転送先"))
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

' マクロモジュールを同期
Private Sub SyncMacro(ByRef oDstWorkBook As Workbook, ByVal oSrcWorkBook As Workbook)
    Dim oFSO
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Dim elm
    Dim mdl

    ' テンポラリフォルダ用意
    Dim sBasePath As String
    sBasePath = CombinePath(BookPath(), "Temp")
    Call RemoveDir(sBasePath, True)
    Call MakeDir(sBasePath)

    ' ①転送元ブックのモジュールをエクスポート
    Debug.Print "EXPORT START"
    For Each mdl In oSrcWorkBook.VBProject.VBComponents
        Debug.Print "EXPORT:" & mdl.Name
        Select Case mdl.Type
            Case 1:     Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".bas")) ' 標準モジュール
            Case 2:     Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".cls")) ' クラス
            Case 3:     Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".frm")) ' フォーム
            Case 100:   Call mdl.Export(CombinePath(sBasePath, mdl.Name & ".dcm")) ' シート
            Case Else:  Debug.Assert False
        End Select
    Next
    Debug.Print "EXPORT END"

    ' ②転送先マクロブックに存在するすべてのマクロを削除する
    Debug.Print "DELETE START"
    Dim bCont As Boolean
    Do
        bCont = False
        For Each mdl In oDstWorkBook.VBProject.VBComponents
            Debug.Print "DELETE:" & mdl.Name
            Select Case mdl.Type
                Case 1:     Call oDstWorkBook.VBProject.VBComponents.Remove(mdl): bCont = True: Exit For    ' 標準モジュール
                Case 2:     Call oDstWorkBook.VBProject.VBComponents.Remove(mdl): bCont = True: Exit For    ' クラス
                Case 3:     Call oDstWorkBook.VBProject.VBComponents.Remove(mdl): bCont = True: Exit For    ' フォーム
                Case 100:   Call mdl.CodeModule.DeleteLines(1, mdl.CodeModule.CountOfLines)                 ' シート
            End Select
        Next
    Loop While bCont
    Debug.Print "DELETE END"

    ' ③転送先へモジュールをインポート
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
                Case "bas": Call oDstWorkBook.VBProject.VBComponents.Import(sPath)  ' 標準モジュール
                Case "cls": Call oDstWorkBook.VBProject.VBComponents.Import(sPath)  ' クラス
                Case "frm": Call oDstWorkBook.VBProject.VBComponents.Import(sPath)  ' フォーム
                Case "dcm"                                                          ' シート
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
