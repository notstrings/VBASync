Attribute VB_Name = "mdlFile"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // 誰でも作りそうなファイル・パス・シェル操作
' // 20210901:初版
' // 202208xx:INIファイル操作追加
' // 20221101:類似ファイル探索追加
' // 20240401:ChooseSimilarFile/ChooseSimilarFolder修正
' //          ByVal/ByRefや戻り値の型指定を徹底
' // 20251010:コマンド存在チェック追加
' // 20251114:億劫だったAPI宣言の修正を実施

' CreateDirectory用
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExW" ( _
    ByVal hWnd As LongPtr, _
    ByVal pszPath As LongPtr, _
    ByVal psa As LongPtr _
) As Long

' PathRelativePathTo用
Private Declare PtrSafe Function PathRelativePathTo Lib "Shlwapi" Alias "PathRelativePathToW" ( _
    ByVal pszPath As LongPtr, _
    ByVal pszFrom As LongPtr, _
    ByVal dwAttrFrom As Long, _
    ByVal pszTo As LongPtr, _
    ByVal dwAttrTo As Long _
) As Long
Private Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10
Private Const FILE_ATTRIBUTE_NORMAL  As Integer = &H80

' FindFirst*用
Private Declare PtrSafe Function FindFirstFileEx Lib "kernel32" Alias "FindFirstFileExW" ( _
    ByVal lpFileName As LongPtr, _
    ByVal fInfoLevelId As FINDEX_INFO_LEVELS, _
    lpFindFileData As WIN32_FIND_DATA, _
    ByVal fSearchOp As FINDEX_SEARCH_OPS, _
    ByVal lpSearchFilter As LongPtr, _
    ByVal dwAdditionalFlags As Long _
) As LongPtr
Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileW" ( _
    ByVal hFindFile As LongPtr, _
    lpFindFileData As WIN32_FIND_DATA _
) As Long
Private Declare PtrSafe Function FindClose Lib "kernel32" ( _
    ByVal hFindFile As LongPtr _
) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Enum FINDEX_INFO_LEVELS
    FindExInfoStandard = 0&
    FindExInfoBasic = 1&
    FindExInfoMaxInfoLevel = 2&
End Enum
Private Enum FINDEX_SEARCH_OPS
    FindExSearchNameMatch = 0&
    FindExSearchLimitToDirectories = 1&
    FindExSearchLimitToDevices = 2&
    FindExSearchMaxSearchOp = 3&
End Enum
Private Const FIND_FIRST_EX_CASE_SENSITIVE = 1&
Private Const FIND_FIRST_EX_LARGE_FETCH = 2&
Private Const FIND_FIRST_EX_ON_DISK_ENTRIES_ONLY = 4&
Private Type FileTime
    LowDateTime                             As Long
    HighDateTime                            As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes                        As Long     ' ファイル属性
    ftCreationTime                          As FileTime ' 作成日
    ftLastAccessTime                        As FileTime ' 最終アクセス日
    ftLastWriteTime                         As FileTime ' 最終更新日
    nFileSizeHigh                           As Long     ' ファイルサイズ（上位３２ビット）
    nFileSizeLow                            As Long     ' ファイルサイズ（下位３２ビット）
    dwReserved0                             As Long     ' 予約済み。リパースタグ
    dwReserved1                             As Long     ' 予約済み。未使用
    cFileName(260 * 2 - 1)                  As Byte     ' ファイル名
    cAlternateFileName(14 * 2 - 1)          As Byte     ' 8.3形式のファイル名
End Type

' StrCmpLogicalW用
Declare PtrSafe Function StrCmpLogicalW Lib "SHLWAPI.DLL" ( _
    ByVal lpStr1 As LongPtr, _
    ByVal lpStr2 As LongPtr _
) As Long

' Ini用
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

' // 共用 /////////////////////////////////////////////////

' FileSystemObject生成
Private Function CreateFSO() As Object
    Static oFSO As Object
    If oFSO Is Nothing Then
        Set oFSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set CreateFSO = oFSO
End Function

' WScript.Shell生成
Private Function CreateWSH() As Object
    Static oWSH As Object
    If oWSH Is Nothing Then
        Set oWSH = CreateObject("WScript.Shell")
    End If
    Set CreateWSH = oWSH
End Function

' Shell.Application生成
Private Function CreateShellApp() As Object
    Static oSH As Object
    If oSH Is Nothing Then
        Set oSH = CreateObject("Shell.Application")
    End If
    Set CreateShellApp = oSH
End Function

' トリム(左端)
Private Function PathTrimL(ByVal sText As String, ByVal sChar As String) As String
    While Left(sText, 1) Like sChar
        sText = Right(sText, Len(sText) - 1)
    Wend
    PathTrimL = sText
End Function

' トリム(右端)
Private Function PathTrimR(ByVal sText As String, ByVal sChar As String) As String
    While Right(sText, 1) = sChar
        sText = Left(sText, Len(sText) - 1)
    Wend
    PathTrimR = sText
End Function

' トリム(両端)
Private Function PathTrim(ByVal sText As String, ByVal sChar As String) As String
    sText = PathTrimL(sText, sChar)
    sText = PathTrimR(sText, sChar)
    PathTrim = sText
End Function

' シングルクォート
Private Function PathSQuote(ByVal sText As String) As String
    PathSQuote = "'" & sText & "'"
End Function

' ダブルクォート
Private Function PathDQuote(ByVal sText As String) As String
    PathDQuote = """" & sText & """"
End Function

' // パス操作 /////////////////////////////////////////////

' パス連結
Function CombinePath(ParamArray sPaths() As Variant) As String
    Dim sPath As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    Dim sSubElm As Variant
    For Each sSubElm In sPaths
        sPath = oFSO.BuildPath(sPath, sSubElm)
    Next
    CombinePath = sPath
End Function

'Sub TestCombinePath()
'    Debug.Assert (CombinePath("", "hoge") = "hoge")
'    Debug.Assert (CombinePath("hoge", "") = "hoge")
'    Debug.Assert (CombinePath("hoge", "hugo") = "hoge\hugo")
'    Debug.Assert (CombinePath("hoge", "hugo\") = "hoge\hugo\")
'    Debug.Assert (CombinePath("hoge\", "hugo") = "hoge\hugo")
'    Debug.Assert (CombinePath("hoge\", "hugo\") = "hoge\hugo\")
'    Debug.Assert (CombinePath("hoge", "hugo") = "hoge\hugo")
'    Debug.Assert (CombinePath("hoge", "\hugo") = "hoge\hugo")
'    Debug.Assert (CombinePath("\hoge", "hugo") = "\hoge\hugo")
'    Debug.Assert (CombinePath("\hoge", "\hugo") = "\hoge\hugo")
'    Debug.Assert (CombinePath("hoge", "hugo") = "hoge\hugo")
'    Debug.Assert (CombinePath("hoge", "C:\hugo") = "hoge\C:\hugo")
'    Debug.Assert (CombinePath("C:\hoge", "hugo") = "C:\hoge\hugo")
'    Debug.Assert (CombinePath("C:\hoge", "C:\hugo") = "C:\hoge\C:\hugo")
'    Debug.Assert (CombinePath("hoge", "hugo") = "hoge\hugo")
'    Debug.Assert (CombinePath("hoge", "\\hugo") = "hoge\\hugo")
'    Debug.Assert (CombinePath("\\hoge", "hugo") = "\\hoge\hugo")
'    Debug.Assert (CombinePath("\\hoge", "\\hugo") = "\\hoge\\hugo")
'End Sub

' ドライブ名取得
Function GetDriveName(ByVal sPath As String, Optional ByVal bFormal As Boolean = True) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    GetDriveName = oFSO.GetDriveName(sPath)
    ' UNCドライブ名は正確には「\\SERVER_NAME\DRIVE_NAME」の形式でなければならない
    ' が、ファイルエクスプローラ等は普通に「\\SERVER_NAME」だけでも反応するわけで
    ' こんなことを普通のユーザにいくら説明しても理解してもらえるワケがない
    ' オッサン多めの職場でイチイチ説明するのとかクッソ面倒なので制限を緩める方法を用意しておく
    If GetDriveName = "" And bFormal = False Then
        If Left(sPath, 2) = "\\" Then
            Dim lSepPos As Long
            lSepPos = InStr(3, sPath, "\")
            If lSepPos > 0 Then
                GetDriveName = Mid(sPath, 1, lSepPos - 1)
            Else
                GetDriveName = sPath
            End If
        End If
    End If
End Function

' 親パス取得
Function GetBasePath(ByVal sPath As String) As String
    GetBasePath = CreateFSO().GetParentFolderName(sPath)
End Function

' フォルダ名取得
Function GetBaseName(ByVal sPath As String) As String
    GetBaseName = CreateFSO().GetBaseName(sPath)
End Function

' ファイル名取得(拡張子あり)
Function GetFileName(ByVal sPath As String) As String
    GetFileName = CreateFSO().GetFileName(sPath)
End Function

' ファイル名取得(拡張子なし)
Function GetFileNameWithoutExtension(ByVal sPath As String) As String
    GetFileNameWithoutExtension = CreateFSO().GetBaseName(sPath)
End Function

' ファイル名変更
Function ChangeFileName(ByVal sPath As String, ByVal sFileName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    ChangeFileName = oFSO.BuildPath(oFSO.GetParentFolderName(sPath), sFileName & "." & LCase(oFSO.GetExtensionName(sPath)))
End Function

' 拡張子判定
Private Function isExtentionName(ByVal sPath As String, ByVal sExtensionNames As String) As Boolean
    Dim bRet As Boolean
    Dim elm As Variant
    For Each elm In Split(sExtensionNames, ",")
        If UCase(CreateFSO().GetExtensionName(sPath)) = UCase(elm) Then bRet = True
    Next
    isExtentionName = bRet
End Function

' 拡張子取得
Function GetExtensionName(ByVal sPath As String) As String
    GetExtensionName = LCase(CreateFSO().GetExtensionName(sPath))
End Function

' 拡張子変更
Function ChangeExtensionName(ByVal sPath As String, ByVal sExtensionName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    ChangeExtensionName = oFSO.BuildPath(oFSO.GetParentFolderName(sPath), oFSO.GetBaseName(sPath) & "." & LCase(sExtensionName))
End Function

' 相対パス取得
Function GetRelPath(ByVal sBasePath As String, ByVal sSpecPath As String) As String
    Dim sBuff As String
    sBuff = String$(260, vbNullChar)
    If PathRelativePathTo(StrPtr(sBuff), StrPtr(sBasePath), FILE_ATTRIBUTE_DIRECTORY, StrPtr(sSpecPath), 0) Then
        GetRelPath = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    Else
        GetRelPath = sSpecPath
    End If
End Function

' 絶対パス取得
Function GetAbsPath(ByVal sSpecPath As String, ByVal sBasePath As String) As String
    Dim sAbsPath As String
    If sSpecPath <> "" And sBasePath <> "" Then
        ' ベースパスは絶対パスを前提とする
        ' ※ベースパスが最初から絶対パスで指定されていれば特に影響のない念のための処理
        sBasePath = CreateFSO().GetAbsolutePathName(sBasePath) ' GetAbsolutePathNameのパス解決はカレントディレクトリ固定
        ' UNCの「\\」がパスセパレータ「\」と混ざって邪魔なので
        ' 内部的に都合の良い表現になるように前処理しておく
        Dim bSUNC As Boolean
        Dim bBUNC As Boolean
        sSpecPath = Replace(sSpecPath, "/", "\")
        sBasePath = Replace(sBasePath, "/", "\")
        If Left(sSpecPath, 2) = "\\" Then
            sSpecPath = Replace(sSpecPath, "\\", "<UNC>", 1, 1) ' <>はファイル名に使えない文字
            bSUNC = True
        End If
        If Left(sBasePath, 2) = "\\" Then
            sBasePath = Replace(sBasePath, "\\", "<UNC>", 1, 1) ' <>はファイル名に使えない文字
            bBUNC = True
        End If
        sSpecPath = PathTrim(sSpecPath, "\")
        sBasePath = PathTrim(sBasePath, "\")
        sSpecPath = IIf(bSUNC = False And InStr(sSpecPath, "\") = 0, ".\", "") & sSpecPath
        sBasePath = IIf(bBUNC = False And InStr(sBasePath, "\") = 0, ".\", "") & sBasePath
        ' パスセパレータでバラして「.」「..」に従ったパス編集操作を実施
        Dim elm As Variant
        For Each elm In Split(sSpecPath, "\")
            Select Case elm
                Case ".", ""
                    sAbsPath = IIf(sAbsPath = "", sBasePath, sAbsPath)
                Case ".."
                    sAbsPath = IIf(sAbsPath = "", sBasePath, sAbsPath)
                    If InStrRev(sAbsPath, "\") > 0 Then
                        sAbsPath = Left(sAbsPath, InStrRev(sAbsPath, "\") - 1)
                    End If
                Case Else
                    sAbsPath = sAbsPath & IIf(sAbsPath = "", "", "\") & elm
            End Select
        Next
        ' パスセパレータの連続を削除
        Dim lAbsPath As Long
        Do
            lAbsPath = Len(sAbsPath)
            sAbsPath = Replace(sAbsPath, "\\", "\")
        Loop While lAbsPath > Len(sAbsPath)
        ' UNC接頭辞を元に戻す
        sAbsPath = Replace(sAbsPath, "<UNC>", "\\")
    End If
    GetAbsPath = sAbsPath
End Function

'Sub TestGetAbsPath()
'    Debug.Assert (GetAbsPath("", "") = "")
'    Debug.Assert (GetAbsPath("", "C:\hoge") = "")
'    Debug.Assert (GetAbsPath(".", "C:\hoge") = "C:\hoge")
'    Debug.Assert (GetAbsPath(".", "C:\hoge\") = "C:\hoge")
'    Debug.Assert (GetAbsPath(".\", "C:\hoge") = "C:\hoge")
'    Debug.Assert (GetAbsPath(".\", "C:\hoge\") = "C:\hoge")
'    Debug.Assert (GetAbsPath("C:\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("C:\.\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("C:\.\.\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("C:\piyo\..\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("C:\piyo\..\..\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("poco", "C:\hoge") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath("poco", "C:\hoge\") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath("\poco", "C:\hoge") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath("\poco", "C:\hoge\") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath(".\poco", "C:\hoge") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath(".\poco", "C:\hoge\") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath(".\poco", "C:\hoge\fuga") = "C:\hoge\fuga\poco")
'    Debug.Assert (GetAbsPath(".\.\poco", "C:\hoge\fuga") = "C:\hoge\fuga\poco")
'    Debug.Assert (GetAbsPath("..\poco", "C:\hoge\fuga") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath("..\..\poco", "C:\hoge\fuga") = "C:\poco")
'    Debug.Assert (GetAbsPath("..\..\..\poco", "C:\hoge\fuga") = "C:\poco")
'    Debug.Assert (GetAbsPath(".\..\poco", "C:\hoge\fuga") = "C:\hoge\poco")
'    Debug.Assert (GetAbsPath(".\..\..\poco", "C:\hoge\fuga") = "C:\poco")
'    Debug.Assert (GetAbsPath(".\..\..\..\poco", "C:\hoge\fuga") = "C:\poco")
'    Debug.Assert (GetAbsPath(".\..\piyo\poco", "C:\hoge\fuga") = "C:\hoge\piyo\poco")
'    Debug.Assert (GetAbsPath(".\..\..\piyo\poco", "C:\hoge\fuga") = "C:\piyo\poco")
'    Debug.Assert (GetAbsPath(".\..\..\..\piyo\poco", "C:\hoge\fuga") = "C:\piyo\poco")
'    Debug.Assert (GetAbsPath("..\.\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("..\.\.\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("..\.\.\..\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath(".\..\..\.\poco", "C:\hoge") = "C:\poco")
'    Debug.Assert (GetAbsPath("piyo", "\\hoge\fuga\") = "\\hoge\fuga\piyo")
'    Debug.Assert (GetAbsPath(".\piyo", "\\hoge\fuga\") = "\\hoge\fuga\piyo")
'    Debug.Assert (GetAbsPath("..\piyo", "\\hoge\fuga\") = "\\hoge\piyo")
'    Debug.Assert (GetAbsPath("..\..\piyo", "\\hoge\fuga\") = "\\hoge\piyo")
'    Debug.Assert (GetAbsPath("C:\piyo", "\\hoge\fuga\") = "C:\piyo")
'    Debug.Assert (GetAbsPath("C:\.\piyo", "\\hoge\fuga\") = "C:\piyo")
'    Debug.Assert (GetAbsPath("C:\..\piyo", "\\hoge\fuga\") = "C:\piyo")
'    Debug.Assert (GetAbsPath("C:\..\..\piyo", "\\hoge\fuga\") = "C:\piyo")
'    Debug.Assert (GetAbsPath("\\hoge", "C:\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge\.", "C:\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge\..", "C:\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge\..\..", "C:\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge", "\\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge\.", "\\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge\..", "\\piyo\pico") = "\\hoge")
'    Debug.Assert (GetAbsPath("\\hoge\..\..", "\\piyo\pico") = "\\hoge")
'End Sub

' パスのサニタイズ
Function SanitizePath(ByVal sText As String) As String
    sText = Trim(sText)
    sText = Replace(sText, vbCr, "")
    sText = Replace(sText, vbLf, "")
    sText = Replace(sText, "/", "\")
    ' ドライブ名部分と、フォルダorファイル名部分を分離
    Dim sDPart As String
    Dim sPPart As String
    If sText <> "" Then
        sDPart = GetDriveName(sText, False)
        sPPart = Mid(sText, InStr(sText, sDPart) + Len(sDPart))
    End If
    ' フォルダorファイル名部分の検証
    Dim sPath As Variant
    Dim sPart As Variant
    For Each sPart In Split(sPPart, "\")
        sPath = CombinePath(sPath, SanitizeFileName(CStr(sPart)))
    Next
    ' 結合
    Dim sRet As String
    If sDPart <> "" Then
        sRet = CombinePath(sDPart & "\", sPath)
    Else
        sRet = sPath
    End If
    SanitizePath = sRet
End Function

' ファイル名のサニタイズ
Function SanitizeFileName(ByVal sText As String) As String
    Dim sRet As String
    sRet = sText
    ' 改行コード削除
    sRet = Replace(sRet, vbCr, "")
    sRet = Replace(sRet, vbLf, "")
    ' パスに使えない文字を発見したら全角に修正する
    ' 間違ってドライブ名とかに使うと「:」が全角になるので注意
    sRet = Replace(sRet, "\", "￥")
    sRet = Replace(sRet, ":", "：")
    sRet = Replace(sRet, "/", "／")
    sRet = Replace(sRet, "*", "＊")
    sRet = Replace(sRet, "?", "？")
    sRet = Replace(sRet, """", Chr(&H8168))
    sRet = Replace(sRet, "<", "＜")
    sRet = Replace(sRet, ">", "＞")
    sRet = Replace(sRet, "|", "｜")
    ' 予約デバイス名に一致するようなら置き換える
    sRet = IIf(sRet = "AUX", "_AUX", sRet)
    sRet = IIf(sRet = "CON", "_CON", sRet)
    sRet = IIf(sRet = "NUL", "_NUL", sRet)
    sRet = IIf(sRet = "PRN", "_PRN", sRet)
    sRet = IIf(sRet = "CLOCK$", "_CLOCK$", sRet)
    sRet = IIf(sRet Like "COM[0-9]", Replace(sRet, "COM", "_COM"), sRet)
    sRet = IIf(sRet Like "LPT[0-9]", Replace(sRet, "LPT", "_LPT"), sRet)
    SanitizeFileName = sRet
End Function

'Sub TestSanitizePath()
'    Debug.Assert (SanitizePath("C:") = "C:\")                               ' 絶対パス:C:\
'    Debug.Assert (SanitizePath("C:\") = "C:\")                              ' 絶対パス:C:\
'    Debug.Assert (SanitizePath("C:\USER\") = "C:\USER")                     ' 絶対パス:C:\USER
'    Debug.Assert (SanitizePath("AAA") = "AAA")                              ' 相対パス:AAA
'    Debug.Assert (SanitizePath(".\AAA") = ".\AAA")                          ' 相対パス:.\AAA
'    Debug.Assert (SanitizePath("..\AAA") = "..\AAA")                        ' 相対パス:..\AAA
'    Debug.Assert (SanitizePath("\\NAS0001") = "\\NAS0001\")                 ' UNCパス :\\NAS0001\           ※WindowsのUNCパスとして本来はNGだけどここではUNC扱いしておく
'    Debug.Assert (SanitizePath("\\NAS0001\D") = "\\NAS0001\D\")             ' UNCパス :\\NAS0001\D\
'    Debug.Assert (SanitizePath("\\NAS0001\D\USER") = "\\NAS0001\D\USER")    ' UNCパス :\\NAS0001\D\USER
'End Sub

' システムドライブパス
Function GetSystemDrivePath() As String
    GetSystemDrivePath = Environ("SystemDrive")
End Function

' デスクトップパス
Function GetDesktopPath() As String
    GetDesktopPath = CreateWSH().SpecialFolders("Desktop")
End Function

' マイドキュメントパス
Function GetMyDocumentsPath() As String
    GetMyDocumentsPath = CreateWSH().SpecialFolders("MyDocuments")
End Function

' ダウンロードパス
Function GetDownloadPath() As String
    GetDownloadPath = CreateShellApp().Namespace("shell:Downloads").Self.Path
End Function

' ユーザプロファイルパス
Function GetUserProfilePath() As String
    GetUserProfilePath = Environ("USERPROFILE")
End Function

' アプリケーションデータパス
Function GetAppDataPath() As String
    GetAppDataPath = Environ("APPDATA")
End Function

' テンポラリパス取得
Function GetTempPath(Optional ByVal sBasePath As String = "", Optional ByVal bFile As Boolean = True) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If sBasePath = "" Then
        sBasePath = oFSO.GetSpecialFolder(2)
    End If
    Dim sPath As String
    If bFile Then
        sPath = oFSO.BuildPath(sBasePath, oFSO.GetTempName())
    Else
        sPath = oFSO.BuildPath(sBasePath, oFSO.GetBaseName(oFSO.GetTempName()))
    End If
    GetTempPath = sPath
End Function

' ブックパス
Function BookPath() As String
    BookPath = ThisWorkbook.Path
End Function

' ハイパーリンクの基点
Function HyperlinkBasePath(ByVal oBook As Workbook) As String
    Dim sRet As String
    sRet = oBook.BuiltinDocumentProperties("Hyperlink base").Value
    If sRet = "" Then
        sRet = oBook.Path
    End If
    HyperlinkBasePath = sRet
End Function

' ドメイン名称
Function DomainName() As String
    DomainName = Environ("USERDOMAIN")
End Function

' コンピュータ名称
Function ComputerName() As String
    ComputerName = Environ("COMPUTERNAME")
End Function

' オペレーティングシステム名称
Function OperatingSystemName() As String
    With CreateObject("WbemScripting.SWbemLocator")
        Dim elm As Object
        For Each elm In .ConnectServer.ExecQuery("Select * From Win32_OperatingSystem")
            OperatingSystemName = elm.Caption & " (" & elm.OSArchitecture & ") Version " & elm.Version
        Next
    End With
End Function

' プロセッサアーキテクチャ名称
Function ArchitectureName() As String
    ArchitectureName = CreateWSH().Environment("Process").Item("PROCESSOR_ARCHITECTURE")
End Function

' OSユーザ名
Function UserName() As String
    UserName = Environ("USERNAME")
End Function

' アプリケーションユーザ名
Function AppUserName(ByVal sDefaultName As String, Optional ByVal bFamilyNameOnly As Boolean = False)
    Dim sText As String
    sText = Application.UserName
    If bFamilyNameOnly Then
        sText = Replace(sText, "　", " ")   ' 全角空白
        sText = Trim(sText)                 ' 前後空白
        If sText <> "" Then
            sText = Split(sText, " ")(0)
        End If
    End If
    If sText = "" Then
        sText = sDefaultName
    End If
    AppUserName = sText
End Function

' 指定フォルダ配下のファイル列挙
Function EnumFile(ByVal sPath As String, ByVal sPtrn As String, Optional ByVal lRecursive As Long = 0) As Collection
    Dim oRet As New Collection  ' 戻り値にすると再帰時に重いが、使い勝手優先で
On Error GoTo ErrExit
    ' FSOだとリスト取得前にフィルタ出来なくて重いのでAPIを使う方法にする
    Dim oFind As WIN32_FIND_DATA
    Dim hFind As LongPtr
    hFind = FindFirstFileEx(StrPtr(sPath & "\" & sPtrn), FindExInfoBasic, oFind, FindExSearchNameMatch, 0&, FIND_FIRST_EX_LARGE_FETCH)
    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            If (oFind.dwFileAttributes And vbDirectory) <> vbDirectory Then
                Dim sEntry As String
                sEntry = Trim$(Left$(oFind.cFileName, InStr(oFind.cFileName, vbNullChar) - 1))
                Call oRet.Add(sPath & "\" & sEntry)
            End If
        Loop Until FindNextFile(hFind, oFind) = 0
        If lRecursive <> 0 Then
            lRecursive = lRecursive - 1
            Dim sSubDElm As Variant
            Dim sSubFElm As Variant
            For Each sSubDElm In EnumFolder(sPath, "*", 0) ' 再帰の場合、子フォルダは名前でフィルタせず全部処理対象にする
                For Each sSubFElm In EnumFile(sSubDElm, sPtrn, lRecursive)
                    Call oRet.Add(sSubFElm)
                Next
            Next
        End If
    End If
ErrExit:
    If hFind <> INVALID_HANDLE_VALUE Then Call FindClose(hFind)
    Set EnumFile = oRet
End Function

' 指定フォルダ配下のフォルダ列挙
' ・sPtrnは目的の階層が単体の場合にのみ指定する事
Function EnumFolder(ByVal sPath As String, ByVal sPtrn As String, Optional ByVal lRecursive As Long = 0) As Collection
    Dim oRet As New Collection  ' 戻り値にすると再帰時に重いが、使い勝手優先で
On Error GoTo ErrExit
    ' FSOだとリスト取得前にフィルタ出来なくて重いのでAPIを使う方法にする
    Dim oFind As WIN32_FIND_DATA
    Dim hFind As LongPtr
    hFind = FindFirstFileEx(StrPtr(sPath & "\" & sPtrn), FindExInfoBasic, oFind, FindExSearchNameMatch, 0&, FIND_FIRST_EX_LARGE_FETCH)
    If hFind <> INVALID_HANDLE_VALUE Then
        Do
            If (oFind.dwFileAttributes And vbDirectory) = vbDirectory Then
                Dim sEntry As String
                sEntry = Trim$(Left$(oFind.cFileName, InStr(oFind.cFileName, vbNullChar) - 1))
                If sEntry <> "." And sEntry <> ".." Then
                    Call oRet.Add(sPath & "\" & sEntry)
                    If lRecursive <> 0 Then
                        lRecursive = lRecursive - 1
                        Dim sSubElm As Variant
                        For Each sSubElm In EnumFolder(sPath & "\" & sEntry, sPtrn, lRecursive)
                            Call oRet.Add(sSubElm)
                        Next
                    End If
                End If
            End If
        Loop Until FindNextFile(hFind, oFind) = 0
    End If
ErrExit:
    If hFind <> INVALID_HANDLE_VALUE Then Call FindClose(hFind)
    Set EnumFolder = oRet
End Function

' 自然順(論理順)ソート
Public Function SortCollectionLogical(ByRef oCollection As Collection) As Collection
    Dim oRet As New Collection
    
    Dim i As Long
    Dim j As Long

    ' Collection→配列
    Dim vArr() As Variant
    ReDim vArr(1 To oCollection.Count)
    For i = 1 To oCollection.Count
        vArr(i) = oCollection(i)
    Next
    
    ' 自然順ソート
    Dim tmp As Variant
    For i = LBound(vArr) To UBound(vArr)
        For j = i To UBound(vArr)
            If StrCmpLogicalW(StrPtr(vArr(i)), StrPtr(vArr(j))) > 0 Then
                Let tmp = vArr(i)
                Let vArr(i) = vArr(j)
                Let vArr(j) = tmp
            End If
       Next
    Next

    ' 配列→Collection
    For i = 1 To UBound(vArr)
        Call oRet.Add(vArr(i))
    Next
    
    Set SortCollectionLogical = oRet
End Function

' 指定したパス＋パターンでパスを補完する
Function PathCompleteFile(ByVal sPath As String, ByVal sPtrn As String, Optional lFlug As Long = vbNormal Or vbHidden) As String
    Dim sFind As String
    sFind = CombinePath(sPath, sPtrn)
    Dim sName As String
    Do
        If sName <> "" And sName <> "." And sName <> ".." Then
            If Not GetAttr(CombinePath(sPath, sName)) And vbDirectory Then
                sName = CombinePath(sPath, sName)
                Exit Do
            End If
        End If
        sName = Dir(sFind, lFlug Or vbNormal)
    Loop Until sName = ""
    PathCompleteFile = sName
End Function

' 指定したパス＋パターンでパスを補完する
Function PathCompleteDir(ByVal sPath As String, ByVal sPtrn As String, Optional lFlug As Long = vbDirectory Or vbHidden) As String
    Dim sFind As String
    sFind = CombinePath(sPath, sPtrn)
    Dim sName As String
    Do
        If sName <> "" And sName <> "." And sName <> ".." Then
            If GetAttr(CombinePath(sPath, sName)) And vbDirectory Then
                sName = CombinePath(sPath, sName)
                Exit Do
            End If
        End If
        sName = Dir(sFind, lFlug Or vbDirectory)
    Loop Until sName = ""
    PathCompleteDir = sName
End Function

' 指定パスで「ある名称に最も似た」名称のファイルを探す
Function ChooseSimilarFile(ByVal sPath As String, ByVal sName As String, ByVal sPtrn As String, ByVal dThr As Double) As String
    Dim sRet As String
    Dim dVal As Double
    Dim dMax As Double
    Dim oFile
    For Each oFile In CreateFSO().GetFolder(sPath).Files
        If oFile.Path Like sPtrn Then
            dVal = LevenshteinRatio(UCase(sName), UCase(CreateFSO().GetBaseName(oFile.Path)))
            If dMax > dVal And dVal > dThr Then
                dMax = dVal
                sRet = oFile.Path
            End If
        End If
    Next
    ChooseSimilarFile = sRet
End Function

' 指定パスで「ある名称に最も似た」名称のフォルダを探す
Function ChooseSimilarFolder(ByVal sPath As String, ByVal sName As String, ByVal sPtrn As String, ByVal dThr As Double) As String
    Dim sRet As String
    Dim dVal As Double
    Dim dMax As Double
    Dim oFolder
    For Each oFolder In CreateFSO().GetFolder(sPath).Folders
        If oFolder.Path Like sPtrn Then
            dVal = LevenshteinRatio(UCase(sName), UCase(CreateFSO().GetBaseName(oFolder.Path)))
            If dMax > dVal And dVal > dThr Then
                dMax = dVal
                sRet = oFolder.Path
            End If
        End If
    Next
    ChooseSimilarFolder = sRet
End Function

' レーベンシュタイン距離比率
Private Function LevenshteinRatio(ByVal sLhs As String, ByVal sRhs As String) As Double
    LevenshteinRatio = 1# - CDbl(LevenshteinDistance(sLhs, sRhs) / WorksheetFunction.Max(Len(sLhs), Len(sRhs)))
End Function

' レーベンシュタイン距離(≒文字列の類似度計算)
Private Function LevenshteinDistance(ByVal sLhs As String, ByVal sRhs As String) As Long
    Dim lLhs As Long
    Dim lRhs As Long
    Dim aLhs() As String
    Dim aRhs() As String
    lLhs = Len(sLhs)
    lRhs = Len(sRhs)
    ReDim d(lLhs, lRhs)
    ReDim aLhs(lLhs)
    ReDim aRhs(lRhs)

    Dim n As Long
    For n = 1 To lLhs
        aLhs(n - 1) = Mid(sLhs, n, 1)
    Next
    For n = 1 To lRhs
        aRhs(n - 1) = Mid(sRhs, n, 1)
    Next

    Dim i As Long
    Dim j As Long
    For i = 0 To lLhs
        d(i, 0) = i
        For j = 1 To lRhs
            d(i, j) = 0
        Next
    Next
    For j = 0 To lRhs
        d(0, j) = j
    Next

    Dim cost As Long
    For i = 1 To lLhs
        For j = 1 To lRhs
            If aLhs(i - 1) = aRhs(j - 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = WorksheetFunction.Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next
    Next

    LevenshteinDistance = d(lLhs, lRhs)
End Function

' ファイル選択ダイアログ
Function DialogOpenFileName( _
    Optional ByVal sTitle As String = "開く", _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFileFilter As String = "全て,*.*", _
    Optional ByVal bMultiSelect As Boolean = False _
) As Collection
    Dim oRet As New Collection ' 単一選択のつもりでもループで書いとけと言っている
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = sTitle
        .InitialFileName = sInitPath & IIf(Right(sInitPath, 1) <> "\", "\", "")
        .Filters.Clear
        ' sFileFilterはこんな感じに指定する:"Excel,*.xls*,その他,*.txt;*.csv"
        Dim oPart As Variant
        oPart = Split(sFileFilter, ",")
        Dim i As Long
        For i = LBound(oPart) To UBound(oPart) Step 2
            Call .Filters.Add(oPart(i + 0), oPart(i + 1))
        Next
        .AllowMultiSelect = bMultiSelect
        If .Show = True Then
            Dim sPath As Variant
            For Each sPath In .SelectedItems
                Call oRet.Add(sPath)
            Next
        End If
    End With
    Set DialogOpenFileName = oRet
End Function

' フォルダ選択ダイアログ
Function DialogOpenFolderName( _
    Optional ByVal sTitle As String = "開く", _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal bMultiSelect As Boolean = False _
) As Collection
    Dim oRet As New Collection ' 単一選択のつもりでもループで書いとけと言っている
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    Do While (sInitPath <> "" And oFSO.FolderExists(sInitPath) = False)
        sInitPath = oFSO.GetParentFolderName(sInitPath)
    Loop
    sInitPath = sInitPath & IIf(Right(sInitPath, 1) <> "\", "\", "")
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = sTitle
        .InitialFileName = sInitPath
        .AllowMultiSelect = bMultiSelect
        If .Show = True Then
            Dim sPath As Variant
            For Each sPath In .SelectedItems
                Call oRet.Add(sPath)
            Next
        End If
    End With
    Set DialogOpenFolderName = oRet
End Function

' // ファイル操作 /////////////////////////////////////////

' ファイル作成
Sub MakeFile(ByVal sPath As String)
    With CreateFSO().OpenTextFile(sPath, 8, True)
        .Close
    End With
End Sub

' ファイル存在判定
Function IsExistFile(ByVal sPath As String)
    IsExistFile = CreateFSO().FileExists(sPath)
End Function

' ファイル削除
Sub RemoveFile(ByVal sPath As String, Optional ByVal bForce As Boolean = False)
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) = True Then
        Call oFSO.DeleteFile(sPath, bForce)
    End If
End Sub

' ファイル複製
Sub CopyFile(ByVal sSrcPath As String, ByVal sDstPath As String, Optional bReplace = False)
    If sSrcPath <> sDstPath Then
        Call CreateFSO().CopyFile(sSrcPath, sDstPath, bReplace)
    End If
End Sub

' ファイル移動
Sub MoveFile(ByVal sSrcPath As String, ByVal sDstPath As String, Optional bReplace = False)
    If sSrcPath <> sDstPath Then
        Dim oFSO As Object
        Set oFSO = CreateFSO()
        If (bReplace = True) And (oFSO.FileExists(sDstPath) = True) Then
            Call oFSO.DeleteFile(sDstPath, True)
        End If
        Call oFSO.MoveFile(sSrcPath, sDstPath)
    End If
End Sub

' フォルダ作成
Sub MakeDir(ByVal sPath As String)
    If CreateFSO().FolderExists(sPath) = False Then
        ' APIで存在しない中間フォルダも一気に作成させる
        Call SHCreateDirectoryEx(0&, StrPtr(sPath), 0&)
    End If
End Sub

' フォルダ存在判定
Function IsExistDir(ByVal sPath As String)
    IsExistDir = CreateFSO().FolderExists(sPath)
End Function

' フォルダ削除
Sub RemoveDir(ByVal sPath As String, Optional ByVal bForce As Boolean = False)
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FolderExists(sPath) = True Then
        Call oFSO.DeleteFolder(sPath, bForce)
    End If
End Sub

' フォルダ複製
Sub CopyDir(ByVal sSrcPath As String, ByVal sDstPath As String, Optional ByVal bReplace As Boolean = False)
    If sSrcPath <> sDstPath Then
        Call CreateFSO().CopyFolder(sSrcPath, sDstPath, bReplace)
    End If
End Sub

' フォルダ移動
Sub MoveDir(ByVal sSrcPath As String, ByVal sDstPath As String, Optional ByVal bReplace As Boolean = False)
    If sSrcPath <> sDstPath Then
        Dim oFSO As Object
        Set oFSO = CreateFSO()
        If (bReplace = True) And (oFSO.FolderExists(sDstPath) = True) Then
            Call oFSO.DeleteFolder(sDstPath, True)
        End If
        Call oFSO.MoveFolder(sSrcPath, sDstPath)
    End If
End Sub

' // ファイル日付 /////////////////////////////////////////

' 作成日時取得
Function GetFileDateCreated(ByVal sPath As String) As Date
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) Then
        GetFileDateCreated = oFSO.GetFile(sPath).DateCreated
    Else
        GetFileDateCreated = vbEmpty
    End If
End Function

' 最終更新日時取得
Function GetFileDateLastModified(ByVal sPath As String) As Date
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) Then
        GetFileDateLastModified = oFSO.GetFile(sPath).DateLastModified
    Else
        GetFileDateLastModified = vbEmpty
    End If
End Function

' 最終更新日時設定
Sub SetFileDateLastModified(ByVal sPath As String, ByVal oDate As Date)
    Dim oFSO As Object
    Dim oSHA As Object
    Set oFSO = CreateFSO()
    Set oSHA = CreateShellApp()
    If oFSO.FileExists(sPath) Then
        sPath = oFSO.GetAbsolutePathName(sPath)
        Dim sFPath As Variant ' ※何故かStringではダメ
        Dim sFName As Variant ' ※何故かStringではダメ
        sFPath = oFSO.GetParentFolderName(sPath)
        sFName = oFSO.GetFileName(sPath)
        Dim oFldr As Object
        Set oFldr = oSHA.Namespace(sFPath)
        Dim oFile As Object
        Set oFile = oFldr.ParseName(sFName)
        oFile.ModifyDate = oDate
    End If
End Sub

' アクセス日時取得
Function GetFileDateLastAccessed(ByVal sPath As String) As Date
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) Then
        GetFileDateLastAccessed = oFSO.GetFile(sPath).DateLastAccessed
    Else
        GetFileDateLastAccessed = vbEmpty
    End If
End Function

' // INI //////////////////////////////////////////////////

' テキスト書き込み
Public Sub WriteIniText(ByVal sPath As String, ByVal sSection As String, ByVal sKey As String, ByVal sVal As String)
    Call WritePrivateProfileString(sSection, sKey, sVal, sPath)
End Sub

' テキスト読み出し
Public Function ReadIniText(ByVal sPath As String, ByVal sSection As String, ByVal sKey As String, Optional ByVal sDef As String = "") As String
    Dim sBuf As String * 1024
    Call GetPrivateProfileString(sSection, sKey, sDef, sBuf, Len(sBuf), sPath)
    ReadIniText = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
End Function

' // 外部プロセス操作 /////////////////////////////////////

' パスを開く
Sub OpenPath(ByVal sPath As String)
    ' ディレクトリならエクスプローラを開く
    ' ファイルなら関連付けに従って開く
    Call CreateShellApp().ShellExecute(sPath)
End Sub

' パスを開く(拡張)
Sub OpenPathEx(ByVal sPath As String, ByVal sFilter As String)
    ' エクスプローラのフィルタを指定して開く
    ' ファイルスキーマ指定の利用でもっといろいろ出来るけど常用できそうなのはこんなもんしかない
    Call CreateShellApp().ShellExecute("search-ms:query=" & sFilter & "&" & "crumb=location:" & sPath)
End Sub

' コマンド存在確認
Function IsExistCmd(ByVal sCommand As String) As Boolean
    Dim bRet As Boolean
    bRet = CreateWSH().Run("%ComSpec% /c where " & sCommand & " > nul 2>&1", 0, True)
    IsExistCmd = Not bRet
End Function

' コマンド同期実行(標準出力無)
Function RunCmd(ByVal sCommand As String) As Long
    Dim lResult As Long
On Error GoTo ErrExit
    lResult = CreateWSH().Run("%ComSpec% /c " & PathDQuote(sCommand), 0, True)
ErrExit:
    RunCmd = lResult
End Function

' コマンド同期実行(標準出力有)
Function RunCmdEx(ByVal sCommand As String, Optional ByVal sEncode As String = "UTF-8") As String
    Dim sResult As String
On Error GoTo ErrExit
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    Dim sTempPath As String
    sTempPath = oFSO.BuildPath(oFSO.GetSpecialFolder(2), oFSO.GetTempName())
    Call RunCmd(sCommand & " > " & PathDQuote(sTempPath))
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2
        .Charset = sEncode
        .LoadFromFile sTempPath
        Do Until .EOS
            sResult = .ReadText
        Loop
        .Close
    End With
ErrExit:
    If oFSO.FileExists(sTempPath) Then Call oFSO.DeleteFile(sTempPath, True)
    RunCmdEx = sResult
End Function

' PowerShell同期実行(標準出力無)
Function RunPsh(ByVal sCommand As String) As Long
    Dim lResult As Long
On Error GoTo ErrExit
    lResult = CreateWSH().Run("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & PathDQuote(sCommand), 0, True)
ErrExit:
    RunPsh = lResult
End Function

' PowerShell同期実行(標準出力有)
Function RunPshEx(ByVal sCommand As String, Optional ByVal sEncode As String = "SJIS") As String
    Dim sResult As String
On Error GoTo ErrExit
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    Dim sTempPath As String
    sTempPath = oFSO.BuildPath(oFSO.GetSpecialFolder(2), oFSO.GetTempName())
    Call RunPsh(sCommand & " | Out-File -filePath " & PathDQuote(sTempPath) & " -encoding Default")
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2
        .Charset = sEncode
        .LoadFromFile sTempPath
        Do Until .EOS
            sResult = .ReadText
        Loop
        .Close
    End With
ErrExit:
    If oFSO.FileExists(sTempPath) Then Call oFSO.DeleteFile(sTempPath, True)
    RunPshEx = sResult
End Function

' ダウンロード
Function DownloadFile(ByVal sDLSrc As String, ByVal sDLDst As String) As Long
    Dim lRet As Long: lRet = -1
    Dim pHTTP As Object
    Set pHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call pHTTP.Open("GET", sDLSrc, False)
    Call pHTTP.Send
    If pHTTP.Status = 200 Then
        Dim pStrm As Object
        Set pStrm = CreateObject("ADODB.Stream")
        pStrm.Type = 1
        Call pStrm.Open
        Call pStrm.Write(pHTTP.responseBody)
        Call pStrm.SaveToFile(sDLDst, 2)
        Call pStrm.Close
        lRet = 0
    End If
    DownloadFile = lRet
End Function

' 圧縮
Function CompressArchive(ByVal sSrcPath As String, ByVal sDstPath As String) As Long
    CompressArchive = RunPsh("Compress-Archive -Path " & sSrcPath & " -DestinationPath " & sDstPath & " -Force")
End Function

' 解凍
Function ExtractArchive(ByVal sSrcPath As String, ByVal sDstPath As String) As Long
    ExtractArchive = RunPsh("Expand-Archive -Path " & sSrcPath & " -DestinationPath " & sDstPath & " -Force")
End Function
