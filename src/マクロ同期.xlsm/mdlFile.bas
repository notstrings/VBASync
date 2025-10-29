Attribute VB_Name = "mdlFile"
Option Explicit
Option Private Module

' // //////////////////////////////////////////////////////////////////////////
' // �N�ł���肻���ȃt�@�C���E�p�X�E�V�F������
' // 20210901:����
' // 202208xx:INI�t�@�C������ǉ�
' // 20221101:�ގ��t�@�C���T���ǉ�
' // 20240401:ChooseSimilarFile/ChooseSimilarFolder�C��
' //          ByVal/ByRef��߂�l�̌^�w���O��
' // 20251010:�R�}���h���݃`�F�b�N�ǉ�

' CreateDirectory�p
Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExW" (ByVal hWnd As Long, ByVal pszPath As LongPtr, ByVal psa As Long) As Long

' PathRelativePathTo�p
Private Declare PtrSafe Function PathRelativePathTo Lib "Shlwapi" Alias "PathRelativePathToW" (ByVal pszPath As LongPtr, ByVal pszFrom As LongPtr, ByVal dwAttrFrom As Long, ByVal pszTo As LongPtr, ByVal dwAttrTo As Long) As Long
Private Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10
Private Const FILE_ATTRIBUTE_NORMAL  As Integer = &H80

' FindFirst*�p
Private Declare PtrSafe Function FindFirstFileEx Lib "kernel32" Alias "FindFirstFileExW" (ByVal lpFileName As LongPtr, ByVal fInfoLevelId As FINDEX_INFO_LEVELS, lpFindFileData As WIN32_FIND_DATA, ByVal fSearchOp As FINDEX_SEARCH_OPS, ByVal lpSearchFilter As LongPtr, ByVal dwAdditionalFlags As Long) As LongPtr
Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As LongPtr
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
    dwFileAttributes                        As Long     ' �t�@�C������
    ftCreationTime                          As FileTime ' �쐬��
    ftLastAccessTime                        As FileTime ' �ŏI�A�N�Z�X��
    ftLastWriteTime                         As FileTime ' �ŏI�X�V��
    nFileSizeHigh                           As Long     ' �t�@�C���T�C�Y�i��ʂR�Q�r�b�g�j
    nFileSizeLow                            As Long     ' �t�@�C���T�C�Y�i���ʂR�Q�r�b�g�j
    dwReserved0                             As Long     ' �\��ς݁B���p�[�X�^�O
    dwReserved1                             As Long     ' �\��ς݁B���g�p
    cFileName(260 * 2 - 1)                  As Byte     ' �t�@�C����
    cAlternateFileName(14 * 2 - 1)          As Byte     ' 8.3�`���̃t�@�C����
End Type

' URLDownloadToFile�p
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

' Ini�p
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String _
) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String _
) As Long

' // ���p /////////////////////////////////////////////////

' FileSystemObject����
Private Function CreateFSO() As Object
    Static oFSO As Object
    If oFSO Is Nothing Then
        Set oFSO = CreateObject("Scripting.FileSystemObject")
    End If
    Set CreateFSO = oFSO
End Function

' WScript.Shell����
Private Function CreateWSH() As Object
    Static oWSH As Object
    If oWSH Is Nothing Then
        Set oWSH = CreateObject("WScript.Shell")
    End If
    Set CreateWSH = oWSH
End Function

' Shell.Application����
Private Function CreateShellApp() As Object
    Static oSH As Object
    If oSH Is Nothing Then
        Set oSH = CreateObject("Shell.Application")
    End If
    Set CreateShellApp = oSH
End Function

' �g����(���[)
Private Function PathTrimL(ByVal sText As String, ByVal sChar As String) As String
    While Left(sText, 1) Like sChar
        sText = Right(sText, Len(sText) - 1)
    Wend
    PathTrimL = sText
End Function

' �g����(�E�[)
Private Function PathTrimR(ByVal sText As String, ByVal sChar As String) As String
    While Right(sText, 1) = sChar
        sText = Left(sText, Len(sText) - 1)
    Wend
    PathTrimR = sText
End Function

' �g����(���[)
Private Function PathTrim(ByVal sText As String, ByVal sChar As String) As String
    sText = PathTrimL(sText, sChar)
    sText = PathTrimR(sText, sChar)
    PathTrim = sText
End Function

' �V���O���N�H�[�g
Private Function PathSQuote(ByVal sText As String) As String
    PathSQuote = "'" & sText & "'"
End Function

' �_�u���N�H�[�g
Private Function PathDQuote(ByVal sText As String) As String
    PathDQuote = """" & sText & """"
End Function

' // �p�X���� /////////////////////////////////////////////

' �p�X�A��
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

' �h���C�u���擾
Function GetDriveName(ByVal sPath As String, Optional ByVal bFormal As Boolean = True) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    GetDriveName = oFSO.GetDriveName(sPath)
    ' UNC�h���C�u���͐��m�ɂ́u\\SERVER_NAME\DRIVE_NAME�v�̌`���łȂ���΂Ȃ�Ȃ�
    ' ���A�t�@�C���G�N�X�v���[�����͕��ʂɁu\\SERVER_NAME�v�����ł���������킯��
    ' ����Ȃ��Ƃ𕁒ʂ̃��[�U�ɂ�����������Ă��������Ă��炦�郏�P���Ȃ�
    ' �I�b�T�����߂̐E��ŃC�`�C�`��������̂Ƃ��N�b�\�ʓ|�Ȃ̂Ő������ɂ߂���@��p�ӂ��Ă���
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

' �e�p�X�擾
Function GetBasePath(ByVal sPath As String) As String
    GetBasePath = CreateFSO().GetParentFolderName(sPath)
End Function

' �t�H���_���擾
Function GetBaseName(ByVal sPath As String) As String
    GetBaseName = CreateFSO().GetBaseName(sPath)
End Function

' �t�@�C�����擾(�g���q����)
Function GetFileName(ByVal sPath As String) As String
    GetFileName = CreateFSO().GetFileName(sPath)
End Function

' �t�@�C�����擾(�g���q�Ȃ�)
Function GetFileNameWithoutExtension(ByVal sPath As String) As String
    GetFileNameWithoutExtension = CreateFSO().GetBaseName(sPath)
End Function

' �t�@�C�����ύX
Function ChangeFileName(ByVal sPath As String, ByVal sFileName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    ChangeFileName = oFSO.BuildPath(oFSO.GetParentFolderName(sPath), sFileName & "." & LCase(oFSO.GetExtensionName(sPath)))
End Function

' �g���q����
Private Function isExtentionName(ByVal sPath As String, ByVal sExtensionNames As String) As Boolean
    Dim bRet As Boolean
    Dim elm As Variant
    For Each elm In Split(sExtensionNames, ",")
        If UCase(CreateFSO().GetExtensionName(sPath)) = UCase(elm) Then bRet = True
    Next
    isExtentionName = bRet
End Function

' �g���q�擾
Function GetExtensionName(ByVal sPath As String) As String
    GetExtensionName = LCase(CreateFSO().GetExtensionName(sPath))
End Function

' �g���q�ύX
Function ChangeExtensionName(ByVal sPath As String, ByVal sExtensionName As String) As String
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    ChangeExtensionName = oFSO.BuildPath(oFSO.GetParentFolderName(sPath), oFSO.GetBaseName(sPath) & "." & LCase(sExtensionName))
End Function

' ���΃p�X�擾
Function GetRelPath(ByVal sBasePath As String, ByVal sSpecPath As String) As String
    Dim sBuff As String
    sBuff = String(255, Chr(0))
    If PathRelativePathTo(StrPtr(sBuff), StrPtr(sBasePath), FILE_ATTRIBUTE_DIRECTORY, StrPtr(sSpecPath), 0) Then
        GetRelPath = Left(sBuff, InStr(sBuff, Chr(0)) - 1)
    Else
        GetRelPath = sSpecPath
    End If
End Function

' ��΃p�X�擾
Function GetAbsPath(ByVal sSpecPath As String, ByVal sBasePath As String) As String
    Dim sAbsPath As String
    If sSpecPath <> "" And sBasePath <> "" Then
        ' �x�[�X�p�X�͐�΃p�X��O��Ƃ���
        ' ���x�[�X�p�X���ŏ������΃p�X�Ŏw�肳��Ă���Γ��ɉe���̂Ȃ��O�̂��߂̏���
        sBasePath = CreateFSO().GetAbsolutePathName(sBasePath) ' GetAbsolutePathName�̃p�X�����̓J�����g�f�B���N�g���Œ�
        ' UNC�́u\\�v���p�X�Z�p���[�^�u\�v�ƍ������Ďז��Ȃ̂�
        ' �����I�ɓs���̗ǂ��\���ɂȂ�悤�ɑO�������Ă���
        Dim bSUNC As Boolean
        Dim bBUNC As Boolean
        sSpecPath = Replace(sSpecPath, "/", "\")
        sBasePath = Replace(sBasePath, "/", "\")
        If Left(sSpecPath, 2) = "\\" Then
            sSpecPath = Replace(sSpecPath, "\\", "<UNC>", 1, 1) ' <>�̓t�@�C�����Ɏg���Ȃ�����
            bSUNC = True
        End If
        If Left(sBasePath, 2) = "\\" Then
            sBasePath = Replace(sBasePath, "\\", "<UNC>", 1, 1) ' <>�̓t�@�C�����Ɏg���Ȃ�����
            bBUNC = True
        End If
        sSpecPath = PathTrim(sSpecPath, "\")
        sBasePath = PathTrim(sBasePath, "\")
        sSpecPath = IIf(bSUNC = False And InStr(sSpecPath, "\") = 0, ".\", "") & sSpecPath
        sBasePath = IIf(bBUNC = False And InStr(sBasePath, "\") = 0, ".\", "") & sBasePath
        ' �p�X�Z�p���[�^�Ńo�����āu.�v�u..�v�ɏ]�����p�X�ҏW��������{
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
        ' �p�X�Z�p���[�^�̘A�����폜
        Dim lAbsPath As Long
        Do
            lAbsPath = Len(sAbsPath)
            sAbsPath = Replace(sAbsPath, "\\", "\")
        Loop While lAbsPath > Len(sAbsPath)
        ' UNC�ړ��������ɖ߂�
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

' �p�X�̃T�j�^�C�Y
Function SanitizePath(ByVal sText As String) As String
    sText = Trim(sText)
    sText = Replace(sText, vbCr, "")
    sText = Replace(sText, vbLf, "")
    sText = Replace(sText, "/", "\")
    ' �h���C�u�������ƁA�t�H���_or�t�@�C���������𕪗�
    Dim sDPart As String
    Dim sPPart As String
    If sText <> "" Then
        sDPart = GetDriveName(sText, False)
        sPPart = Mid(sText, InStr(sText, sDPart) + Len(sDPart))
    End If
    ' �t�H���_or�t�@�C���������̌���
    Dim sPath As Variant
    Dim sPart As Variant
    For Each sPart In Split(sPPart, "\")
        sPath = CombinePath(sPath, SanitizeFileName(CStr(sPart)))
    Next
    ' ����
    Dim sRet As String
    If sDPart <> "" Then
        sRet = CombinePath(sDPart & "\", sPath)
    Else
        sRet = sPath
    End If
    SanitizePath = sRet
End Function

' �t�@�C�����̃T�j�^�C�Y
Function SanitizeFileName(ByVal sText As String) As String
    Dim sRet As String
    sRet = sText
    ' ���s�R�[�h�폜
    sRet = Replace(sRet, vbCr, "")
    sRet = Replace(sRet, vbLf, "")
    ' �p�X�Ɏg���Ȃ������𔭌�������S�p�ɏC������
    ' �Ԉ���ăh���C�u���Ƃ��Ɏg���Ɓu:�v���S�p�ɂȂ�̂Œ���
    sRet = Replace(sRet, "\", "��")
    sRet = Replace(sRet, ":", "�F")
    sRet = Replace(sRet, "/", "�^")
    sRet = Replace(sRet, "*", "��")
    sRet = Replace(sRet, "?", "�H")
    sRet = Replace(sRet, """", Chr(&H8168))
    sRet = Replace(sRet, "<", "��")
    sRet = Replace(sRet, ">", "��")
    sRet = Replace(sRet, "|", "�b")
    ' �\��f�o�C�X���Ɉ�v����悤�Ȃ�u��������
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
'    Debug.Assert (SanitizePath("C:") = "C:\")                               ' ��΃p�X:C:\
'    Debug.Assert (SanitizePath("C:\") = "C:\")                              ' ��΃p�X:C:\
'    Debug.Assert (SanitizePath("C:\USER\") = "C:\USER")                     ' ��΃p�X:C:\USER
'    Debug.Assert (SanitizePath("AAA") = "AAA")                              ' ���΃p�X:AAA
'    Debug.Assert (SanitizePath(".\AAA") = ".\AAA")                          ' ���΃p�X:.\AAA
'    Debug.Assert (SanitizePath("..\AAA") = "..\AAA")                        ' ���΃p�X:..\AAA
'    Debug.Assert (SanitizePath("\\NAS0001") = "\\NAS0001\")                 ' UNC�p�X :\\NAS0001\           ��Windows��UNC�p�X�Ƃ��Ė{����NG�����ǂ����ł�UNC�������Ă���
'    Debug.Assert (SanitizePath("\\NAS0001\D") = "\\NAS0001\D\")             ' UNC�p�X :\\NAS0001\D\
'    Debug.Assert (SanitizePath("\\NAS0001\D\USER") = "\\NAS0001\D\USER")    ' UNC�p�X :\\NAS0001\D\USER
'End Sub

' �V�X�e���h���C�u�p�X
Function GetSystemDrivePath() As String
    GetSystemDrivePath = Environ("SystemDrive")
End Function

' �f�X�N�g�b�v�p�X
Function GetDesktopPath() As String
    GetDesktopPath = CreateWSH().SpecialFolders("Desktop")
End Function

' �}�C�h�L�������g�p�X
Function GetMyDocumentsPath() As String
    GetMyDocumentsPath = CreateWSH().SpecialFolders("MyDocuments")
End Function

' �_�E�����[�h�p�X
Function GetDownloadPath() As String
    GetDownloadPath = CreateShellApp().Namespace("shell:Downloads").Self.Path
End Function

' ���[�U�v���t�@�C���p�X
Function GetUserProfilePath() As String
    GetUserProfilePath = Environ("USERPROFILE")
End Function

' �A�v���P�[�V�����f�[�^�p�X
Function GetAppDataPath() As String
    GetAppDataPath = Environ("APPDATA")
End Function

' �e���|�����p�X�擾
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

' �u�b�N�p�X
Function BookPath() As String
    BookPath = ThisWorkbook.Path
End Function

' �n�C�p�[�����N�̊�_
Function HyperlinkBasePath(ByVal oBook As Workbook) As String
    Dim sRet As String
    sRet = oBook.BuiltinDocumentProperties("Hyperlink base").Value
    If sRet = "" Then
        sRet = oBook.Path
    End If
    HyperlinkBasePath = sRet
End Function

' �h���C������
Function DomainName() As String
    DomainName = Environ("USERDOMAIN")
End Function

' �R���s���[�^����
Function ComputerName() As String
    ComputerName = Environ("COMPUTERNAME")
End Function

' �I�y���[�e�B���O�V�X�e������
Function OperatingSystemName() As String
    With CreateObject("WbemScripting.SWbemLocator")
        Dim elm As Object
        For Each elm In .ConnectServer.ExecQuery("Select * From Win32_OperatingSystem")
            OperatingSystemName = elm.Caption & " (" & elm.OSArchitecture & ") Version " & elm.Version
        Next
    End With
End Function

' �v���Z�b�T�A�[�L�e�N�`������
Function ArchitectureName() As String
    ArchitectureName = CreateWSH().Environment("Process").Item("PROCESSOR_ARCHITECTURE")
End Function

' OS���[�U��
Function UserName() As String
    UserName = Environ("USERNAME")
End Function

' �A�v���P�[�V�������[�U��
Function AppUserName(ByVal sDefaultName As String, Optional ByVal bFamilyNameOnly As Boolean = False)
    Dim sText As String
    sText = Application.UserName
    If bFamilyNameOnly Then
        sText = Replace(sText, "�@", " ")   ' �S�p��
        sText = Trim(sText)                 ' �O���
        If sText <> "" Then
            sText = Split(sText, " ")(0)
        End If
    End If
    If sText = "" Then
        sText = sDefaultName
    End If
    AppUserName = sText
End Function

' �w��t�H���_�z���̃t�@�C����
Function EnumFile(ByVal sPath As String, ByVal sPtrn As String, Optional ByVal lRecursive As Long = 0) As Collection
    Dim oRet As New Collection  ' �߂�l�ɂ���ƍċA���ɏd�����A�g������D���
On Error GoTo ErrExit
    ' FSO���ƃ��X�g�擾�O�Ƀt�B���^�o���Ȃ��ďd���̂�API���g�����@�ɂ���
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
            For Each sSubDElm In EnumFolder(sPath, "*", 0) ' �ċA�̏ꍇ�A�q�t�H���_�͖��O�Ńt�B���^�����S�������Ώۂɂ���
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

' �w��t�H���_�z���̃t�H���_��
' �EsPtrn�͖ړI�̊K�w���P�̂̏ꍇ�ɂ̂ݎw�肷�鎖
Function EnumFolder(ByVal sPath As String, ByVal sPtrn As String, Optional ByVal lRecursive As Long = 0) As Collection
    Dim oRet As New Collection  ' �߂�l�ɂ���ƍċA���ɏd�����A�g������D���
On Error GoTo ErrExit
    ' FSO���ƃ��X�g�擾�O�Ƀt�B���^�o���Ȃ��ďd���̂�API���g�����@�ɂ���
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

' �w�肵���p�X�{�p�^�[���Ńp�X��⊮����
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

' �w�肵���p�X�{�p�^�[���Ńp�X��⊮����
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

' �w��p�X�Łu���閼�̂ɍł������v���̂̃t�@�C����T��
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

' �w��p�X�Łu���閼�̂ɍł������v���̂̃t�H���_��T��
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

' ���[�x���V���^�C�������䗦
Private Function LevenshteinRatio(ByVal sLhs As String, ByVal sRhs As String) As Double
    LevenshteinRatio = 1# - CDbl(LevenshteinDistance(sLhs, sRhs) / WorksheetFunction.Max(Len(sLhs), Len(sRhs)))
End Function

' ���[�x���V���^�C������(��������̗ގ��x�v�Z)
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

' �t�@�C���I���_�C�A���O
Function DialogOpenFileName( _
    Optional ByVal sTitle As String = "�J��", _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal sFileFilter As String = "�S��,*.*", _
    Optional ByVal bMultiSelect As Boolean = False _
) As Collection
    Dim oRet As New Collection ' �P��I���̂���ł����[�v�ŏ����Ƃ��ƌ����Ă���
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = sTitle
        .InitialFileName = sInitPath & IIf(Right(sInitPath, 1) <> "\", "\", "")
        .Filters.Clear
        ' sFileFilter�͂���Ȋ����Ɏw�肷��:"Excel,*.xls*,���̑�,*.txt;*.csv"
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

' �t�H���_�I���_�C�A���O
Function DialogOpenFolderName( _
    Optional ByVal sTitle As String = "�J��", _
    Optional ByVal sInitPath As String = "", _
    Optional ByVal bMultiSelect As Boolean = False _
) As Collection
    Dim oRet As New Collection ' �P��I���̂���ł����[�v�ŏ����Ƃ��ƌ����Ă���
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

' // �t�@�C������ /////////////////////////////////////////

' �t�@�C���쐬
Sub MakeFile(ByVal sPath As String)
    With CreateFSO().OpenTextFile(sPath, 8, True)
        .Close
    End With
End Sub

' �t�@�C�����ݔ���
Function IsExistFile(ByVal sPath As String)
    IsExistFile = CreateFSO().FileExists(sPath)
End Function

' �t�@�C���폜
Sub RemoveFile(ByVal sPath As String, Optional ByVal bForce As Boolean = False)
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) = True Then
        Call oFSO.DeleteFile(sPath, bForce)
    End If
End Sub

' �t�@�C������
Sub CopyFile(ByVal sSrcPath As String, ByVal sDstPath As String, Optional bReplace = False)
    If sSrcPath <> sDstPath Then
        Call CreateFSO().CopyFile(sSrcPath, sDstPath, bReplace)
    End If
End Sub

' �t�@�C���ړ�
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

' �t�H���_�쐬
Sub MakeDir(ByVal sPath As String)
    If CreateFSO().FolderExists(sPath) = False Then
        ' API�ő��݂��Ȃ����ԃt�H���_����C�ɍ쐬������
        Call SHCreateDirectoryEx(0&, StrPtr(sPath), 0&)
    End If
End Sub

' �t�H���_���ݔ���
Function IsExistDir(ByVal sPath As String)
    IsExistDir = CreateFSO().FolderExists(sPath)
End Function

' �t�H���_�폜
Sub RemoveDir(ByVal sPath As String, Optional ByVal bForce As Boolean = False)
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FolderExists(sPath) = True Then
        Call oFSO.DeleteFolder(sPath, bForce)
    End If
End Sub

' �t�H���_����
Sub CopyDir(ByVal sSrcPath As String, ByVal sDstPath As String, Optional ByVal bReplace As Boolean = False)
    If sSrcPath <> sDstPath Then
        Call CreateFSO().CopyFolder(sSrcPath, sDstPath, bReplace)
    End If
End Sub

' �t�H���_�ړ�
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

' // �t�@�C�����t /////////////////////////////////////////

' �쐬�����擾
Function GetFileDateCreated(ByVal sPath As String) As Date
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) Then
        GetFileDateCreated = oFSO.GetFile(sPath).DateCreated
    Else
        GetFileDateCreated = vbEmpty
    End If
End Function

' �ŏI�X�V�����擾
Function GetFileDateLastModified(ByVal sPath As String) As Date
    Dim oFSO As Object
    Set oFSO = CreateFSO()
    If oFSO.FileExists(sPath) Then
        GetFileDateLastModified = oFSO.GetFile(sPath).DateLastModified
    Else
        GetFileDateLastModified = vbEmpty
    End If
End Function

' �ŏI�X�V�����ݒ�
Sub SetFileDateLastModified(ByVal sPath As String, ByVal oDate As Date)
    Dim oFSO As Object
    Dim oSHA As Object
    Set oFSO = CreateFSO()
    Set oSHA = CreateShellApp()
    If oFSO.FileExists(sPath) Then
        sPath = oFSO.GetAbsolutePathName(sPath)
        Dim sFPath As Variant ' �����̂�String�ł̓_��
        Dim sFName As Variant ' �����̂�String�ł̓_��
        sFPath = oFSO.GetParentFolderName(sPath)
        sFName = oFSO.GetFileName(sPath)
        Dim oFldr As Object
        Set oFldr = oSHA.Namespace(sFPath)
        Dim oFile As Object
        Set oFile = oFldr.ParseName(sFName)
        oFile.ModifyDate = oDate
    End If
End Sub

' �A�N�Z�X�����擾
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

' �e�L�X�g��������
Public Sub WriteIniText(ByVal sPath As String, ByVal sSection As String, ByVal sKey As String, ByVal sVal As String)
    Call WritePrivateProfileString(sSection, sKey, sVal, sPath)
End Sub

' �e�L�X�g�ǂݏo��
Public Function ReadIniText(ByVal sPath As String, ByVal sSection As String, ByVal sKey As String, Optional ByVal sDef As String = "") As String
    Dim sBuf As String * 1024
    Call GetPrivateProfileString(sSection, sKey, sDef, sBuf, Len(sBuf), sPath)
    ReadIniText = Left$(sBuf, InStr(sBuf, vbNullChar) - 1)
End Function

' // �O���v���Z�X���� /////////////////////////////////////

' �p�X���J��
Sub OpenPath(ByVal sPath As String)
    ' �f�B���N�g���Ȃ�G�N�X�v���[�����J��
    ' �t�@�C���Ȃ�֘A�t���ɏ]���ĊJ��
    Call CreateShellApp().ShellExecute(sPath)
End Sub

' �p�X���J��(�g��)
Sub OpenPathEx(ByVal sPath As String, ByVal sFilter As String)
    ' �G�N�X�v���[���̃t�B���^���w�肵�ĊJ��
    ' �t�@�C���X�L�[�}�w��̗��p�ł����Ƃ��낢��o���邯�Ǐ�p�ł������Ȃ̂͂���Ȃ��񂵂��Ȃ�
    Call CreateShellApp().ShellExecute("search-ms:query=" & sFilter & "&" & "crumb=location:" & sPath)
End Sub

' �R�}���h���݊m�F
Function IsExistCmd(ByVal sCommand As String) As Boolean
    Dim bRet As Boolean
    bRet = CreateWSH().Run("%ComSpec% /c where " & sCommand & " > nul 2>&1", 0, True)
    IsExistCmd = Not bRet
End Function

' �R�}���h�������s(�W���o�͖�)
Function RunCmd(ByVal sCommand As String) As Long
    Dim lResult As Long
On Error GoTo ErrExit
    lResult = CreateWSH().Run("%ComSpec% /c " & PathDQuote(sCommand), 0, True)
ErrExit:
    RunCmd = lResult
End Function

' �R�}���h�������s(�W���o�͗L)
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

' PowerShell�������s(�W���o�͖�)
Function RunPsh(ByVal sCommand As String) As Long
    Dim lResult As Long
On Error GoTo ErrExit
    lResult = CreateWSH().Run("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & PathDQuote(sCommand), 0, True)
ErrExit:
    RunPsh = lResult
End Function

' PowerShell�������s(�W���o�͗L)
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

' �_�E�����[�h
Function DownloadFile(ByVal sDLSrc As String, ByVal sDLDst As String) As Long
    Dim lRet As Long
    lRet = URLDownloadToFile(0, sDLSrc, sDLDst, 0, 0)
    If lRet = 0 Then
        Do Until IsExistFile(sDLDst) = True
            DoEvents
        Loop
    End If
    DownloadFile = lRet
End Function

' ���k
Function CompressArchive(ByVal sSrcPath As String, ByVal sDstPath As String) As Long
    CompressArchive = RunPsh("Compress-Archive -Path " & sSrcPath & " -DestinationPath " & sDstPath & " -Force")
End Function

' ��
Function ExtractArchive(ByVal sSrcPath As String, ByVal sDstPath As String) As Long
    ExtractArchive = RunPsh("Expand-Archive -Path " & sSrcPath & " -DestinationPath " & sDstPath & " -Force")
End Function

