Attribute VB_Name = "modCmax"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public CheckCh(9) As Integer

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number
'MakeFileType "txt", "Text Document", "C:\windows\notepad.exe,0", "open", "C:\windows\notepad.exe %1", False, True


Private Function GetString(hKey As Long, strPath As String, strValue As String, DefaultStr As Long)

  'EXAMPLE:
  '
  'text1.text = getstring(HKEY_CURRENT_USE
  '     R, "Software\VBW\Registry", "String")
  '
  
  Dim keyhand As Long
  Dim Datatype As Long
  Dim lResult As Long
  Dim strBuf As String
  Dim lDataBufSize As Long
  Dim intZeroPos As Integer
  Dim datas1 As String, datas2 As String
  Dim fle As Integer, r, lValueType

    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))

            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
              Else
                GetString = strBuf
            End If
        End If
    End If
    If strBuf = "" Then GetString = DefaultStr

End Function

Public Function ReadOptions(rt As CodeMax)

    Call rt.SetColor(cmClrComment, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "comment", vbGreen))
    Call rt.SetColor(cmClrBookmark, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "bookmark", vbWhite))
    Call rt.SetColor(cmClrBookmarkBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "bookmarkbk", vbRed))
    Call rt.SetColor(cmClrCommentBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "commentbk", -1))
    Call rt.SetColor(cmClrHDividerLines, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "divider", -1))
    Call rt.SetColor(cmClrHighlightedLine, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "highlight", 65535))
    Call rt.SetColor(cmClrKeyword, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "keyword", 16711680))
    Call rt.SetColor(cmClrKeywordBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "keywordbk", -1))
    Call rt.SetColor(cmClrLeftMargin, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "left", 8421504))
    Call rt.SetColor(cmClrLineNumber, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "linenum", 16777215))
    Call rt.SetColor(cmClrLineNumberBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "linenumbk", 8421504))
    Call rt.SetColor(cmClrNumber, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "number", 0))
    Call rt.SetColor(cmClrNumberBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "numberbk", -1))
    Call rt.SetColor(cmClrOperator, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "operator", vbBlack))
    Call rt.SetColor(cmClrOperatorBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "operatorbk", -1))
    Call rt.SetColor(cmClrScopeKeyword, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "scope", 16711680))
    Call rt.SetColor(cmClrScopeKeywordBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "scopebk", -1))
    Call rt.SetColor(cmClrString, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "string", vbBlack))
    Call rt.SetColor(cmClrStringBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "stringbk", -1))
    Call rt.SetColor(cmClrTagAttributeName, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagattrib", 16711680))
    Call rt.SetColor(cmClrTagAttributeNameBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagattribbk", -1))
    Call rt.SetColor(cmClrTagElementName, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagele", 128))
    Call rt.SetColor(cmClrTagElementNameBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagelebk", -1))
    Call rt.SetColor(cmClrTagEntity, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagent", 255))
    Call rt.SetColor(cmClrTagEntityBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagentbk", -1))
    Call rt.SetColor(cmClrTagText, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagtxt", 0))
    Call rt.SetColor(cmClrTagTextBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "tagtxtbk", -1))
    Call rt.SetColor(cmClrText, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "text", 0))
    Call rt.SetColor(cmClrTextBk, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "textbk", -1))
    Call rt.SetColor(cmClrVDividerLines, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "vdivider", -1))
    Call rt.SetColor(cmClrWindow, GetString(HKEY_CLASSES_ROOT, "SE\colors\", "window", -1))
  
    Call rt.SetFontStyle(cmStyComment, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k1", 0))
    Call rt.SetFontStyle(cmStyKeyword, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k2", 0))
    Call rt.SetFontStyle(cmStyLineNumber, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k3", 0))
    Call rt.SetFontStyle(cmStyNumber, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k4", 0))
    Call rt.SetFontStyle(cmStyOperator, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k5", 0))
    Call rt.SetFontStyle(cmStyScopeKeyword, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k6", 0))
    Call rt.SetFontStyle(cmStyString, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k7", 0))
    Call rt.SetFontStyle(cmStyTagAttributeName, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k8", 0))
    Call rt.SetFontStyle(cmStyTagElementName, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k9", 0))
    Call rt.SetFontStyle(cmStyTagEntity, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k10", 0))
    Call rt.SetFontStyle(cmStyTagText, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k11", 0))
    Call rt.SetFontStyle(cmStyText, GetString(HKEY_CLASSES_ROOT, "SE\fonts\", "k12", 0))
  
    rt.SelBounds = GetString(HKEY_CLASSES_ROOT, "SE\options\", "selbounds", True)
    rt.DisplayLeftMargin = GetString(HKEY_CLASSES_ROOT, "SE\options\", "leftmargin", True)
    rt.LineNumbering = GetString(HKEY_CLASSES_ROOT, "SE\data\", "numbering", True)
    rt.LineToolTips = GetString(HKEY_CLASSES_ROOT, "SE\options\", "lttips", True)
    rt.LineNumberStyle = GetString(HKEY_CLASSES_ROOT, "SE\data\", "numberingstyle", 1)
    rt.LineNumberStart = GetString(HKEY_CLASSES_ROOT, "SE\data\", "numberingstart", 1)
    rt.Font.Size = GetString(HKEY_CLASSES_ROOT, "SE\data\", "Fontsize", 10)

End Function

