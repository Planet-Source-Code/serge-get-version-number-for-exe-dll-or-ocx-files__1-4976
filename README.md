<div align="center">

## Get Version Number for EXE, DLL or OCX files


</div>

### Description

This function will retrieve the version number, product name, original program name (like if you right click on the EXE file and select properties, then select Version tab, it shows you all that information) etc
 
### More Info
 
Label (named Label1 and make it wide enough, also increase the height of the label to have size of the form), Common Dilaog Box (CommonDialog1) and a Command Button (Command1)

FileInfo structure


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Serge](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/serge.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/serge-get-version-number-for-exe-dll-or-ocx-files__1-4976/archive/master.zip)

### API Declarations

```
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Type FILEINFO
  CompanyName As String
  FileDescription As String
  FileVersion As String
  InternalName  As String
  LegalCopyright As String
  OriginalFileName As String
  ProductName As String
  ProductVersion As String
End Type
Public Enum VerisonReturnValue
  eOK = 1
  eNoVersion = 2
End Enum
```


### Source Code

```
Public Function GetFileVersionInformation(ByRef pstrFieName As String, ByRef tFileInfo As FILEINFO) As VerisonReturnValue
  Dim lBufferLen As Long, lDummy As Long
  Dim sBuffer() As Byte
  Dim lVerPointer As Long
  Dim lRet As Long
  Dim Lang_Charset_String As String
  Dim HexNumber As Long
  Dim i As Integer
  Dim strTemp As String
  'Clear the Buffer tFileInfo
  tFileInfo.CompanyName = ""
  tFileInfo.FileDescription = ""
  tFileInfo.FileVersion = ""
  tFileInfo.InternalName = ""
  tFileInfo.LegalCopyright = ""
  tFileInfo.OriginalFileName = ""
  tFileInfo.ProductName = ""
  tFileInfo.ProductVersion = ""
  lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)
  If lBufferLen < 1 Then
    GetFileVersionInformation = eNoVersion
    Exit Function
  End If
  ReDim sBuffer(lBufferLen)
  lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))
  If lRet = 0 Then
    GetFileVersionInformation = eNoVersion
    Exit Function
  End If
  lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)
  If lRet = 0 Then
    GetFileVersionInformation = eNoVersion
    Exit Function
  End If
  Dim bytebuffer(255) As Byte
  MoveMemory bytebuffer(0), lVerPointer, lBufferLen
  HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
  Lang_Charset_String = Hex(HexNumber)
  'Pull it all apart:
  '04------    = SUBLANG_ENGLISH_USA
  '--09----    = LANG_ENGLISH
  ' ----04E4 = 1252 = Codepage for Windows:Multilingual
  Do While Len(Lang_Charset_String) < 8
    Lang_Charset_String = "0" & Lang_Charset_String
  Loop
  Dim strVersionInfo(7) As String
  strVersionInfo(0) = "CompanyName"
  strVersionInfo(1) = "FileDescription"
  strVersionInfo(2) = "FileVersion"
  strVersionInfo(3) = "InternalName"
  strVersionInfo(4) = "LegalCopyright"
  strVersionInfo(5) = "OriginalFileName"
  strVersionInfo(6) = "ProductName"
  strVersionInfo(7) = "ProductVersion"
  Dim buffer As String
  For i = 0 To 7
    buffer = String(255, 0)
    strTemp = "\StringFileInfo\" & Lang_Charset_String _
    & "\" & strVersionInfo(i)
    lRet = VerQueryValue(sBuffer(0), strTemp, _
    lVerPointer, lBufferLen)
    If lRet = 0 Then
      GetFileVersionInformation = eNoVersion
      Exit Function
    End If
    lstrcpy buffer, lVerPointer
    buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)
    Select Case i
      Case 0
        tFileInfo.CompanyName = buffer
      Case 1
        tFileInfo.FileDescription = buffer
      Case 2
        tFileInfo.FileVersion = buffer
      Case 3
        tFileInfo.InternalName = buffer
      Case 4
        tFileInfo.LegalCopyright = buffer
      Case 5
        tFileInfo.OriginalFileName = buffer
      Case 6
        tFileInfo.ProductName = buffer
      Case 7
        tFileInfo.ProductVersion = buffer
    End Select
  Next i
  GetFileVersionInformation = eOK
End Function
'-----------
Private Sub Command1_Click()
  Dim strFile As String
  Dim udtFileInfo As FILEINFO
  On Error Resume Next
  With CommonDialog1
    .Filter = "All Files (*.*)|*.*"
    .ShowOpen
    strFile = .FileName
    If Err.Number = cdlCancel Or strFile = "" Then Exit Sub
  End With
  If GetFileVersionInformation(strFile, udtFileInfo) = eNoVersion Then
    MsgBox "No version available for this file", vbInformation
    Exit Sub
  End If
  Label1 = "Company Name:           " & udtFileInfo.CompanyName & vbCrLf
  Label1 = Label1 & "File Description:    " & udtFileInfo.FileDescription & vbCrLf
  Label1 = Label1 & "File Version:      " & udtFileInfo.FileVersion & vbCrLf
  Label1 = Label1 & "Internal Name:     " & udtFileInfo.InternalName & vbCrLf
  Label1 = Label1 & "Legal Copyright:   " & udtFileInfo.LegalCopyright & vbCrLf
  Label1 = Label1 & "Original FileName:  " & udtFileInfo.OriginalFileName & vbCrLf
  Label1 = Label1 & "Product Name:    " & udtFileInfo.ProductName & vbCrLf
  Label1 = Label1 & "Product Version:   " & udtFileInfo.ProductVersion & vbCrLf
End Sub
```

