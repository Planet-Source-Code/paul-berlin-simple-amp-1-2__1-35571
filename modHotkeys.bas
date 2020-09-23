Attribute VB_Name = "modHotkeys"
Option Explicit
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function ReadINI(SectionName As String, KeyName As String, FileName As String) As String
  'Reads value from ini-file with API
  Dim tmpBuffer As String * 255
  GetPrivateProfileString SectionName, KeyName, "NOT FOUND", tmpBuffer, Len(tmpBuffer), FileName
  ReadINI = Mid(tmpBuffer, 1, InStr(1, tmpBuffer, vbNullChar) - 1)
End Function

Public Function FileExists(FileName As String) As Boolean
  'Cheks if file exists
  FileExists = Not (Dir(FileName) = "")
End Function
