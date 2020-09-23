Attribute VB_Name = "ShellExec"
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function ExtractFileName(ByVal strPath As String) As String
  ' StrReverse is only working in VB6
  strPath = StrReverse(strPath)
  strPath = Left(strPath, InStr(strPath, "\") - 1)
  ExtractFileName = StrReverse(strPath)
End Function

Public Function ExtractPath(ByVal strPath As String)
  Dim strtmp As String
  Dim a As Long
  
  strtmp = StrReverse(strPath)
  a = Len(strPath) - InStr(strtmp, "\")
  strPath = Left(strPath, a)
  ExtractPath = strPath
End Function


Function AddASlash(ByVal path As String)
  If Right(path, 1) = "\" Then
    AddASlash = path
  Else
    AddASlash = path & "\"
  End If
End Function

Public Function FileExists(strPath As String) As Integer
  FileExists = Not (Dir(strPath) = "")
End Function

