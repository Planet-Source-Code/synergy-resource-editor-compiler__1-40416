Attribute VB_Name = "modFileOps"
Option Explicit

Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private F As Long
Public Const MAX_PATH = 255
Public Sub binSave(ByVal srcFile As String, ByVal txt As String)
F = FreeFile
  Open srcFile For Binary As F
    Put F, , txt
  Close F
End Sub

Public Function binOpen(ByVal aFile As String) As String
F = FreeFile
Dim txt As String

  Open aFile For Binary As F
    txt = String(LOF(F), " ")
    Get F, , txt
  Close F
binOpen = txt
End Function


Public Function GetPath()
GetPath = App.Path
If Not Right$(GetPath, 1) = "\" Then GetPath = GetPath & "\"
End Function

Public Function strFilePath(ByVal aPath As String) As String
  strFilePath = Left$(aPath, InStrRev(aPath, "\"))
End Function


Public Function strFileName(ByVal aPath As String) As String
Dim X As Byte, Y As Byte
 
  X = InStrRev(aPath, "\") + 1
  Y = InStr(aPath, ".")
  strFileName = Mid$(aPath, X, Y - X)
  
'to get the filename with extension:
'strFileName = Right$(aPath, Len(aPath) - InStrRev(aPath, "\"))
End Function


