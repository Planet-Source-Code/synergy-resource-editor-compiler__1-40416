Attribute VB_Name = "ModGeneral"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal wHWND As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public strResource As String
Public strFile  As String
Public strSample As String
Public tBox As Byte

'Track changes
Public Saved As Boolean
Public IsSaved As Boolean

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const EM_SETMARGINS = &HD3
Public Const EC_LEFTMARGIN = &H1
Public Const EC_RIGHTMARGIN = &H2

Public Const EM_SETTABSTOPS = &HCB

Public Sub SetMargins()
  Dim X As Long, Y As Long, z As Long
  X = 5: Y = 5
  z = X * &H10000 + Y
  SendMessage frmEditor.txtString(2).hWnd, EM_SETMARGINS, EC_LEFTMARGIN Or EC_RIGHTMARGIN, z
End Sub

Public Function GetInserts(ByVal strTxt As String)
  GetInserts = Replace(strTxt, "/7", Chr$(38))
  GetInserts = Replace(strTxt, "/9", Chr$(60))
  GetInserts = Replace(strTxt, "/0", Chr$(62))
  GetInserts = Replace(strTxt, "/n", Chr$(34))
  GetInserts = Replace(strTxt, "/t", vbTab)
  GetInserts = Replace(strTxt, "/crlf ", vbCrLf)
  GetInserts = Replace(strTxt, "/cr ", vbCr)
  GetInserts = Replace(strTxt, "/lf ", vbLf)
End Function

Public Function PutInserts(ByVal strTxt As String)
'Each reference in the precompiled resource file
'must be a single line of text.  These characters
'are replace to keep references to one line or
'remove characters that will cause compiler/xml errors
  PutInserts = Replace(strTxt, Chr$(38), "/7")
  PutInserts = Replace(strTxt, Chr$(60), "/9")
  PutInserts = Replace(strTxt, Chr$(62), "/0")
  PutInserts = Replace(strTxt, Chr$(34), "/n")
  PutInserts = Replace(strTxt, vbTab, "/t")
  PutInserts = Replace(strTxt, vbCrLf, "/crlf ")
  PutInserts = Replace(strTxt, vbCr, "/cr ")
  PutInserts = Replace(strTxt, vbLf, "/lf ")
End Function


Public Sub CheckSave(ByRef frm As Form, ByVal srcFile As String, ByVal sSave As Boolean)
'Watch changes and Update
Dim msgFile As String

If sSave Then
  Exit Sub
Else
  If InStr(srcFile, ".") Then
    msgFile = strFileName(srcFile)
  Else
    msgFile = srcFile
  End If
    If MsgBox("   Save Changes to '" & UCase$(msgFile) & "' Template   ", vbYesNo Or vbQuestion, "Template has Changed") = vbYes Then
      Select Case frm.Name
      Case "frmEditor"
        frm.SaveXML
      Case "frmStrings"
        frm.SaveXML srcFile
      End Select
    End If
End If
End Sub

Public Sub PrintCode(ByVal strText As String)
'Wordwrap long lines
On Error Resume Next
Dim txt As String
Dim bytWrap As Byte, bytCount As Byte, bytSpace As Byte
'Set this value to your liking
bytWrap = 80

Y = Len(strText)
For X = 1 To Y
  'Start adding the length of the line
  bytCount = bytCount + 1
  'Check each character
  txt = Asc(Mid$(strText, X, 1))
  'Save this value for long or hyphenated words at the endofline
    If txt = 32 Or txt = 45 Then bytSpace = bytCount
  
    'First check if at the end of text in textbox
    If X = Y Then
      'If yes, print it
      Printer.Print Mid$(strText, X - bytCount + 1, bytCount)
    
    'Check for endofparagraph
    ElseIf txt = 13 Then
      Printer.Print Mid$(strText, X - bytCount + 1, bytCount)
      'Move passed vbCrLf
      X = X + 1
      'Reset line length
      bytCount = 0
      
    'Check for long words at endofline
    ElseIf bytCount >= bytWrap + 5 Then
      'if yes, backup and print at last space or hyphen
      Printer.Print Mid$(strText, X - bytCount + 1, bytSpace)
      'Backup the loop
      X = X - bytCount + bytSpace
      bytCount = 0
      
    'Break all other line at a space...chr$(32)
    ElseIf txt = 32 And bytCount >= bytWrap Then
      Printer.Print Mid$(strText, X - bytCount + 1, bytCount)
      bytCount = 0
      
    End If
  'Check for interrupts?
  DoEvents
Next

End Sub

Public Sub selTXT()
'Cuts down on repeditive code
'Select the entire textbox on GotFocus
With Screen.ActiveForm.ActiveControl
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Function GetEnvironmentVar(sName As String) As String
'One method for passing varables between forms
  GetEnvironmentVar = String(255, 0)
  GetEnvironmentVariable sName, GetEnvironmentVar, Len(GetEnvironmentVar)
End Function

Public Function InStrRev(ByVal strSearch As String, ByVal strFind As String) As Integer
'For vb4 & 5 programmers
On Error Resume Next
Dim X As Integer

  For X = Len(strSearch) To 1 Step -1
    InStrRev = InStr(X, strSearch, strFind)
    If InStrRev > 0 Then Exit For
  Next
End Function

Public Function Replace(ByVal strOld As String, ByVal strSearch As String, ByVal strReplace As String) As String
'For vb4 & 5 programmers
Dim X As Integer
Dim Y As Integer

  Y = Len(strSearch)
  
  For X = 1 To Len(strOld) - (Y - 1)
  If InStr(X, strOld, strSearch) > 0 Then
    X = InStr(X, strOld, strSearch)
        txt = Left$(strOld, X - 1)
        txts = Right$(strOld, Len(strOld) - X - Len(strSearch) + 1)
      strOld = txt & strReplace & txts
      X = Len(txt & strReplace)
    
    If X > Len(strOld) - Y Then Exit For
  End If
  Next X
    
Replace = strOld
End Function

