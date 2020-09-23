Attribute VB_Name = "modClipBoard"
'Adapted from The KPD-Team
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)
Dim PrevProc As Long


Public Sub HookForm(F As Form)
    PrevProc = SetWindowLong(F.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub UnHookForm(F As Form)
    SetWindowLong F.hWnd, GWL_WNDPROC, PrevProc
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
Dim ClpFmt As Integer
Dim clpAction As String, clpEX As String
Dim txt As String, txt1 As String
Dim msg As String, msg1 As String

WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
If uMsg = WM_DRAWCLIPBOARD Then
    
  If Clipboard.GetFormat(vbCFText) Then
    clpAction = "Text"
    clpEX = ".txt"
  ElseIf Clipboard.GetFormat(vbCFBitmap) Then
    clpAction = "Bitmap"
    clpEX = ".bmp"
  ElseIf Clipboard.GetFormat(vbCFMetafile) Then
    clpAction = "wMetaFile"
    clpEX = ".wmf"
  ElseIf Clipboard.GetFormat(vbCFDIB) Then
    clpAction = "DIBitmap"
    clpEX = ".dib"
  ElseIf Clipboard.GetFormat(vbCFRTF) Then
    clpAction = "RichText"
    clpEX = ".rtf"
  Else
    clpAction = ""
  End If
  
    msg = "There is " & clpAction & " Data on the ClipBoard." & vbCrLf & vbCrLf
    msg = msg & "Do you want to get this Data?"
  If clpAction <> "" Then
    If MsgBox(msg, vbYesNo Or vbQuestion Or vbSystemModal, App.Title & " _ ClipBoard Data") = vbYes Then
    
    
      If frmEditor.WindowState = vbMinimized Then
        frmEditor.WindowState = vbNormal
      End If
      
      
      If clpAction = "Text" Then
        txt = Clipboard.GetText()
        
        If Len(txt) < 240 Then
          frmEditor.txtString(0) = clpAction & "Name"
          frmEditor.txtString(1) = txt
        Else 'Text too Long, put it in a file
        
          msg1 = InputBox("Enter fileName for new text file", "Get Text Resource", "newTextFileName")
          frmEditor.txtString(0) = msg1
          
          'Check inputbox return
          If InStr(msg1, ".") Then
            txt1 = Left$(msg1, InStr(msg1, ".") - 1)
            clpEX = Right$(msg1, (InStrRev(msg1, ".") - 1) - Len(msg1))
          Else
            txt1 = msg1
          End If
          If CBool(PathFileExists(txt1 & clpEX)) Then
            For X = 1 To 999
              If Not CBool(PathFileExists(txt1 & X & clpEX)) Then
                binSave strFilePath(strFile) & txt1 & X & clpEX, txt
                frmEditor.txtString(1) = txt1 & X & clpEX
                Exit For
              End If
            Next
          Else
            binSave strFilePath(strFile) & txt1 & clpEX, txt
            frmEditor.txtString(1) = txt1 & clpEX
          End If
          
          If Not Err Then MsgBox "To Include this new Resource," & vbCrLf & vbCrLf & "Click [Add Item] on the Resource Editor.", vbInformation, "Reminder"
        End If 'Len text
        
      Else 'Not text
      
          msg1 = InputBox("Enter fileName for new " & clpAction & " file", "Get Text Resource", "newTextFileName")
          frmEditor.txtString(0) = msg1
          'Check inputbox return
          If InStr(msg1, ".") Then
            txt1 = Left$(msg1, InStr(msg1, ".") - 1)
            clpEX = Right$(msg1, (InStrRev(msg1, ".") - 1) - Len(msg1))
          Else
            txt1 = msg1
          End If
          If CBool(PathFileExists(txt1 & clpEX)) Then
            For X = 1 To 999
              If Not CBool(PathFileExists(txt1 & X & clpEX)) Then
                binSave strFilePath(strFile) & txt1 & X & clpEX, txt
                frmEditor.txtString(1) = txt1 & X & clpEX
                Exit For
              End If
            Next
          Else
            binSave strFilePath(strFile) & txt1 & clpEX, txt
            frmEditor.txtString(1) = txt1 & clpEX
          End If
        
          If Not Err Then MsgBox "To Include this new Resource," & vbCrLf & vbCrLf & "Click [Add Item] on the Resource Editor.", vbInformation, "Reminder"
      End If 'Graphics
      
      
    End If 'msgbox save it
  End If 'clpAction
End If 'uMsg

If uMsg = WM_DROPFILES Then

End If
If Err Then MsgBox "Could not include new Resource.", vbCritical, "Error - ClipBoard"
End Function


