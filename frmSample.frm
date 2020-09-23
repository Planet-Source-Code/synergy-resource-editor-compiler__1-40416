VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sample - "
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Run Mode
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private picWeb As PictureBox
Private Sub Form_Load()
'The webBrowser Control is called from
'other form's procedures, so don't
'run this code unless called from
'sample's menu click.
If strSample = "" Then Exit Sub
Dim X As Integer, Y As Integer
Me.Icon = frmEditor.Icon
Me.Caption = "Resource Editor - " & strSample

  'Adjust for ScreenResolution
  X = Screen.Width / 15
  Y = Screen.Height / 15
  If X = 640 And Y = 480 Then
    Me.Move Me.Left, Me.Top, Me.Width - 2000, Me.Height - 1000
  ElseIf X = 1024 And Y = 768 Then
    Me.Move Me.Left, Me.Top, Me.Width + 1000, Me.Height + 1000
  End If
  


'This rarely works from the vb.IDE
'Compile an exe before running this form
Dim strName As String
Dim lngRet As Long
    
  strName = String(MAX_PATH, 0)
  lngRet = GetModuleFileName(App.hInstance, strName, MAX_PATH)
  strName = Left(strName, lngRet)
  If InStr(LCase$(strName), LCase$(App.EXEName)) = 0 Then
    MsgBox "You should run this form from a Compiled exe.", _
      vbInformation, "Mode:  vb.IDE"
  End If

  frmBrowse.Picture1.Move 0, 0, ScaleWidth, ScaleHeight
  frmBrowse.WebBrowser1.Move -30, -30, _
    frmBrowse.Picture1.Width, frmBrowse.Picture1.Height
Set picWeb = frmBrowse.Picture1

'Use this method for making ebooks...
frmBrowse.WebBrowser1.Navigate "res://ResourceEditor.exe/" & strSample

SetParent picWeb.hWnd, Me.hWnd
picWeb.Visible = True
picWeb.Move 0, 0, ScaleWidth, ScaleHeight
strSample = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set frmBrowse = Nothing
  Set frmSample = Nothing
End Sub



Private Sub mBack_Click()
On Error Resume Next
frmBrowse.WebBrowser1.GoBack
End Sub


Private Sub mForward_Click()
On Error Resume Next
frmBrowse.WebBrowser1.GoForward
End Sub


Private Sub mPrint_Click()
frmBrowse.WebBrowser1.SetFocus
frmBrowse.WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub


