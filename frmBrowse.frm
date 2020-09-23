VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowse 
   Caption         =   "Browse for Files - Resource Editor"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   2640
      ScaleHeight     =   4875
      ScaleWidth      =   4635
      TabIndex        =   6
      Top             =   360
      Width           =   4695
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4935
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   4695
         ExtentX         =   8281
         ExtentY         =   8705
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   45
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2490
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   4
      Top             =   30
      Width           =   2490
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   45
      TabIndex        =   3
      Top             =   390
      Width           =   2490
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   330
      Left            =   3840
      TabIndex        =   2
      Top             =   15
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   330
      Left            =   2640
      TabIndex        =   1
      Top             =   15
      Width           =   1185
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   45
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   2880
      Width           =   2490
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
  File1.Pattern = Combo1.Text
End Sub

Private Sub Combo1_Click()
  File1.Pattern = Combo1.Text
End Sub


Private Sub Command1_Click()
On Error Resume Next
Dim strdata As String
Dim strName As String
Dim tvParent As Node
Dim booParent As Boolean
Dim noDx As Node

For X = 0 To File1.ListCount - 1
  If File1.Selected(X) Then
    strdata = UCase$(Right$(File1.List(X), 3))
    Select Case strdata
      Case "CUR"
        strdata = "CURSORS"
      Case "BMP"
        strdata = "BITMAPS"
      Case "ICO"
        strdata = "ICONS"
      Case "AVI"
        strdata = "VIDEOS"
      Case "WAV"
        strdata = "SOUNDS"
      Case "HTM", "HTML", "SHTML", "JS", "CSS", "XML", "XSL", "GIF", "JPG", "JPEG"
        strdata = "HTML"
      Case Else
        strdata = "CUSTOM"
    End Select
    
  'Parse the path in txt to get the name
   strName = strFileName(File1.List(X))
          
  'Check that file isn't already
  'listed on the Resource Editor's Tree
  For Y = 1 To frmEditor.tv.Nodes.Count
    If strName = frmEditor.tv.Nodes(Y).Text Then
      MsgBox "KeyItem '" & strName & "' is already in use by Resource Editor" & vbCrLf & vbCrLf & "All Identifiers must be Unique", vbOKOnly Or vbCritical, "Insert Operation Cancelled"
      GoTo hell
    End If
  Next
  
  booParent = True
  For Y = 1 To frmEditor.tv.Nodes.Count
    If strdata = UCase$(frmEditor.tv.Nodes(Y).Text) Then
      Set tvParent = frmEditor.tv.Nodes(Y)
      booParent = False
      Exit For
    End If
  Next
  If booParent Then
    Set tvParent = frmEditor.tv.Nodes.Add("root", tvwChild, strdata, strdata, 2)
      tvParent.Tag = strdata
  End If
  
  strName = StrConv(strName, vbProperCase)
  Set noDx = frmEditor.tv.Nodes.Add(tvParent, tvwChild, , strName, 3)
      noDx.Tag = LCase$(File1.List(X))
      noDx.EnsureVisible
      
  'Copy file to resource section for later compilation
  FileCopy File1.Path & "\" & File1.List(X), GetPath & "templates\" & strResource & "\" & File1.List(X)
  
  Saved = False
  End If
hell:
Next
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Dir1_Change()
  File1 = Dir1
End Sub

Private Sub Dir1_Click()
  Dir1.Path = Dir1.List(Dir1.ListIndex)
  File1 = Dir1
End Sub

Private Sub Drive1_Change()
  Dir1 = Drive1
End Sub

Private Sub File1_Click()
On Error Resume Next
Dim txt As String
  txt = LCase$(Right$(File1.filename, 3))
  
  'webbrowser doesn't show cursors
  If txt = "cur" Then
    WebBrowser1.Visible = False
    Picture1 = LoadPicture(Dir1.Path & "\" & File1.filename)
  Else
    WebBrowser1.Visible = True
    WebBrowser1.Navigate Dir1.Path & "\" & File1.filename
  End If

Y = 0
For X = 0 To File1.ListCount - 1
  If File1.Selected(X) Then Y = Y + 1
Next
If Y > 1 Then
  Command1.Caption = "&Insert All"
Else
  Command1.Caption = "&Insert"
End If
End Sub

Private Sub Form_Load()
Dim ThePattern As String
Me.Icon = frmEditor.Icon
  'Start things where they were last session
  On Error Resume Next
  ThePath = GetSetting(App.Title, "Browse", "ThePath", "Empty")
  ThePattern = GetSetting(App.Title, "Browse", "Pattern", "Empty")
  
On Error GoTo errHandler
Combo1.AddItem "*.avi;*.bmp;*.cur;*.ico;*.wav"
Combo1.AddItem "*.gif;*.jpg;*.jpeg"
Combo1.AddItem "*.txt;*.ini;*.dat;*.csv"
Combo1.AddItem "*.htm;*.html;*.shtml"
Combo1.AddItem "*.doc; *.xls; *.ppt"
Combo1.AddItem "*.izs;*.js"
Combo1.AddItem "*.css;*.xml;*.xsl"
Combo1.AddItem "*.*"
Combo1.Text = ThePattern
  File1.Pattern = Combo1.Text
  

If CStr(CBool(PathIsDirectory(ThePath))) Then
Drive1 = ThePath
Dir1 = ThePath
End If

Exit Sub
errHandler:
SaveSetting App.Title, "Browse", "ThePath", GetPath
Drive1 = GetPath
Dir1 = GetPath
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormCode Then
  'Save path for next session
  SaveSetting App.Title, "Browse", "ThePath", File1.Path
  SaveSetting App.Title, "Browse", "Pattern", Combo1.Text
End If
End Sub


Private Sub Form_Resize()
Dim lf As Integer, tf As Integer, wf As Integer
Dim lw As Integer, tw As Integer
lf = 45: tf = 2880: wf = 2490
lw = 2640: tw = 360

With Me
  If .Width < 5235 Then .Width = 5235
  If .Height < 5205 Then .Height = 5205
  If .WindowState <> 1 Then
  
  File1.Move lf, tf, wf, ScaleHeight - (tf + 60)
  Picture1.Move lw, tw, _
    ScaleWidth - (Picture1.Left + 90), _
    ScaleHeight - 480
  WebBrowser1.Move -30, -30, _
    Picture1.Width, Picture1.Height
  End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmBrowse = Nothing
End Sub

