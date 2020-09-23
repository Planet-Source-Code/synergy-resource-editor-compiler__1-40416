VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditor 
   Caption         =   "Resource Editor"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   6375
   Icon            =   "frmEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBoxLt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      ForeColor       =   &H80000016&
      Height          =   5175
      Left            =   60
      ScaleHeight     =   5115
      ScaleWidth      =   2385
      TabIndex        =   11
      Top             =   0
      Width           =   2445
      Begin VB.CommandButton cmdGroup 
         Caption         =   "Add &Group"
         Height          =   360
         Index           =   0
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         ToolTipText     =   "Add New Group"
         Top             =   0
         Width           =   1185
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "Delete"
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         ToolTipText     =   "Delete Selected Group"
         Top             =   0
         Width           =   1185
      End
      Begin MSComctlLib.TreeView tv 
         Height          =   4125
         Left            =   -30
         TabIndex        =   14
         Top             =   330
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   7276
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   18
         Style           =   5
         HotTracking     =   -1  'True
         ImageList       =   "imgTree"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picSplitV 
      BackColor       =   &H80000001&
      Height          =   5175
      Left            =   2520
      MousePointer    =   99  'Custom
      ScaleHeight     =   5115
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   0
      Width           =   120
   End
   Begin VB.PictureBox picBoxRt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   5175
      Left            =   2640
      ScaleHeight     =   5115
      ScaleWidth      =   3585
      TabIndex        =   8
      Top             =   0
      Width           =   3640
      Begin VB.PictureBox picSplitH 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   360
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   300
         ScaleWidth      =   3510
         TabIndex        =   10
         ToolTipText     =   "Double-Click to Refresh Code."
         Top             =   3135
         Width           =   3575
      End
      Begin VB.TextBox txtString 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Index           =   2
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3360
         Width           =   3575
      End
      Begin VB.TextBox txtString 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   3575
      End
      Begin VB.TextBox txtString 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Index           =   1
         Left            =   0
         MultiLine       =   -1  'True
         OLEDropMode     =   2  'Automatic
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1410
         Width           =   3575
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "&Change"
         Height          =   360
         Index           =   2
         Left            =   2400
         TabIndex        =   2
         ToolTipText     =   "Change Selected Section"
         Top             =   0
         Width           =   1185
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Add &Item"
         Height          =   360
         Index           =   0
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   0
         ToolTipText     =   "Add New Item to selected Section"
         Top             =   0
         Width           =   1185
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "Delete"
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Delete Selected Item"
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Value:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   5
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Key:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList imgTree 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSaveAs 
         Caption         =   "Save &As"
         Shortcut        =   ^A
      End
      Begin VB.Menu mLine 
         Caption         =   "-"
      End
      Begin VB.Menu mImport 
         Caption         =   "&Import XML"
      End
      Begin VB.Menu mCompile 
         Caption         =   "&Compile Res"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mPrint 
         Caption         =   "&Print"
         Begin VB.Menu mPrtXML 
            Caption         =   "Print &XML File"
         End
         Begin VB.Menu mPrtRC 
            Caption         =   "Print &PreCompile"
         End
      End
      Begin VB.Menu mUnInstall 
         Caption         =   "&UnInstall"
      End
      Begin VB.Menu mLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mAddGroup 
         Caption         =   "Add &Group"
         Shortcut        =   ^G
      End
      Begin VB.Menu mDeleteGroup 
         Caption         =   "Delete Group"
      End
      Begin VB.Menu mLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mAddItem 
         Caption         =   "Add &Item"
         Shortcut        =   ^I
      End
      Begin VB.Menu mDeleteItem 
         Caption         =   "Delete Item"
      End
      Begin VB.Menu mChangeItem 
         Caption         =   "&Change Item"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      Begin VB.Menu mShowRC 
         Caption         =   "Show &RC Precompile Data"
      End
      Begin VB.Menu mShowXML 
         Caption         =   "&Show &XML Data"
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "&Tools"
      Begin VB.Menu mBrowse 
         Caption         =   "&Browse"
         Shortcut        =   ^B
      End
      Begin VB.Menu mExtract 
         Caption         =   "&Extract"
         Shortcut        =   ^E
      End
      Begin VB.Menu mClip 
         Caption         =   "&ClipBoard"
      End
   End
   Begin VB.Menu mEditors 
      Caption         =   "E&ditors"
      Begin VB.Menu mEditCursor 
         Caption         =   "&Cursor"
      End
      Begin VB.Menu mEditIcon 
         Caption         =   "&Icon"
      End
      Begin VB.Menu mEditPicture 
         Caption         =   "&Picture"
      End
      Begin VB.Menu mEditStringTable 
         Caption         =   "&String Table"
      End
   End
   Begin VB.Menu mExample 
      Caption         =   "E&xamples"
      Begin VB.Menu mEx 
         Caption         =   "Example"
         Index           =   1
      End
      Begin VB.Menu mLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mSound 
         Caption         =   "Play Sound"
      End
      Begin VB.Menu mVideo 
         Caption         =   "Play Video"
      End
   End
   Begin VB.Menu mLanguage 
      Caption         =   "&Language"
      Begin VB.Menu mLang 
         Caption         =   "Lang"
         Index           =   1
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SearchTreeForFile Lib "imagehlp" (ByVal RootPath As String, ByVal InputPathName As String, ByVal OutputPathBuffer As String) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Use registry
Private booSettings As Boolean

'Used alot
Private X As Integer
Private Y As Integer
Private txt As String

'Track Selected Tree Values
Private strdata As String
Private strName As String
Private strValue As String

'Vert & Hori Splitters Resizing
Private bytMoveNow As Byte

'Track treeNodes
Private Enum nodType
    otNone = 0
    otRoot = 1
    otData = 2
    otName = 3
End Enum
Private otSource As nodType
Private objSourceNode As Object
Private objTargetNode As Object
Private intNode As Integer
Private noDx As Node

Private Const lngEnglish = 1000
Private Const lngGerman = 2000
Private Const lngFrench = 3000
Private Const lngSpanish = 4000
Private varLanguage As Integer

Private Sub GetChildren(tvw As TreeView, nodN As Node, nodP As Node)
Dim nodC As Node, nodT As Node
Dim i As Integer

  With tvw
    'For each children in the tree
    For i = 1 To nodN.Children
      'If it's the first child:
      If i = 1 Then
        'Add the node:
        Set nodC = .Nodes.Add(nodP.Index, tvwChild, , nodN.Child.Text, 3)
            nodC.Tag = nodN.Child.Tag
        'Set us up for the next child:
        Set nodT = nodN.Child.Next
        'Get the added nodes children:
        If nodN.Child.Children <> 0 Then
          GetChildren tvw, nodN.Child, nodC
        End If
      'It's not the first child, so:
      Else
        On Error Resume Next
        'Add the node:
        Set nodC = .Nodes.Add(nodP.Index, tvwChild, , nodT.Text, 3)
            nodC.Tag = nodT.Tag
        'Get the added nodes children:
        If nodT.Children <> 0 Then
          GetChildren tvw, nodT, nodC
        End If
        'Set us up again:
        Set nodT = nodT.Next
      End If
    Next
  End With

Set nodC = Nothing
Set nodT = Nothing
End Sub

Private Sub MoveNode(tvw As TreeView, noDx As Node, Direction As String)
'Credit: Andrew Murphy
Dim nodN As Node
Dim strName As String
  
  'All we do here is copy the node and set it as the previous
  'Nodes previous node. A little confusing, but it works.
  'We then add all the children and delete the original
  'Node
  
  With tvw
    Select Case Direction
      Case "UP"
        If Not noDx.Previous Is Nothing Then
          'This modification is to allow for
          'the correct Images and Tags
          'to be included.
          If otSource = otData Then
            Set nodN = .Nodes.Add(noDx.Previous, tvwPrevious, , noDx.Text, 2)
            nodN.Tag = noDx.Tag
          Else
            Set nodN = .Nodes.Add(noDx.Previous, tvwPrevious, , noDx.Text, 3)
            nodN.Tag = noDx.Tag
          End If
          Saved = False
        Else
          Exit Sub
        End If
      Case "DOWN"
        If Not noDx.Next Is Nothing Then
          If otSource = otData Then
            Set nodN = .Nodes.Add(noDx.Next, tvwNext, , noDx.Text, 2)
            nodN.Tag = noDx.Tag
          Else
            Set nodN = .Nodes.Add(noDx.Next, tvwNext, , noDx.Text, 3)
            nodN.Tag = noDx.Tag
          End If
        Else
          Exit Sub
        End If
        Saved = False
    End Select
      
    nodN.Selected = True
    
    'Move the Child nodes, if any.
    If noDx.Children <> 0 Then
      GetChildren tvw, noDx, nodN
    End If
    
    'Delete the Old Node.
    strName = noDx.Key
    .Nodes.Remove noDx.Index
    Set noDx = Nothing
    nodN.Key = strName
  End With

Set nodN = Nothing
End Sub

Private Sub ShowXML()
Dim strTxt As String

txt = "<" & UCase$(strResource) & ">" & vbCrLf

For X = 2 To tv.Nodes.Count
  
   If tv.Nodes(X).Children > 0 Then
    txt = txt & vbTab & "<resDATA id=" & Chr$(34) & tv.Nodes(X).Text & Chr$(34) & ">" & vbCrLf

      'Get first child's text
      txt = txt & vbTab & vbTab & "<resNAME>" & tv.Nodes(X).Child.Text & "</resNAME>" & vbCrLf
        strTxt = tv.Nodes(X).Child.Tag
        strTxt = PutInserts(strTxt)
      txt = txt & vbTab & vbTab & "<resVALUE>" & strTxt & "</resVALUE>" & vbCrLf
      
      Y = tv.Nodes(X).Child.Index

      'Get next sibling's text
      While Y <> tv.Nodes(X).Child.LastSibling.Index
        txt = txt & vbTab & vbTab & "<resNAME>" & tv.Nodes(Y).Next.Text & "</resNAME>" & vbCrLf
        strTxt = tv.Nodes(Y).Next.Tag
        strTxt = PutInserts(strTxt)
        txt = txt & vbTab & vbTab & "<resVALUE>" & strTxt & "</resVALUE>" & vbCrLf
         'Set y to next sibling's index
         Y = tv.Nodes(Y).Next.Index
      Wend
   
    txt = txt & vbTab & "</resDATA>" & vbCrLf
   End If
Next

txt = txt & "</" & UCase$(strResource) & ">" & vbCrLf

txtString(2) = txt
picSplitH.Cls
picSplitH.Print "   XML Data Code"
End Sub

Private Sub cmdGroup_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sPath As String, sExT As String
Dim tvParent As Node
Dim booParent As Boolean
Dim noDx As Node

If Data.GetFormat(vbCFFiles) Then
  Y = Data.Files.Count

For X = 1 To Y
txt = Data.Files(X)

sPath = strFilePath(txt)
sExT = Right$(txt, 4)
strName = strFileName(txt)

    strdata = UCase$(sExT)
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

  booParent = True
  For Y = 1 To tv.Nodes.Count
    If strdata = UCase$(tv.Nodes(Y).Text) Then
      Set tvParent = tv.Nodes(Y)
      booParent = False
      Exit For
    End If
  Next
  If booParent Then
    Set tvParent = tv.Nodes.Add("root", tvwChild, strdata, strdata, 2)
      tvParent.Tag = strdata
  End If
  
  If Shift = vbShiftMask Then
    'Move files
    Name txt As strFilePath(strFile) & strName & sExT
  Else
    'Copy files
    FileCopy txt, strFilePath(strFile) & strName & sExT
  End If

strdata = tvParent
txtString(1) = strName & sExT
strValue = txtString(1)
strName = StrConv(strFileName(txt), vbProperCase)
txtString(0) = strName
   
   Set noDx = tv.Nodes.Add(strdata, tvwChild, , strName, 3)
     noDx.Tag = strValue
      noDx.EnsureVisible
Next
End If

otSource = otName
Saved = False
ShowPreCompile
Exit Sub
errHandler:
If Err.Number = 35602 Then
txt = "Reference Key Name already in use..." & vbCrLf & vbCrLf
txt = txt & "New Reference Key Name must be unique."
MsgBox txt, vbExclamation + vbOKOnly, "Add New Reference Key"
txtString(0).SetFocus
End If
End Sub

Private Sub cmdGroup_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'If not a file, then no dropping
If Not Data.GetFormat(vbCFFiles) Then
  Effect = vbDropEffectNone
End If
End Sub

Private Sub cmdItem_Click(Index As Integer)
Select Case Index
Case 0
  mAddItem_Click
Case 1
  mDeleteItem_Click
Case 2
  mChangeItem_Click
End Select
  txtString(0).SetFocus
End Sub


Private Sub cmdGroup_Click(Index As Integer)
'Use arrays of commandButtons
'whenever possible and, if they are
'backed up with the same menu commands
'just send a click to the menu

Select Case Index
Case 0
  mAddGroup_Click
Case 1
  mDeleteGroup_Click
End Select
End Sub


Private Sub cmdItem_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim sPath As String
Dim sExT As String
Dim noDx As Node

If otSource = otRoot Or otSource = otNone Then
  MsgBox "Select a 'Group' to insert New Reference Key.", vbExclamation + vbOKOnly, "Add New Resource Item"
  Exit Sub
End If
 
If Data.GetFormat(vbCFFiles) Then
txt = LCase$(Data.Files(1))

sPath = strFilePath(txt)
sExT = Right$(txt, 4)
strName = strFileName(txt)

  If Shift = vbShiftMask Then
    'Move files
    Name txt As strFilePath(strFile) & strName & sExT
  Else
    'Copy files
    FileCopy txt, strFilePath(strFile) & strName & sExT
  End If

txtString(1) = strName & sExT
strValue = txtString(1)
strName = StrConv(strFileName(txt), vbProperCase)
txtString(0) = strName
   
   Set noDx = tv.Nodes.Add(strdata, tvwChild, , strName, 3)
     noDx.Tag = strValue
End If

otSource = otName
Saved = False
ShowPreCompile
Exit Sub
errHandler:
If Err.Number = 35602 Then
  txt = "Reference Key Name already in use..." & vbCrLf & vbCrLf
  txt = txt & "New Reference Key Name must be unique."
  MsgBox txt, vbExclamation + vbOKOnly, "Add New Reference Key"
  txtString(0).SetFocus
Else
  MsgBox "Error adding reference key: " & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Add New Reference Key"
End If
End Sub

Private Sub cmdItem_OLEDragOver(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'If not a file, then no dropping
If Not Data.GetFormat(vbCFFiles) Then
  Effect = vbDropEffectNone
End If
End Sub


Private Sub Form_Load()
Dim imgX As ListImage

'Initialize user-defined settings
UserSettings
SetClipboardViewer Me.hWnd
If mClip.Checked Then HookForm Me

'Resize the Fonts that are not default
SetResolution
SetMargins

  Set imgX = imgTree.ListImages. _
  Add(, , LoadResPicture("Resource", vbResIcon))
  Set imgX = imgTree.ListImages. _
  Add(, , LoadResPicture("Section", vbResIcon))
  Set imgX = imgTree.ListImages. _
  Add(, , LoadResPicture("Item", vbResIcon))

tv.DragIcon = LoadResPicture("Drag", vbResIcon)
picSplitH.MouseIcon = LoadResPicture("Horisize", vbResCursor)
picSplitV.MouseIcon = LoadResPicture("Vertsize", vbResCursor)

  'Reduce the size of the idle
  'exe file by using arrays of
  'commandButtons and menus that can
  'be dynamically loaded at runtime.
  varLanguage = lngEnglish
  'Loading the Language strings
  mLang(1).Caption = LoadResString(1 + varLanguage)
  For X = 2 To 4
    Load mLang(X)
    mLang(X).Caption = LoadResString(X + varLanguage)
    mLang(X).Visible = True
  Next

  mEx(1).Caption = LoadResString(101)
  For X = 2 To 6
    Load mEx(X)
    mEx(X).Caption = LoadResString(100 + X)
    mEx(X).Visible = True
  Next

If Not CBool(PathFileExists(strFile)) Then
  strFile = GetPath & "templates\resource\resource.xml"
End If

'Load the TreeView from the XML File
LoadXML strFile
Set imgX = Nothing
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CheckSave Me, strFile, Saved

'Registry cleared?
If Not booSettings Then

  'Use previous settings if Minimized or Maximized
  If Me.WindowState = vbMinimized Then
    'Reopen as vbNormal
    SaveSetting App.Title, "Editor", "State", vbNormal
  ElseIf Me.WindowState = vbMaximized Then
    'Reopen as vbMaximized
    SaveSetting App.Title, "Editor", "State", Me.WindowState
  Else 'vbNormal
    'Save new vbNormal sizes
    SaveSetting App.Title, "Editor", "Splitv", picSplitV.Left
    SaveSetting App.Title, "Editor", "SplitH", picSplitH.Top
    SaveSetting App.Title, "Editor", "Left", Me.Left
    SaveSetting App.Title, "Editor", "Top", Me.Top
    SaveSetting App.Title, "Editor", "Width", Me.Width
    SaveSetting App.Title, "Editor", "Height", Me.Height
    SaveSetting App.Title, "Editor", "State", Me.WindowState
  End If
    
    'Reopen with the current template
    SaveSetting App.Title, "Editor", "Clip", mClip.Checked
    SaveSetting App.Title, "Editor", "Last", strFile
End If

  UnHookForm Me
End Sub

Private Sub Form_Resize()
'Small adjustments to make everything look better

'Skip Minimize event
If Me.WindowState <> 1 Then

picSplitV.Move picSplitV.Left, picSplitV.Top, _
  picSplitV.Width, Me.Height - (picSplitV.Top + 780)

picBoxLt.Move picBoxLt.Left, picBoxLt.Top, _
  picSplitV.Left - 60, Me.Height - (picBoxLt.Top + 780)
picBoxRt.Move picSplitV.Left + 120, picBoxRt.Top, _
  Me.Width - (picSplitV.Left + picSplitV.Width + 180), Me.Height - (picBoxRt.Top + 780)

tv.Move -45, cmdGroup(0).Height - 30, _
  picBoxLt.Width + 45, picBoxLt.Height - tv.Top - 15

picSplitH.Move -30, picSplitH.Top, _
  picBoxRt.Width, picSplitH.Height
txtString(0).Move -30, txtString(0).Top, _
  picBoxRt.Width, txtString(0).Height
txtString(1).Move -30, txtString(1).Top, _
  picBoxRt.Width, picSplitH.Top - txtString(1).Top
txtString(2).Move -30, picSplitH.Top + 360, _
  picBoxRt.Width, picBoxRt.Height - (picSplitH.Top + picSplitH.Height + 30)

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim frm As Form, ctl As Control
For Each frm In Forms
  For Each ctl In frm.Controls
    Set ctl = Nothing
  Next ctl
    Unload frm
    Set frm = Nothing
Next frm
End Sub


Private Sub mAddGroup_Click()
On Error GoTo errHandler

  txt = "Add New Group to Resource file" & vbCrLf
  txt = txt & "It must be unique..."
  txt = InputBox(txt, "Add New Group", "NewGroup")
  If txt = "" Or txt = "NewGroup" Then Exit Sub
  
  If IsNumeric(txt) Then txt = "#" & txt
  strdata = UCase$(Trim$(txt))

Select Case strdata
  Case "CURSOR"
    strdata = "CURSORS"
  Case "BITMAP"
    strdata = "BITMAPS"
  Case "ICON"
    strdata = "ICONS"
  Case "STRING"
    strdata = "STRINGS"
  Case "AVI"
    strdata = "VIDEOS"
  Case "WAV"
    strdata = "SOUNDS"
End Select

  Set noDx = tv.Nodes.Add("root", tvwChild, strdata, strdata, 2)
    noDx.Tag = txt
  
  noDx.EnsureVisible
  tv_NodeClick noDx

Saved = False
ShowPreCompile
Exit Sub
errHandler:
If Err.Number = 35602 Then
txt = "Identifier already in use..." & vbCrLf & vbCrLf
txt = txt & "New Identifier must be unique."
MsgBox txt, vbExclamation + vbOKOnly, "Add New Group"
Err.Clear
Else
Resume Next
End If
End Sub

Private Sub mAddItem_Click()
On Error GoTo errHandler

If otSource = otRoot Or otSource = otNone Then
  MsgBox "Select a 'Group' to insert New Reference Key.", vbExclamation + vbOKOnly, "Add New Resource Item"
  Exit Sub
End If
 
 If txtString(0) = "" Then
   MsgBox "Enter a Reference Name or Number.", vbExclamation Or vbOKOnly, "Add New Resource Item"
   txtString(0).SetFocus
   Exit Sub
 End If
 
 If txtString(1) = "" Then
   MsgBox "Enter a Reference Value.", vbExclamation Or vbOKOnly, "Add New Resource Item"
   txtString(1).SetFocus
   Exit Sub
 End If
 
 If InStr(txtString(1), ":\") Then
   If Not InStr(txtString(1), ":\\") Then
   MsgBox "In the Resource file," & vbCrLf & vbCrLf & "Qualified-Path Values must have a Double Slash (ie: 'C:\\vb') after the Drive.   ", vbInformation Or vbOKOnly, "Add New Resource Item"
   txtString(1).SetFocus
   Exit Sub
   End If
 End If
 
   strName = txtString(0)
   strValue = txtString(1)
   Set noDx = tv.Nodes.Add(strdata, tvwChild, , strName, 3)
     noDx.Tag = strValue

     noDx.EnsureVisible
     tv_NodeClick noDx

otSource = otName
Saved = False
ShowPreCompile
Exit Sub
errHandler:
If Err.Number = 35602 Then
txt = "Reference Key Name already in use..." & vbCrLf & vbCrLf
txt = txt & "New Reference Key Name must be unique."
MsgBox txt, vbExclamation + vbOKOnly, "Add New Reference Key"
Err.Clear
txtString(0).SetFocus
Else
Err.Clear
End If
End Sub

Private Sub mBrowse_Click()
 frmBrowse.Show vbModal
 ShowPreCompile
End Sub

Private Sub mChangeItem_Click()
'On Error GoTo errHandler

Dim txt As String
Dim i


If otSource = otName Then
  If txtString(0) = "" Then
    MsgBox "Select a 'Reference Key' in the Tree.", _
      vbExclamation + vbOKOnly, "Change Reference Item"
    Exit Sub
  End If

  If InStr(txtString(1), ":\") Then
   If Not InStr(txtString(1), ":\\") Then
   MsgBox "In the Resource file," & vbCrLf & vbCrLf & "Qualified-Path Values must have a Double Slash (ie: 'C:\\vb') after the Drive.   ", vbInformation Or vbOKOnly, "Add New Resource Item"
   txtString(1).SetFocus
   Exit Sub
   End If
  End If
 
  strName = txtString(0)
  strValue = txtString(1)
  tv.SelectedItem.Text = strName
  tv.SelectedItem.Tag = strValue
    Saved = False

End If

ShowPreCompile
Exit Sub
errHandler:
txt = "Error #" & Err.Number & vbCrLf
txt = txt & Err.Description & vbCrLf & vbCrLf
MsgBox txt, vbExclamation + vbOKOnly, "Error - Change Reference Key"
Err.Clear
txtString(0).SetFocus
End Sub

Private Sub mClip_Click()
  mClip.Checked = Not mClip.Checked
  SaveSetting App.Title, "Editor", "Clip", False
  If mClip.Checked Then
    HookForm Me
  Else
    UnHookForm Me
  End If
End Sub

Private Sub mCompile_Click()
On Error Resume Next
Dim objBrowse As Object
Dim fPath As Variant
Dim F As Long
Dim lngRet As Long
Dim intTime As Integer
Dim strPath As String
Dim strTmp As String
Dim newFolder As String
Dim resPath As String
Dim strShortFile As String * 128

  'Parse to get path w/o filename
  strPath = strFilePath(strFile)

'Now, are rc.exe & rcdll.dll in
'this folder? Note: this is an
'excellent file search routine.
If CBool(PathFileExists(GetPath & "rc.exe")) = False Then
  If MsgBox("This program requires rc.exe and rcdll.dll" & vbCrLf _
    & "be located in this folder." & vbCrLf & vbCrLf _
    & "Move the essential files to this folder?", vbYesNo Or vbCritical, "Compiler EXEs not found") = vbYes Then
    Screen.MousePointer = vbHourglass
    txt = String$(MAX_PATH, Chr$(0))
    'Chars c to z
    For X = 99 To 122
      'Make drive path
      resPath = Chr$(X) & ":\"
      If CBool(PathIsDirectory(resPath)) Then
      F = SearchTreeForFile(resPath, "rc.exe", txt)
        'Found
        If F = 1 Then
          'Remove nulls
          resPath = Left$(strTmp, InStr(strTmp, Chr$(0)) - 1)
          FileCopy resPath, GetPath & "rc.exe"
          FileCopy resPath, GetPath & "rcdll.dll"
          Exit For
        End If
      End If
    Next X 'Drive
    Screen.MousePointer = vbDefault
    If F = 0 Then
      MsgBox "Files not found.", , "Compile Canceled"
      Exit Sub
    End If
  End If
Else
  Exit Sub
End If

If Saved = False Then
  If MsgBox("Save Changes to '" & strResource & "' File", vbYesNo Or vbQuestion, "File has Changed") = vbYes Then
    SaveXML
  End If
End If

  'Refresh Resource File
  ShowPreCompile
  txt = txtString(2)

  'Save unCompiled '.rc' in the template folder
  Call binSave(strPath & strResource & ".rc", txt)

  'Get User's Project Folder
  'Use the webBrowser's BrowseForFolder object and NameSpaces
  'See: http://www.microsoft.com/mind/0898/cutting0898.asp
  Set objBrowse = CreateObject("Shell.Application")
  Set fPath = objBrowse.BrowseForFolder(frmBrowse.hWnd, _
    "Select the     Project 's 'Folder' to Save Resource (.res) File", _
    0, Left$(App.Path, 3))
  If fPath.Items.Item.Path = "" Then
    cmdItem(0).SetFocus
    Me.Caption = App.Title
    Exit Sub
  Else
    Me.Caption = "Compiling..."
    newFolder = fPath.Items.Item.Path & "\"
  End If

  'Put all necessary files in the same folder
  'Compile them into '.res', then delete
  'them all, except the '.res' file.
  'By putting everything together in the template
  'folder, this compile procedure has never failed.
  FileCopy GetPath & "rc.exe", strPath & "rc.exe"
  FileCopy GetPath & "rcdll.dll", strPath & "rcdll.dll"
  
  'Delete the old res - this is essential,
  'the batch file won't overwrite it
  If CBool(PathFileExists(newFolder & strResource & ".res")) _
    Then DeleteFile newFolder & strResource & ".res"
  'Move new resource file to selected folder
  'This will error out until the .res is moved
  F = GetTickCount
  Do Until CBool(PathFileExists(newFolder & strResource & ".res")) = False
    DeleteFile newFolder & strResource & ".res"
    'Just in case - give it 15 secs
    If (GetTickCount - F) > 15000 Then
      MsgBox "Can't remove old resource file..." & vbCrLf & vbCrLf _
        & "Is the file Locked (Read Only) or Opened (Being Used)?", "Resource Editor - Compile Error"
      Me.Caption = App.Title
      cmdItem(0).SetFocus
      Exit Sub
    End If
  Loop
  
  'The resource compler - rc.exe won't run from vb
  'It must be run from a batch file,
  'which requires the short path name
  lngRet = GetShortPathName(strPath, strShortFile, Len(strShortFile))
  resPath = Left(strShortFile, lngRet)
  
  'Build the dos batch file to
  'compile the resource file
  txt = "ChDir " & resPath & vbCrLf & "CALL RC.exe /r " & strResource & ".rc"
  Call binSave(strPath & "Compile.bat", txt)

  'vbNormalFocus will show any compilation errors
  lngRet = Shell(strPath & "Compile.bat", vbHide)
  
  'Move new resource file to selected folder
  'This will error out until the .res is moved
  F = GetTickCount
  Do Until CBool(PathFileExists(newFolder & strResource & ".res")) = True
  Name strPath & strResource & ".res" As newFolder & strResource & ".res"
    'Just in case - give it 15 secs
    If (GetTickCount - F) > 15000 Then
      MsgBox "  Can't create new resource file  ", vbCritical, "Compile Error - Resource Editor"
      Exit Do
    End If
  Loop
  
  Me.Caption = App.Title
  cmdItem(0).SetFocus
  If CBool(PathFileExists(newFolder & strResource & ".res")) = True Then
  MsgBox vbCrLf & newFolder & strResource & ".res" & vbCrLf & vbCrLf & _
    vbTab & "Date:" & vbTab & Date & vbCrLf & _
    vbTab & "Time:" & vbTab & Time, , App.Title & " - Compiled:  " & strResource
  End If
  
  'Remove extra files used for compilation
  DeleteFile strPath & "Compile.bat"
  DeleteFile strPath & strResource & ".rc"
  DeleteFile strPath & "rcdll.dll"
  DeleteFile strPath & "rc.exe"

Set objBrowse = Nothing
End Sub

Private Sub mDeleteItem_Click()
On Error GoTo errHandler

If otSource = otName Then
    If MsgBox("Confirm to delete Item '" & tv.SelectedItem.Text & _
        "'", vbYesNo, "Delete Item") = vbYes Then
      tv.Nodes.Remove tv.SelectedItem.Index
      txtString(0) = ""
      txtString(1) = ""
      txtString(0).SetFocus
        Saved = False
    End If
End If

ShowPreCompile
Exit Sub
errHandler:
txt = "Error #" & Err.Number & vbCrLf
txt = txt & Err.Description & vbCrLf & vbCrLf
MsgBox txt, vbExclamation + vbOKOnly, "Error - Delete Reference Item"
Err.Clear
End Sub

Private Sub mDeleteGroup_Click()
On Error GoTo errHandler

If otSource = otData Then
  If MsgBox("Confirm Delete Group '" & tv.SelectedItem.Text & _
      "'", vbYesNo, "Delete Group") = vbYes Then
      tv.Nodes.Remove tv.SelectedItem.Key
      tv.Nodes(1).Selected = True
          Saved = False
  End If
Else
  MsgBox "Select a Group to Delete", vbOKOnly Or vbInformation, "Delete Group"
End If

ShowPreCompile
Exit Sub
errHandler:
txt = "Error #" & Err.Number & vbCrLf
txt = txt & Err.Description & vbCrLf & vbCrLf
MsgBox txt, vbExclamation + vbOKOnly, "Error - Delete Group"
Err.Clear
End Sub

Private Sub mEditCursor_Click()
'Imagedit.exe is on the vb cd
On Error Resume Next
  txt = GetSetting(App.Title, "Editor", "EditCursor", "Empty")
  
  If CBool(PathFileExists(txt)) Then
    Shell txt, vbNormalFocus
  Else
    With CommonDialog1
     .DialogTitle = "Locate Cursor Editor"
     .filename = "CursorEditor"
     .Flags = cdlOFNFileMustExist
     .Filter = "Executable Files (*.exe)|*.exe"
     .InitDir = Left$(App.Path, 3)
     .ShowOpen
     
     If Err Then Exit Sub
     SaveSetting App.Title, "Editor", "EditCursor", CommonDialog1.filename
     Shell .filename, vbNormalFocus
    End With
  End If
End Sub

Private Sub mEditIcon_Click()
'http://www.pcmag.com/' for IconEdit32
On Error Resume Next
  txt = GetSetting(App.Title, "Editor", "EditIcon", "Empty")
  
  If CBool(PathFileExists(txt)) Then
    Shell txt, vbNormalFocus
  Else
    With CommonDialog1
     .DialogTitle = "Locate Icon Editor"
     .filename = "IconEditor"
     .Flags = cdlOFNFileMustExist
     .Filter = "Executable Files (*.exe)|*.exe"
     .InitDir = Left$(App.Path, 3)
     .ShowOpen
     
     If Err Then Exit Sub
     SaveSetting App.Title, "Editor", "EditIcon", CommonDialog1.filename
     Shell .filename, vbNormalFocus
    End With
  End If
End Sub


Private Sub mEditPicture_Click()
On Error Resume Next
  txt = GetSetting(App.Title, "Editor", "EditPicture", "Empty")
  
  If CBool(PathFileExists(txt)) Then
    Shell txt, vbNormalFocus
  Else
    With CommonDialog1
     .DialogTitle = "Locate Graphics Editor"
     .filename = "GraphicsEditor"
     .Flags = cdlOFNFileMustExist
     .Filter = "Executable Files (*.exe)|*.exe"
     .InitDir = Left$(App.Path, 3)
     .ShowOpen
     
     If Err Then Exit Sub
     SaveSetting App.Title, "Editor", "EditPicture", CommonDialog1.filename
     Shell .filename, vbNormalFocus
    End With
  End If
End Sub

Private Sub mEditStringTable_Click()
  frmStrings.Show
  ShowPreCompile
End Sub

Private Sub mExCustom_Click()
  strSample = "Custom"
  frmSample.Show vbModal
End Sub

Private Sub mExGen_Click()
  strSample = "General"
  frmSample.Show vbModal
End Sub

Private Sub mEx_Click(Index As Integer)
  strSample = mEx(Index).Caption
  frmSample.Show vbModal
End Sub

Private Sub mExit_Click()
  Unload Me
End Sub

Private Sub mExtract_Click()
'It is very easy to add a resource
'extraction ultility here.  Just
'examine the code from frmBrowse and
'frmStrings to see how.  The concept
'is to 1) check the treeview for
'duplicates, 2) add to the treeview,
'3) move the resource data files to
'the proper templates folder.
'Many extractors on PSC will access
'more resources than vb can ever use
'and can be adapted like frmstrings
'or frmbrowse.

'Otherwise, get Angus Johnson's at:
' http://www.users.on.net/johnson/resourcehacker/
'then extract and save resource to working template
'folder, and add it to the treeview via frmBrowse
On Error Resume Next
  txt = GetSetting(App.Title, "Editor", "resHack", "Empty")
  
  If CBool(PathFileExists(txt)) Then
    Shell txt, vbNormalFocus
  Else
    CommonDialog1.DialogTitle = "Locate Resource Extractor"
    CommonDialog1.filename = "ResExtractor"
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Filter = "Executable Files (*.exe)|*.exe"
    CommonDialog1.InitDir = Left$(App.Path, 3)
    CommonDialog1.ShowOpen
    If Err Then Exit Sub
    SaveSetting App.Title, "Editor", "resHack", CommonDialog1.filename
    Shell CommonDialog1.filename, vbNormalFocus
  End If
End Sub

Private Sub mHelp_Click()
Dim F As Long
F = ShellExecute(Me.hWnd, vbNullString, "ResourceEditor.chm", vbNullString, App.Path, SW_SHOWNORMAL)
End Sub

Private Sub mImport_Click()
With CommonDialog1
  .DialogTitle = "Import XML Resource"
  .filename = "xmlResource.xml"
  .Flags = cdlOFNFileMustExist
  .Filter = "XML Files (*.xml)|*.xml"
  .InitDir = strFilePath(strFile)
  .ShowOpen
  txt = Trim$(.filename)
End With
If Not Len(txt) Then Exit Sub

Dim xmlDoc As DOMDocument
Dim objResource As IXMLDOMElement
Dim objData As IXMLDOMElement
Dim objAttributes As IXMLDOMNamedNodeMap
Dim objAttributeNode As IXMLDOMNode
Dim objDOMElement As IXMLDOMElement
Dim tvParent As Node
Dim booParent As Boolean
  
  Set xmlDoc = New DOMDocument
  xmlDoc.async = False
  xmlDoc.validateOnParse = True
  xmlDoc.resolveExternals = True

  xmlDoc.Load txt
  
  If xmlDoc.parseError.reason <> "" Then
    MsgBox xmlDoc.parseError.reason
    Exit Sub
  End If

  Set objResource = xmlDoc.documentElement

For Each objData In objResource.childNodes
booParent = True

  Set objAttributes = objData.attributes
  Set objAttributeNode = objAttributes.getNamedItem("id")
      For X = 1 To tv.Nodes.Count
        If objAttributeNode.nodeValue = tv.Nodes(X).Text Then
          Set tvParent = tv.Nodes(X)
          booParent = False
          Exit For
        End If
      Next
  If booParent Then
    Set tvParent = tv.Nodes.Add("root", tvwChild, objAttributeNode.nodeValue, objAttributeNode.nodeValue, 2)
      tvParent.Tag = objAttributeNode.nodeValue
      tvParent.EnsureVisible
  End If
  
  For Each objDOMElement In objData.childNodes
    Select Case objDOMElement.nodeName
    Case "resNAME"
      booParent = True
      For X = 1 To tv.Nodes.Count
        If objDOMElement.nodeTypedValue = tv.Nodes(X).Text Then
          booParent = False
          Exit For
        End If
      Next
      If booParent Then
        Set noDx = tv.Nodes.Add(tvParent, tvwChild, , objDOMElement.nodeTypedValue, 3)
      End If
    Case "resVALUE"
      If booParent Then
          'If its a file Copy it to template folder
          If CBool(PathFileExists(objDOMElement.nodeTypedValue)) Then
            FileCopy objDOMElement.nodeTypedValue, GetPath & "templates\" & strResource & "\" & strFileName(objDOMElement.nodeTypedValue)
            noDx.Tag = strFileName(objDOMElement.nodeTypedValue)
          ElseIf CBool(PathFileExists(strFilePath(txt) & objDOMElement.nodeTypedValue)) Then
            FileCopy strFilePath(txt) & objDOMElement.nodeTypedValue, GetPath & "templates\" & strResource & "\" & objDOMElement.nodeTypedValue
            noDx.Tag = objDOMElement.nodeTypedValue
          Else
            noDx.Tag = GetInserts(objDOMElement.nodeTypedValue)
          End If
      End If
    End Select
  Next
  
Next

Saved = False
ShowXML
  Set objResource = Nothing
  Set objData = Nothing
  Set objAttributes = Nothing
  Set objAttributeNode = Nothing
  Set objDOMElement = Nothing
  Set tvParent = Nothing
  Set noDx = Nothing
  Set xmlDoc = Nothing
End Sub

Private Sub mLang_Click(Index As Integer)
  varLanguage = Index * 1000

  For X = 1 To 4
    mLang(X).Caption = LoadResString(X + varLanguage)
  Next

End Sub

Private Sub mNew_Click()
CheckSave Me, strFile, Saved

  strResource = "NEWRESOURCE"
  tv.Nodes.Clear
  Set noDx = tv.Nodes.Add(, , "root", strResource, 1)
  
  strFile = GetPath & "templates\" & strResource & "\" & strResource & ".xml"
  
  Saved = False
End Sub

Private Sub mOpen_Click()
With CommonDialog1
  .DialogTitle = "Open XML Resource"
  .filename = "xmlResource.xml"
  .Flags = cdlOFNFileMustExist
  .Filter = "XML Files (*.xml)|*.xml"
  .InitDir = strFilePath(strFile)
  .ShowOpen
  txt = .filename
End With
  If Not Len(txt) Then Exit Sub

  CheckSave Me, strFile, Saved
  
  strFile = txt
  LoadXML strFile
End Sub

Private Sub mPrtRC_Click()
  PrintCode txtString(2)
End Sub

Private Sub mPrtXML_Click()
  txt = binOpen(strFile)

PrintCode txt
End Sub

Private Sub mSave_Click()
  SaveXML
End Sub

Private Sub mSaveAs_Click()
CheckSave Me, strFile, Saved

Dim newResource As String
Dim newName As String
Dim newFile As String

'Do not to use the inputbox if
'the tv.root label has been changed
    newName = UCase$(tv.Nodes(1).Text)
Select Case newName
  'Is the root label the 'default' or
  'was the label changed...
  Case "RESOURCE", strResource
    'Then get new name via inputbox
    GoTo here
  
  'If Edited tv.root.label using menu-File|New
  Case Else
    'Confirm changedLabel as new name...
    If MsgBox("Save As:  " & newName & "   ", vbYesNo Or vbQuestion, "Save New Template As") = vbYes Then
      'Skip inputbox
      GoTo there
    Else
      'Get new name via inputbox
      GoTo here
    End If
End Select

here:
  'Get new Name
  newName = InputBox("Create New Resource File", "New Resource file - Resource Editor", "NewResourceFileName")
  'Can't use default
  If UCase$(newName) = "RESOURCE" Then
    MsgBox " 'RESOURCE' is Reserved as the Default Template  "
    Exit Sub
  
  'Cancelled or...
  ElseIf newName = "" Or newName = "NewResourceFileName" Then
    Exit Sub
  End If
  
there:
  'Fix the entry as new resource name and file
  If Right$(LCase$(newName), 4) = ".xml" Then
     newResource = Left$(newFile, Len(newFile) - 4)
   newFile = GetPath & "templates\" & newResource & "\" & newName
  Else
      newResource = newName
    newFile = GetPath & "templates\" & newResource & "\" & newName & ".xml"
  End If
  
  'Does the new template file already exist???
  If CBool(PathFileExists(newFile)) Then
    'If so, don't overwrite it???
    If MsgBox("OverWrite Existing Template  '" & newResource & "'   ", vbYesNo Or vbQuestion, "Template Exists") = vbNo Then
      'No means start over
      Exit Sub
    End If
  End If

  'Create new Template folder
  SaveALL UCase$(newResource), newFile
  strResource = UCase$(newResource)
  strFile = newFile
  
  'If menu - File|New, make it UCase
  tv.Nodes(1).Text = strResource
  
End Sub


Private Sub mShowRC_Click()
  ShowPreCompile
End Sub

Private Sub mShowXML_Click()
  ShowXML
End Sub

Private Sub mSound_Click()
  sndArray = LoadResData("Chimes", "SOUND")
  
X = waveOutGetNumDevs()

If X > 0 Then
  X = sndPlaySound(sndArray(0), SND_SYNC Or SND_NODEFAULT Or SND_MEMORY)
Else
  PlaySound
  MsgBox "No Sound Cards Installed"
End If
End Sub

Private Sub mUnInstall_Click()
On Error Resume Next
'Clear the Registry of all settings
'made by this application until next use
If MsgBox("Delete all Registry Settings for Resource Editor", vbOKCancel + vbCritical, "UnInstall Resource Editor") = vbOK Then
  DeleteSetting App.Title
  booSettings = True
End If
End Sub

Private Sub mVideo_Click()
  frmVideo.Show vbModal
End Sub

Private Sub picSplitH_DblClick()
  ShowPreCompile
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bytMoveNow = 1
End Sub


Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bytMoveNow Then
picSplitH.Move picSplitH.Left, picSplitH.Top + Y, _
  picBoxRt.Width - 60, picSplitH.Height

If picSplitH.Top < txtString(1).Top + 900 _
  Then picSplitH.Top = txtString(1).Top + 900
If picSplitH.Top > (picBoxRt.Height - 450) _
  Then picSplitH.Top = (picBoxRt.Height - 450)

txtString(1).Move txtString(1).Left, txtString(1).Top, _
  txtString(0).Width, picSplitH.Top - txtString(1).Top
txtString(2).Move txtString(2).Left, picSplitH.Top + 360, _
  picBoxRt.Width - 60, picBoxRt.Height - (picSplitH.Top + picSplitH.Height + 30)

End If
End Sub


Private Sub picSplitH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bytMoveNow = 0
End Sub


Private Sub picSplitv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bytMoveNow = 1
End Sub


Private Sub picSplitv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bytMoveNow Then
picSplitV.Move picSplitV.Left + X, picSplitV.Top, picSplitV.Width, picSplitV.Height
  
If picSplitV.Left < 1200 Then picSplitV.Left = 1200
If picSplitV.Left > (Me.Width - 3000) Then picSplitV.Left = (Me.Width - 3000)

picBoxLt.Move picBoxLt.Left, picBoxLt.Top, picSplitV.Left - 60, picBoxLt.Height
picBoxRt.Move picSplitV.Left + 120, picBoxRt.Top, _
  Me.Width - (picSplitV.Left + picSplitV.Width + 180), picBoxRt.Height

tv.Move -45, tv.Top, picBoxLt.Width + 45, tv.Height

txtString(0).Move txtString(0).Left, txtString(0).Top, picBoxRt.Width - 60, txtString(0).Height
txtString(1).Move txtString(1).Left, txtString(1).Top, picBoxRt.Width - 60, txtString(1).Height
picSplitH.Move picSplitH.Left, picSplitH.Top, picBoxRt.Width - 60, picSplitH.Height
txtString(2).Move txtString(2).Left, txtString(2).Top, _
  picBoxRt.Width - 60, picBoxRt.Height - (picSplitH.Top + picSplitH.Height)

End If
End Sub

Private Sub picSplitv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bytMoveNow = 0
End Sub


Private Sub tv_AfterLabelEdit(Cancel As Integer, NewString As String)
  If Len(NewString) = 0 Or UCase$(NewString) = "RESOURCE" Then
     Cancel = True
     Exit Sub
  End If
  If otSource = otData Then
   NewString = UCase$(NewString)
   strdata = NewString
  
    Select Case strdata
      Case "CURSOR"
        strdata = "CURSORS"
      Case "BITMAP"
        strdata = "BITMAPS"
      Case "ICON"
        strdata = "ICONS"
      Case "STRING"
        strdata = "STRINGS"
      Case "AVI"
        strdata = "VIDEOS"
      Case "WAV"
        strdata = "SOUNDS"
    End Select
    
    tv.SelectedItem.Tag = strdata
    Saved = False
  ElseIf otSource = otName Then
   strName = NewString
    txtString(0) = strName
    Saved = False
  Else
   Cancel = True
  End If
  
'Changes to treeview control are not updated while a
'treeview event is in progress.
'So, ShowPreCompile here will not see the
'changes to show them in the preview.
End Sub

Private Sub tv_BeforeLabelEdit(Cancel As Integer)
  If otSource = otRoot Then
    'Use mNew or mSaveAs to change
    'Root name
    Cancel = True
  End If
End Sub


Private Sub tv_Collapse(ByVal Node As MSComctlLib.Node)
  If Node.Index = 1 Then
    Node.Expanded = True
  End If
End Sub

Private Sub tv_DragDrop(Source As Control, X As Single, Y As Single)
'Credit DragDrop: srinivas
Dim i As Integer

  If Not (tv.DropHighlight Is Nothing) Then
    Set objSourceNode.Parent = tv.DropHighlight
    Set tv.DropHighlight = Nothing
    strdata = tv.SelectedItem.Parent
    Saved = False
  End If
End Sub

Private Sub tv_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Dim target As Node
Dim highlight As Boolean

  'See what node we're above.
  Set target = tv.HitTest(X, Y)
  
  'Only DragDrop items to different Groups.
  'Use the keyboard to move
  'items up & down within a group or
  'to move Groups Up & down
  If target Is objTargetNode Then Exit Sub
  Set objTargetNode = target
  
  highlight = False
  If Not (objTargetNode Is Nothing) Then
      'See what kind of node were above.
      If NodeType(objTargetNode) = otData Then highlight = True
  End If
  
  If highlight Then
      Set tv.DropHighlight = objTargetNode
  Else
      Set tv.DropHighlight = Nothing
  End If
End Sub

Private Function NodeType(test_node As Node) As nodType
'Got this from someone's ini-2-tv code
'Credit: Unknown
  If test_node Is Nothing Then
      NodeType = otNone
  Else
      If test_node.Key = "root" Then
          NodeType = otRoot
      ElseIf test_node.Parent.Key = "root" Then
          NodeType = otData
      Else
          NodeType = otName
      End If
  End If
End Function


Private Sub tv_KeyDown(KeyCode As Integer, Shift As Integer)
  
'Move selection with keyboard
If KeyCode = vbKeyUp Then
'This moves Groups better than DragDrop
'But use dragdrop to move items to
'different groups
If Shift And vbCtrlMask Then
  'Move the node
  MoveNode tv, tv.SelectedItem, "UP"
  ShowPreCompile
  Saved = False
  Exit Sub
End If
  'Show data if an item is selected
  'Clear textBoxes if Group is selected
  If otSource = otName Then
    'Knowing when addressing a Group
    'is the only reason for making
    'the Group Nodes' Text and Tag equal
    If tv.Nodes(intNode).Parent.Text = tv.Nodes(intNode).Parent.Tag Then
      otSource = otData
      txtString(0) = ""
      txtString(1) = ""
    Else
      strdata = tv.Nodes(intNode).Parent.Text
      strName = tv.Nodes(intNode).Previous.Text
      strValue = tv.Nodes(intNode).Previous.Tag
        txtString(0) = strName
        txtString(1) = strValue
    End If
  ElseIf otSource = otData Then
    If Not tv.Nodes(strdata).Previous Is Nothing Then
      If tv.Nodes(strdata).Previous.Expanded Then
      otSource = otName
        strdata = tv.Nodes(strdata).Previous.Text
        strName = tv.Nodes(strdata).Child.Text
        strName = tv.Nodes(intNode).LastSibling.Text
        strValue = tv.Nodes(intNode).LastSibling.Tag
          txtString(0) = strName
          txtString(1) = strValue
      Else
        strdata = tv.Nodes(strdata).Previous.Text
      End If
    Else
      otSource = otRoot
      strdata = tv.Nodes(strdata).Parent.Child.Text
    End If
  End If
  Saved = False


ElseIf KeyCode = vbKeyDown Then
If Shift And vbCtrlMask Then
  MoveNode tv, tv.SelectedItem, "DOWN"
  ShowPreCompile
  Saved = False
  Exit Sub
End If
  If otSource = otName Then
    'This pops an error when addressing a Group node
    'or the last node.
    'On Error Resume Next gives the
    'desired results
    On Error GoTo errHandler
    If tv.Nodes(intNode).Next.Text = tv.Nodes(intNode).Next.Tag Then
      otSource = otData
      strdata = tv.Nodes(strdata).Next.Text
      txtString(0) = ""
      txtString(1) = ""
    Else
      strdata = tv.Nodes(intNode).Parent.Text
      strName = tv.Nodes(intNode).Next.Text
      strValue = tv.Nodes(intNode).Next.Tag
        txtString(0) = strName
        txtString(1) = strValue
    End If
    'if at the end of the treeview,
    'textboxes will toggle showing and
    'clearing data, this stops the toggling
errHandler:
    'At the end of the list
    If Err = 76 Then
      strdata = tv.Nodes(intNode).Parent.Text
      strName = tv.Nodes(intNode).Next.Text
      strValue = tv.Nodes(intNode).Next.Tag
        txtString(0) = strName
        txtString(1) = strValue
    End If
    'Disable Error Handler
    On Error GoTo 0
  ElseIf otSource = otData Then
    If Not tv.Nodes(strdata).Child Is Nothing Then
      If tv.Nodes(strdata).Expanded Then
      otSource = otName
      
        strName = tv.Nodes(strdata).Child.Text
        strValue = tv.Nodes(strdata).Child.Tag
          txtString(0) = strName
          txtString(1) = strValue
      Else
        If tv.Nodes(strdata).Next Is Nothing Then Exit Sub
        strdata = tv.Nodes(strdata).Next.Text
      End If
    Else
      strdata = tv.Nodes(strdata).Next.Text
    End If
  ElseIf otSource = otRoot Then
      otSource = otData
      strdata = tv.Nodes(strdata).Parent.Child.Text
  End If
  Saved = False

End If
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Prepare for a DragDrop
Set objSourceNode = tv.HitTest(X, Y)
otSource = NodeType(objSourceNode)

'This updates the textboxes
'if dragged and Dropped
If otSource = otName Then
  strdata = objSourceNode.Parent.Text
  strName = objSourceNode.Text
  strValue = objSourceNode.Tag
    txtString(0) = strName
    txtString(1) = strValue
ElseIf otSource = otData Then
  strdata = objSourceNode.Text
    txtString(0) = ""
    txtString(1) = ""
End If
End Sub

Private Sub tv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Don't DragDrop Groups (get a Group under and item)
'only items
If Button = vbLeftButton And Not otSource = otData Then
  otSource = NodeType(objSourceNode)
  tv.Drag vbBeginDrag
End If
End Sub

Public Sub SaveXML()
'I first tried to used msxml 4.0
'but the indent property in the writer
'never indented, so I wrote this
'and referenced msxml 3.0
'This is unnecessary, and only used to
'format the file cosmetically for viewing
'in txtString(2) or a text editor like notepad.
'4.0 will work for this prog. Get 4.0 at:
'http://msdn.microsoft.com/downloads/default.asp?url=/downloads/sample.asp?url=/msdn-files/027/001/766/msdncompositedoc.xml
Dim z As Integer
Dim sPath As String
Dim sName As String
Dim booFile As Boolean
Dim strTxt As String
Dim arrFile() As String

If CBool(PathFileExists(strFile)) Then DeleteFile strFile
sPath = strFilePath(strFile)

'Some formatting for consistency
txt = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "ISO-8859-1" & Chr$(34) & "?>" & vbCrLf
txt = txt & "<!-- ********** " & strResource & " Template ********** -->" & vbCrLf
txt = txt & "<" & UCase$(strResource) & ">" & vbCrLf

For X = 2 To tv.Nodes.Count
  
   'Check if Group
   If tv.Nodes(X).Children > 0 Then
    'Add Group Name
    txt = txt & vbTab & "<resDATA id=" & Chr$(34) & tv.Nodes(X).Text & Chr$(34) & ">" & vbCrLf

      'Add first child's data
      txt = txt & vbTab & vbTab & "<resNAME>" & tv.Nodes(X).Child.Text & "</resNAME>" & vbCrLf
        strTxt = tv.Nodes(X).Child.Tag
        strTxt = PutInserts(strTxt)
      txt = txt & vbTab & vbTab & "<resVALUE>" & strTxt & "</resVALUE>" & vbCrLf
          'If this is a file, add it to array
          If CBool(PathFileExists(sPath & strTxt)) Then
            ReDim Preserve arrFile(z)
            arrFile(z) = strTxt
            z = z + 1
          End If
      
      Y = tv.Nodes(X).Child.Index

      'Add sibling's data
      While Y <> tv.Nodes(X).Child.LastSibling.Index
        txt = txt & vbTab & vbTab & "<resNAME>" & tv.Nodes(Y).Next.Text & "</resNAME>" & vbCrLf
        strTxt = tv.Nodes(Y).Next.Tag
        strTxt = PutInserts(strTxt)
        txt = txt & vbTab & vbTab & "<resVALUE>" & strTxt & "</resVALUE>" & vbCrLf
          'If this is a file, add it to array
          If CBool(PathFileExists(sPath & strTxt)) Then
            ReDim Preserve arrFile(z)
            arrFile(z) = strTxt
            z = z + 1
          End If
         ' Reset y to next sibling's index.
         Y = tv.Nodes(Y).Next.Index
      Wend
   
    txt = txt & vbTab & "</resDATA>" & vbCrLf
   End If
Next

txt = txt & "</" & UCase$(strResource) & ">"

'Create a temp file
If Not CBool(PathIsDirectory(strFilePath(strFile))) Then
  MkDir strFilePath(strFile)
End If
binSave strFile, txt

'Check file array for references
'that have been deleted
'Check each file in folder (sPath)...
sName = Dir(sPath, vbNormal)
Do While sName <> ""
  'Against each file in array
  For X = LBound(arrFile) To UBound(arrFile)
    'If a match...
    If LCase$(sName) = LCase$(arrFile(X)) Then
      'Mark it
      booFile = True
    End If
  Next
  'If NOT marked...
  If booFile = False And Right$(sName, 3) <> "xml" Then
      'Delete it
      DeleteFile sPath & sName
  End If
  'Reset marker
  booFile = False
  'Check next file
  sName = Dir
Loop

Saved = True
  txtString(2) = txt
  picSplitH.Cls
  picSplitH.Print "   XML Data - DblClick for Pre-Compiled (rc) Code"
End Sub

Public Sub SaveALL(ByVal nResource As String, ByVal nFile As String)
On Error Resume Next
Dim nPath As String
Dim sPath As String
Dim strTxt As String

nPath = strFilePath(nFile)
If Not CBool(PathIsDirectory(nPath)) Then
  MkDir nPath
End If
sPath = strFilePath(strFile)

'Some formatting for consistency
txt = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "ISO-8859-1" & Chr$(34) & "?>" & vbCrLf
txt = txt & "<!-- ********** " & nResource & " Template ********** -->" & vbCrLf
txt = txt & "<" & UCase$(nResource) & ">" & vbCrLf

For X = 2 To tv.Nodes.Count
  
   'Check if Group
   If tv.Nodes(X).Children > 0 Then
    'Add Group Name
    txt = txt & vbTab & "<resDATA id=" & Chr$(34) & tv.Nodes(X).Text & Chr$(34) & ">" & vbCrLf

      'Add first child's data
      txt = txt & vbTab & vbTab & "<resNAME>" & tv.Nodes(X).Child.Text & "</resNAME>" & vbCrLf
        strTxt = tv.Nodes(X).Child.Tag
        strTxt = PutInserts(strTxt)
      txt = txt & vbTab & vbTab & "<resVALUE>" & strTxt & "</resVALUE>" & vbCrLf
          'If this is a file, copy it
          If CBool(PathFileExists(strFile)) Then
            FileCopy sPath & strTxt, nPath & strTxt
          End If
      
      Y = tv.Nodes(X).Child.Index

      'Add sibling's data
      While Y <> tv.Nodes(X).Child.LastSibling.Index
        txt = txt & vbTab & vbTab & "<resNAME>" & tv.Nodes(Y).Next.Text & "</resNAME>" & vbCrLf
        strTxt = tv.Nodes(Y).Next.Tag
        strTxt = PutInserts(strTxt)
        txt = txt & vbTab & vbTab & "<resVALUE>" & strTxt & "</resVALUE>" & vbCrLf
          'If this is a file, copy it
          If CBool(PathFileExists(strFile)) Then
            FileCopy sPath & strTxt, nPath & strTxt
          End If
         
         ' Reset y to next sibling's index.
         Y = tv.Nodes(Y).Next.Index
      Wend
   
    txt = txt & vbTab & "</resDATA>" & vbCrLf
   End If
Next

txt = txt & "</" & UCase$(nResource) & ">"

'Create a temp file
binSave nFile, txt

Saved = True
  txtString(2) = txt
  picSplitH.Cls
  picSplitH.Print "   XML Data - DblClick for Pre-Compiled (rc) Code"
End Sub

Private Sub ShowPreCompile()
Dim strGroup As String
Dim strItem As String
Dim strTag As String

picSplitH.Cls
picSplitH.Print "   Pre-Compiled Code"
'Some formatting for consistency
strGroup = StrConv(tv.Nodes(1).Text, vbProperCase)
txt = String(75, "/") & vbCrLf
txt = txt & "//  " & strGroup & vbCrLf & vbCrLf

For X = 2 To tv.Nodes.Count
  
   If tv.Nodes(X).Children > 0 Then ' There are children.
    strGroup = tv.Nodes(X).Text
    If Left$(strGroup, 1) = "#" Then strGroup = Right$(strGroup, Len(strGroup) - 1)
    
        txt = txt & String(75, "/") & vbCrLf
        'Use the word "String" to
        'format String Tables
        If InStr(UCase$(strGroup), "STRING") > 0 Then
          
          'For String Tables made with the ST Editor
          If InStr(strGroup, " - ") > 0 Then
            Y = InStr(strGroup, "- ")
            txt = txt & "// " & UCase$(Mid$(strGroup, Y + 1)) & vbCrLf
          Else
            txt = txt & "// " & strGroup & vbCrLf
          End If
          
          txt = txt & "STRINGTABLE" & vbCrLf
          txt = txt & "BEGIN" & vbCrLf
        Else
          'Use "ICONS", "BITMAPS", and other
          'Reserved words For Resource files
          'to make loading them easier
          txt = txt & "//  " & UCase$(strGroup) & vbCrLf
        End If

        'Get first child's text, and set N to its index value.
        strItem = tv.Nodes(X).Child.Text
        strTag = tv.Nodes(X).Child.Tag
        strTag = PutInserts(strTag)
        If InStr(UCase$(strGroup), "STRING") > 0 Then
          txt = txt & " " & strItem & "," & vbTab & Chr$(34) & strTag & Chr$(34) & vbCrLf
        ElseIf Right$(UCase$(strGroup), 1) = "S" Then
          txt = txt & "  " & strItem & vbTab & Left$(strGroup, Len(strGroup) - 1) & _
            vbTab & Chr$(34) & strTag & Chr$(34) & vbCrLf
        ElseIf InStr(UCase$(strGroup), "HTML") > 0 Then
          'Change HTML to 2110
          txt = txt & " " & strItem & vbTab & "2110" & vbTab & _
            Chr$(34) & strTag & Chr$(34) & vbCrLf
         Else
          txt = txt & "  " & strItem & vbTab & strGroup & _
            vbTab & Chr$(34) & strTag & Chr$(34) & vbCrLf
        End If
      
      'Get next siblings
      Y = tv.Nodes(X).Child.Index
      While Y <> tv.Nodes(X).Child.LastSibling.Index
        strItem = tv.Nodes(Y).Next.Text
        strTag = tv.Nodes(Y).Next.Tag
        strTag = PutInserts(strTag)
        If InStr(UCase$(strGroup), "STRING") > 0 Then
          txt = txt & " " & strItem & "," & vbTab & Chr$(34) & strTag & Chr$(34) & vbCrLf
        ElseIf InStr(UCase$(strGroup), "HTML") > 0 Then
          txt = txt & " " & strItem & vbTab & "2110" & vbTab & _
            Chr$(34) & strTag & Chr$(34) & vbCrLf
        ElseIf Right$(UCase$(strGroup), 1) = "S" Then
          txt = txt & "  " & strItem & vbTab & Left$(strGroup, Len(strGroup) - 1) & _
            vbTab & Chr$(34) & strTag & Chr$(34) & vbCrLf
        Else
          txt = txt & "  " & strItem & vbTab & strGroup & _
            vbTab & Chr$(34) & strTag & Chr$(34) & vbCrLf
        End If
         'Set y to next sibling's index.
         Y = tv.Nodes(Y).Next.Index
      Wend
   
    txt = txt & vbCrLf
    
    If InStr(UCase$(strGroup), "STRING") > 0 Then
      txt = Left$(txt, Len(txt) - 2)
      txt = txt & "END" & vbCrLf & vbCrLf
    End If
   End If
Next

  txtString(2) = txt
End Sub


Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
intNode = Node.Index
'This identifies the node clicked to
'other procedures, and...
'Updates the interface
If Node.Key = "root" Then
  otSource = otRoot
    cmdGroup(1).Enabled = False
    mDeleteGroup.Enabled = False
    cmdItem(0).Enabled = False
    cmdItem(1).Enabled = False
    cmdItem(2).Enabled = False
    mAddItem.Enabled = False
    mDeleteItem.Enabled = False
    mChangeItem.Enabled = False
ElseIf Node.Parent.Key = "root" Then
  otSource = otData
  strdata = Node.Text
  txtString(0) = ""
  txtString(1) = ""
    txtString(1).Enabled = True
    txtString(1).BackColor = &H80000005
    cmdGroup(1).Enabled = True
    mDeleteGroup.Enabled = True
    cmdItem(0).Enabled = True
    cmdItem(1).Enabled = False
    cmdItem(2).Enabled = False
    mAddItem.Enabled = True
    mDeleteItem.Enabled = False
    mChangeItem.Enabled = False
Else
  otSource = otName
  strName = Node.Text
  strValue = Node.Tag
  txtString(0) = strName
  txtString(1) = strValue
  
    'Disable editing file paths because they are
    'are stored in the template directory.
    'If the paths to these Resources are changed,
    'they won't be compiled - except strings...
    If InStr(UCase$(Node.Parent.Key), "STRING") = 0 Then
      txtString(1).Enabled = False
      txtString(1).BackColor = &H80000016
    Else
      txtString(1).Enabled = True
      txtString(1).BackColor = &H80000005
    End If
    cmdGroup(1).Enabled = True
    mDeleteGroup.Enabled = True
    cmdItem(0).Enabled = True
    cmdItem(1).Enabled = True
    cmdItem(2).Enabled = True
    mAddItem.Enabled = True
    mDeleteItem.Enabled = True
    mChangeItem.Enabled = True
End If
End Sub

Private Sub txtString_GotFocus(Index As Integer)
If Index = 2 Then txtString(0).SetFocus
  selTXT
End Sub


Private Sub txtString_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
  If Len(txtString(0)) >= 60 Then
    MsgBox "Key too long."
  End If
Case 1
  If Len(txtString(1)) >= 220 Then
    MsgBox "Maximum of 255 characters"
  End If
End Select
End Sub



Public Sub UserSettings()
On Error GoTo errHandler

'Save and retreive User Changes to
'the form and file, to the program will
'start as it finished
Me.Left = GetSetting(App.Title, "Editor", "Left", "Empty")
Me.Top = GetSetting(App.Title, "Editor", "Top", "Empty")
Me.Width = GetSetting(App.Title, "Editor", "Width", "Empty")
Me.Height = GetSetting(App.Title, "Editor", "Height", "Empty")
picSplitV.Left = GetSetting(App.Title, "Editor", "Splitv", "Empty")
picSplitH.Top = GetSetting(App.Title, "Editor", "SplitH", "Empty")
Me.WindowState = GetSetting(App.Title, "Editor", "State", "Empty")
strFile = GetSetting(App.Title, "Editor", "Last", "Empty")
mClip.Checked = GetSetting(App.Title, "Editor", "Clip", "Empty")

Exit Sub
'If the last file was deleted or moved...
errHandler:
'Reset default values
SaveSetting App.Title, "Editor", "Clip", False
SaveSetting App.Title, "Editor", "Left", 1000
SaveSetting App.Title, "Editor", "Top", 1000
SaveSetting App.Title, "Editor", "Width", 6465
SaveSetting App.Title, "Editor", "Height", 6285
SaveSetting App.Title, "Editor", "Splitv", 2520
SaveSetting App.Title, "Editor", "SplitH", 3135
SaveSetting App.Title, "Editor", "State", Me.WindowState
  strFile = GetPath & "templates\resource\resource.xml"
SaveSetting App.Title, "Editor", "Last", strFile

mClip.Checked = GetSetting(App.Title, "Editor", "Clip", "Empty")
Me.Left = GetSetting(App.Title, "Editor", "Left", "Empty")
Me.Top = GetSetting(App.Title, "Editor", "Top", "Empty")
Me.Width = GetSetting(App.Title, "Editor", "Width", "Empty")
Me.Height = GetSetting(App.Title, "Editor", "Height", "Empty")
picSplitV.Left = GetSetting(App.Title, "Editor", "Splitv", "Empty")
picSplitH.Top = GetSetting(App.Title, "Editor", "SplitH", "Empty")
Me.WindowState = GetSetting(App.Title, "Editor", "State", "Empty")
strFile = GetSetting(App.Title, "Editor", "Last", "Empty")
End Sub

Private Sub LoadXML(ByVal strFile)
Dim xmlDoc As DOMDocument
Dim objResource As IXMLDOMElement
Dim objData As IXMLDOMElement
Dim tvRoot As Node
Dim objAttributes As IXMLDOMNamedNodeMap
Dim objAttributeNode As IXMLDOMNode
Dim objDOMElement As IXMLDOMElement
Dim tvParent As Node
  
  Set xmlDoc = New DOMDocument
  xmlDoc.async = False
  xmlDoc.validateOnParse = True
  xmlDoc.resolveExternals = True

  xmlDoc.Load strFile
  
  If xmlDoc.parseError.reason <> "" Then
    MsgBox xmlDoc.parseError.reason
    Exit Sub
  End If

  tv.Nodes.Clear
  Set objResource = xmlDoc.documentElement
  txt = objResource.baseName
  Set tvRoot = tv.Nodes.Add(, , "root", txt, 1)
  strResource = txt

For Each objData In objResource.childNodes
  Set objAttributes = objData.attributes
  Set objAttributeNode = objAttributes.getNamedItem("id")
  Set tvParent = tv.Nodes.Add("root", tvwChild, objAttributeNode.nodeValue, objAttributeNode.nodeValue, 2)
    tvParent.Tag = objAttributeNode.nodeValue
    tvParent.EnsureVisible
  
  For Each objDOMElement In objData.childNodes
    Select Case objDOMElement.nodeName
      Case "resNAME"
        Set noDx = tv.Nodes.Add(tvParent, tvwChild, , objDOMElement.nodeTypedValue, 3)
      Case "resVALUE"
          noDx.Tag = GetInserts(objDOMElement.nodeTypedValue)
    End Select
  Next
Next

Saved = True
ShowXML
End Sub


Private Sub SetResolution()
'Get the screen resolution
X = Screen.Width / 15
Y = Screen.Height / 15

'Reset only those controls that I changed from
'the standard window's settings - everything
'else will still be the user's system settings

'Make smaller for smaller resolution
If X = 640 And Y = 480 Then
  tv.Font.Size = tv.Font.Size - 2
  picSplitH.Font.Size = picSplitH.Font.Size - 2
  txtString(0).Font.Size = txtString(0).Font.Size - 2
  txtString(1).Font.Size = txtString(1).Font.Size - 2
  txtString(2).Font.Size = txtString(2).Font.Size - 2

'Larger resolution
ElseIf X = 1024 And Y = 768 Then
  tv.Font.Size = tv.Font.Size + 2
  picSplitH.Font.Size = picSplitH.Font.Size + 2
  txtString(0).Font.Size = txtString(0).Font.Size + 2
  txtString(1).Font.Size = txtString(1).Font.Size + 2
  txtString(2).Font.Size = txtString(2).Font.Size + 2
End If
End Sub

