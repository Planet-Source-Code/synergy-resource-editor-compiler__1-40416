VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStrings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "     Language Table Editor"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7710
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox arrStrings 
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox arrNumbers 
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox lstLanguages 
      Appearance      =   0  'Flat
      Height          =   1200
      Left            =   4080
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      Index           =   1
      Left            =   4080
      TabIndex        =   12
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Clear"
      Height          =   330
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      ToolTipText     =   "Clear All Strings"
      Top             =   1365
      Width           =   1630
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Update"
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Replace Selected String"
      Top             =   1365
      Width           =   1630
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Add"
      Height          =   345
      Index           =   2
      Left            =   6360
      TabIndex        =   13
      ToolTipText     =   "Add New Language"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Delete"
      Height          =   345
      Index           =   3
      Left            =   6360
      TabIndex        =   14
      ToolTipText     =   "Delete Selected Language"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "&Translate"
      Height          =   345
      Index           =   1
      Left            =   5160
      TabIndex        =   9
      ToolTipText     =   "Translate Selected String"
      Top             =   1365
      Width           =   1060
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Char &Map"
      Height          =   345
      Index           =   0
      Left            =   4080
      TabIndex        =   8
      ToolTipText     =   "Show Character Map"
      Top             =   1365
      Width           =   1060
   End
   Begin VB.ComboBox cboLanguage 
      Height          =   315
      Index           =   0
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Delete"
      Height          =   330
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "Delete Selected String"
      Top             =   1020
      Width           =   1630
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   30
      Top             =   1800
      Width           =   7455
      Begin VB.CommandButton cmdMove 
         Appearance      =   0  'Flat
         Caption         =   "#"
         Height          =   300
         Index           =   2
         Left            =   3540
         TabIndex        =   32
         ToolTipText     =   "ReNumber"
         Top             =   1440
         Width           =   270
      End
      Begin VB.TextBox txtNumber 
         Height          =   315
         Index           =   2
         Left            =   3960
         TabIndex        =   2
         Top             =   720
         Width           =   900
      End
      Begin VB.ListBox lstStrings 
         Height          =   1620
         Index           =   2
         Left            =   4920
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2400
      End
      Begin VB.ListBox lstNumbers 
         Height          =   1620
         Index           =   2
         Left            =   3960
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtString 
         Height          =   1005
         HideSelection   =   0   'False
         Index           =   2
         Left            =   4920
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   2400
      End
      Begin VB.ListBox lstStrings 
         Height          =   1620
         Index           =   1
         Left            =   1080
         TabIndex        =   23
         Top             =   1800
         Width           =   2400
      End
      Begin VB.ListBox lstNumbers 
         Height          =   1620
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtString 
         Height          =   1005
         HideSelection   =   0   'False
         Index           =   1
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   2400
      End
      Begin VB.CommandButton cmdMove 
         Appearance      =   0  'Flat
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   14.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3540
         TabIndex        =   17
         ToolTipText     =   "Move Selected Up"
         Top             =   720
         Width           =   270
      End
      Begin VB.CommandButton cmdMove 
         Appearance      =   0  'Flat
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   14.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3540
         TabIndex        =   18
         ToolTipText     =   "Move Selected Down"
         Top             =   1020
         Width           =   270
      End
      Begin VB.TextBox txtNumber 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Translated Language"
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
         Index           =   8
         Left            =   4200
         TabIndex        =   36
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Primary Language"
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
         Index           =   7
         Left            =   360
         TabIndex        =   35
         Top             =   210
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "In&dex #"
         Height          =   195
         Index           =   1
         Left            =   4035
         TabIndex        =   21
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Translated &Language String:"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   22
         Top             =   480
         Width           =   2010
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&String:"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   20
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I&ndex #"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdList 
      Caption         =   "&Add"
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Add New String"
      Top             =   1020
      Width           =   1630
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   1560
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   345
      Index           =   1
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Cancel All Changes and Close"
      Top             =   480
      Width           =   1320
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Insert"
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Insert All Languages Into Resource Editor"
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Translated Language"
      Height          =   195
      Index           =   6
      Left            =   4320
      TabIndex        =   29
      Top             =   720
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Primary Language"
      Height          =   195
      Index           =   5
      Left            =   4440
      TabIndex        =   28
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "String Table &Title:"
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
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   27
      Top             =   240
      Width           =   1695
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
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
      Begin VB.Menu mInsert 
         Caption         =   "&Insert"
         Shortcut        =   ^I
      End
      Begin VB.Menu mPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mL2 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mAdd 
         Caption         =   "&Add String"
      End
      Begin VB.Menu mDelete 
         Caption         =   "&Delete String"
      End
      Begin VB.Menu mUpdate 
         Caption         =   "&Update Change"
      End
      Begin VB.Menu mClear 
         Caption         =   "&Clear All Indexes"
      End
      Begin VB.Menu mL3 
         Caption         =   "-"
      End
      Begin VB.Menu mMove 
         Caption         =   "Move Up"
         Index           =   0
         Shortcut        =   ^U
      End
      Begin VB.Menu mMove 
         Caption         =   "Move Down"
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu mL4 
         Caption         =   "-"
      End
      Begin VB.Menu mReNumber 
         Caption         =   "&ReNumber All"
      End
   End
   Begin VB.Menu mLanguages 
      Caption         =   "&Languages"
      Begin VB.Menu mLangAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mLangDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "&Tools"
      Begin VB.Menu mChar 
         Caption         =   "&Character Chart"
         Shortcut        =   ^C
      End
      Begin VB.Menu mTran 
         Caption         =   "&Translations"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'There is no advantage using String Editor
'for single language applications.
'Use the Resource Editor's main form
Option Explicit

'Note: using a resource extractor, you will
'see that vb stores these strings as
'Language = en (English)
'subLanguage = en_us
'or whatever your Regional settings are
'in the Control Panel.

'These are only important for Moving and
'ReNumbering listBox Items which required
'using another array of listboxes  to hold
'all languages items to avoid the
'intermediate level programming of collections
'or programming directly on the XML file
'and lose the indented formatting for easy direct
'reading of the file by beginners.
Dim bytLangCount As Byte
Dim arrLang() As String

Dim StringFile As String

Dim arrSection As Variant, arrKey As Variant
Dim X As Integer, Y As Integer, z As Integer
Dim strSection As String, strKey As String, strValue As String
Dim txt As String

Dim strPrime As String
Dim strTrans As String

Dim xmlDoc As DOMDocument
Dim objResource As IXMLDOMElement
Dim objData As IXMLDOMElement
Dim objAttributes As IXMLDOMNamedNodeMap
Dim objAttributeNode As IXMLDOMNode
Dim objDOMElement As IXMLDOMElement

Public Sub SetResolution()
  X = Screen.Width / 15
  Y = Screen.Height / 15
  If X = 640 And Y = 480 Then
    txtTitle.Font.Size = txtTitle.Font.Size - 2
    cboLanguage(0).Font.Size = cboLanguage(0).Font.Size - 2
    cboLanguage(1).Font.Size = cboLanguage(1).Font.Size - 2
    lstLanguages.Font.Size = lstLanguages.Font.Size - 2
    For Y = 1 To 2
      txtNumber(Y).Font.Size = txtNumber(Y).Font.Size - 2
      txtString(Y).Font.Size = txtString(Y).Font.Size - 2
      lstNumbers(Y).Font.Size = lstNumbers(Y).Font.Size - 2
      lstStrings(Y).Font.Size = lstStrings(Y).Font.Size - 2
    Next
  ElseIf X = 1024 And Y = 768 Then
    txtTitle.Font.Size = txtTitle.Font.Size + 2
    cboLanguage(0).Font.Size = cboLanguage(0).Font.Size + 2
    cboLanguage(1).Font.Size = cboLanguage(1).Font.Size + 2
    lstLanguages.Font.Size = lstLanguages.Font.Size + 2
    For Y = 1 To 2
      txtNumber(Y).Font.Size = txtNumber(Y).Font.Size + 2
      txtString(Y).Font.Size = txtString(Y).Font.Size + 2
      lstNumbers(Y).Font.Size = lstNumbers(Y).Font.Size + 2
      lstStrings(Y).Font.Size = lstStrings(Y).Font.Size + 2
    Next
  End If
End Sub

Private Sub cboLanguage_Click(Index As Integer)
On Error Resume Next
    
Select Case Index
Case 0 'Change Primary
  'Redo the listboxes
  lstNumbers(1).Clear
  lstStrings(1).Clear
  For X = 0 To arrNumbers(1).ListCount - 1
    lstNumbers(1).AddItem 1 & arrNumbers(cboLanguage(0).ListIndex).List(X)
    lstStrings(1).AddItem arrStrings(cboLanguage(0).ListIndex).List(X)
  Next X
  
  'Reset the Translated list
  cboLanguage(1).Clear
  For X = 0 To bytLangCount
    If Not arrLang(X) = cboLanguage(0).List(cboLanguage(0).ListIndex) Then
      cboLanguage(1).AddItem arrLang(X)
    End If
  Next X
  cboLanguage(1).ListIndex = 0
  
Case 1 'Change Translation Language
    lstNumbers(2).Clear
    lstStrings(2).Clear

  For Y = 0 To UBound(arrLang)
    If cboLanguage(1) = arrLang(Y) Then Exit For
  Next Y

  For X = 0 To arrNumbers(Y).ListCount - 1
    lstNumbers(2).AddItem (cboLanguage(1).ListIndex + 2) & arrNumbers(Y).List(X)
    lstStrings(2).AddItem arrStrings(Y).List(X)
  Next X
   
End Select

'Select the last item
lstNumbers(2).ListIndex = lstNumbers(2).ListCount - 1
lstNumbers(1).ListIndex = lstNumbers(1).ListCount - 1
txtNumber(1).SetFocus

'Check and select first unfinished item
For X = 0 To lstNumbers(1).ListCount - 1
  If lstStrings(1).List(X) = " " Then
    lstNumbers(2).ListIndex = X
    lstNumbers(1).ListIndex = X
    txtNumber(1).SetFocus
    Exit For
  ElseIf lstStrings(2).List(X) = " " Then
    lstNumbers(1).ListIndex = X
    lstNumbers(2).ListIndex = X
    txtNumber(2).SetFocus
    Exit For
  End If
Next

End Sub
Private Sub cboLanguage_GotFocus(Index As Integer)
  selTXT
End Sub


Private Sub cboLanguage_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Index = 1 Then
 cmdTools_Click (2)
End If
End Sub


Private Sub cmdButton_Click(Index As Integer)
Dim noDx As Node
Dim z As Integer, i As Integer

If Index = 0 Then 'Insert in Resource Editor

  If txtTitle = "" Then
    MsgBox "Create a 'Unique' Title for Table", vbOKOnly Or vbCritical, "Error - No Title"
    txtTitle.SetFocus
    Exit Sub
  End If
  
  'Check that Primary Language is complete
  For X = 0 To lstNumbers(1).ListCount - 1
  If lstStrings(1).List(X) = " " Then
    MsgBox "Primary Language '" & cboLanguage(0) & "' is not complete", vbOKOnly Or vbCritical, "Checking Missing or Duplicate Entries"
    'Set on first incomplete item
    lstStrings(1).ListIndex = X
    'Focus on edit box for incomplete item
    txtString(1).SetFocus
    Exit Sub
  End If
  Next
  
  'Check that Translations are complete
  For z = 0 To cboLanguage(1).ListCount - 1
  For Y = 0 To bytLangCount
  'Compare Translation to arrLang to get y
  If cboLanguage(1).List(z) = arrLang(Y) Then
    'Check each item in array y
    For X = 0 To arrNumbers(Y).ListCount - 1
      If arrStrings(Y).List(X) = " " Then
        MsgBox "Translated Language '" & cboLanguage(1).List(z) & "' is not complete", vbOKOnly Or vbCritical, "Checking Missing or Duplicate Entries"
        'Set to incomplete language
        cboLanguage(0).ListIndex = z
        'Set to first incomplete item
        lstStrings(2).ListIndex = X
        'Focus on edit box of incomplete item
        txtString(2).SetFocus
        Exit Sub
      End If
    Next
  End If
  Next
  Next
  
  'Save all changes
  Call SaveXML(StringFile)
  
  'Insert xml data into Editor treeview
  Y = 0
  For Each objData In objResource.childNodes
    Set objAttributes = objData.attributes
    Set objAttributeNode = objAttributes.getNamedItem("id")
      'Create new GroupName
      strSection = "String Table - " & txtTitle & " (" & objAttributeNode.nodeValue & ")"
      'Create treeview node for new Group
      Set noDx = frmEditor.tv.Nodes.Add("root", tvwChild, strSection, strSection, 2)
      
      'Increment KeyNumber for Group items
      Y = Y + 1
      For Each objDOMElement In objData.childNodes
        If objDOMElement.nodeName = "resNAME" Then
          'Add groupNumber to String's ItemNumber
          strKey = Y & objDOMElement.nodeTypedValue
          Set noDx = frmEditor.tv.Nodes.Add(strSection, tvwChild, , strKey, 3)
        End If
        If objDOMElement.nodeName = "resVALUE" Then
          'Add string's text value to treeview
          strKey = objDOMElement.nodeTypedValue
          noDx.Tag = strKey
        End If
      Next
  Next

End If

Unload Me
End Sub

Private Sub cmdList_Click(Index As Integer)
Dim booAdd As Boolean
booAdd = False

Select Case Index
Case 0  'Add
  mAdd_Click
  
Case 1  'Update
  mUpdate_Click
  
Case 2  'Delete
  mDelete_Click
  
Case 3  'ClearAll
  mClear_Click
  
End Select

IsSaved = False
txtNumber(1).SetFocus
End Sub


Private Sub cmdMove_Click(Index As Integer)
Select Case Index
  Case 0
    mMove_Click 0
    
  Case 1
    mMove_Click 1
    
  Case 2
    mReNumber_Click
  
End Select

IsSaved = False
End Sub


Private Sub cmdTools_Click(Index As Integer)
  Select Case Index
  Case 0
    mChar_Click
    
  Case 1
    mTran_Click
    
  Case 2  'Add new language
    mLangAdd_Click
  
  Case 3  'Delete language from list
    mLangDelete_Click
  End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Icon = frmEditor.Icon
SetResolution
IsSaved = True

'For testing without main editor
If strResource = "" Then strResource = "Resource"

Me.Top = GetSetting(App.Title, "Strings", "Top", "Empty")
Me.Left = GetSetting(App.Title, "Strings", "Left", "Empty")
StringFile = GetSetting(App.Title, "Strings", "Last", "Empty")

  If Not CBool(PathFileExists(StringFile)) Then
      StringFile = GetPath & "templates\strings\strings.xml"
  End If

LoadXML StringFile
End Sub

Private Sub LoadXML(ByVal srcFile As String)
On Error Resume Next
Dim ctl As Control
  
  Set xmlDoc = New DOMDocument
  xmlDoc.async = False
  xmlDoc.validateOnParse = True
  xmlDoc.resolveExternals = True

  xmlDoc.Load srcFile
  
  If xmlDoc.parseError.reason <> "" Then
    MsgBox xmlDoc.parseError.reason
    Exit Sub
  End If

For Each ctl In Me.Controls
  ctl.Text = ""
  ctl.Clear
Next
AddLanguages
   
  Set objResource = xmlDoc.documentElement
  txtTitle = objResource.baseName

For Each objData In objResource.childNodes

  Set objAttributes = objData.attributes
  Set objAttributeNode = objAttributes.getNamedItem("id")
    ReDim Preserve arrLang(bytLangCount)
    arrLang(bytLangCount) = objAttributeNode.nodeValue
    cboLanguage(0).AddItem arrLang(bytLangCount)
    
    Load arrNumbers(bytLangCount)
    Load arrStrings(bytLangCount)
    For Each objDOMElement In objData.childNodes
        Select Case objDOMElement.nodeName
        Case "resNAME"
          arrNumbers(bytLangCount).AddItem objDOMElement.nodeTypedValue
        Case "resVALUE"
          arrStrings(bytLangCount).AddItem objDOMElement.nodeTypedValue
        End Select
    Next
  
bytLangCount = bytLangCount + 1
Next
bytLangCount = bytLangCount - 1

Me.Visible = True
cboLanguage(0).ListIndex = 0
cboLanguage(1).ListIndex = 0
lstNumbers(1).ListIndex = lstNumbers(1).ListCount - 1
IsSaved = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  CheckSave Me, StringFile, IsSaved
If UnloadMode = vbFormCode Then
  SaveSetting App.Title, "Strings", "Last", StringFile
  SaveSetting App.Title, "Strings", "Top", Me.Top
  SaveSetting App.Title, "Strings", "Left", Me.Left
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmStrings = Nothing
End Sub


Private Sub lstLanguages_Click()
  cboLanguage(1).AddItem lstLanguages.List(lstLanguages.ListIndex)
  cboLanguage(1).ListIndex = cboLanguage(1).ListCount - 1
  
lstLanguages.Visible = False
End Sub

Private Sub lstNumbers_Click(Index As Integer)
On Error Resume Next

'In sync
For X = 0 To bytLangCount
  lstNumbers(X).Selected(lstNumbers(Index).ListIndex) = True
  lstStrings(X).Selected(lstNumbers(Index).ListIndex) = True
  arrNumbers(X).Selected(lstNumbers(Index).ListIndex) = True
  arrStrings(X).Selected(lstNumbers(Index).ListIndex) = True
Next
  
  txtNumber(1) = lstNumbers(1).List(lstNumbers(1).ListIndex)
  txtString(1) = lstStrings(1).List(lstStrings(1).ListIndex)
  txtNumber(2) = lstNumbers(2).List(lstNumbers(2).ListIndex)
  txtString(2) = lstStrings(2).List(lstStrings(2).ListIndex)
End Sub


Private Sub lstNumbers_Scroll(Index As Integer)

'In sync
If Index = 1 Then
  lstNumbers(2).TopIndex = lstNumbers(1).TopIndex
Else
  lstNumbers(1).TopIndex = lstNumbers(2).TopIndex
End If
  lstStrings(1).TopIndex = lstNumbers(1).TopIndex
  lstStrings(2).TopIndex = lstNumbers(1).TopIndex
End Sub


Private Sub lstStrings_Click(Index As Integer)
On Error Resume Next

'In sync
For X = 0 To bytLangCount
  lstStrings(X).Selected(lstStrings(Index).ListIndex) = True
  lstNumbers(X).Selected(lstStrings(Index).ListIndex) = True
  arrNumbers(X).Selected(lstNumbers(Index).ListIndex) = True
  arrStrings(X).Selected(lstNumbers(Index).ListIndex) = True
Next

  txtNumber(1) = lstNumbers(1).List(lstNumbers(1).ListIndex)
  txtString(1) = lstStrings(1).List(lstStrings(1).ListIndex)
  txtNumber(2) = lstNumbers(2).List(lstNumbers(2).ListIndex)
  txtString(2) = lstStrings(2).List(lstStrings(2).ListIndex)
End Sub


Private Sub lstStrings_Scroll(Index As Integer)

'In sync
If Index = 1 Then
  lstStrings(2).TopIndex = lstStrings(1).TopIndex
Else
  lstStrings(1).TopIndex = lstStrings(2).TopIndex
End If
  lstNumbers(1).TopIndex = lstStrings(1).TopIndex
  lstNumbers(2).TopIndex = lstStrings(1).TopIndex
End Sub



Private Sub mAdd_Click()
Dim booAdd As Boolean
If txtNumber(1) = "" Then
  MsgBox "String Tables require a numeric identifier", vbOKOnly Or vbCritical, "Error:  Identifier Missing"
  txtNumber(1).SetFocus
  Exit Sub
End If
If txtString(1) = "" Then txtString(1) = " "

If txtNumber(2) = "" Then
  MsgBox "String Tables require a numeric identifier", vbOKOnly Or vbCritical, "Error:  Identifier Missing"
  txtNumber(2).SetFocus
  Exit Sub
End If
If txtString(2) = "" Then txtString(2) = " "

For X = 0 To lstNumbers(1).ListCount - 1
  If lstNumbers(1).List(X) = txtNumber(1) Then
    MsgBox "String Tables require a 'Unique' numeric identifier", vbOKOnly Or vbCritical, "Error:  Identifier Exists"
    txtNumber(1).SetFocus
    Exit Sub
  End If
Next

If Not lstNumbers(1).ListCount = 0 Then
  For X = 0 To lstNumbers(1).ListCount - 1
      'Insert according to value
      If Val(lstNumbers(1).List(X)) > Val(txtNumber(1)) Then
        lstNumbers(1).AddItem txtNumber(1), X
        lstStrings(1).AddItem txtString(1), X
        lstNumbers(2).AddItem txtNumber(2), X
        lstStrings(2).AddItem txtString(2), X
        lstNumbers(1).ListIndex = X
        booAdd = True
        Exit For
      End If
  Next
  'Case largest
  If Not booAdd Then
    lstNumbers(1).AddItem txtNumber(1)
    lstStrings(1).AddItem txtString(1)
    lstNumbers(2).AddItem txtNumber(2)
    lstStrings(2).AddItem txtString(2)
    lstNumbers(1).ListIndex = lstNumbers(1).ListCount - 1
  End If
Else
  'First entry
  lstNumbers(1).AddItem txtNumber(1)
  lstStrings(1).AddItem txtString(1)
  lstNumbers(2).AddItem txtNumber(2)
  lstStrings(2).AddItem txtString(2)
  lstNumbers(1).ListIndex = 0
End If

  'AutoIncrement numbers
  X = Val(lstNumbers(1).List(lstNumbers(1).ListCount - 1)) + 1
  txtNumber(1) = X
  X = Val(lstNumbers(2).List(lstNumbers(2).ListCount - 1)) + 1
  txtNumber(2) = X

End Sub

Private Sub mChar_Click()
'Can't use the flexgrid, then
'comment this out and...
'See:  txtString_KeyDown
  frmChar.Show vbModal
  txtString(tBox).SetFocus
End Sub

Private Sub mClear_Click()
  txtNumber(1) = "1001"
  txtNumber(2) = "2001"
  txtString(1) = " "
  txtString(2) = " "
  Dim ctl As Control
  For Each ctl In Me
    If TypeOf ctl Is ListBox Then
      ctl.Clear
    End If
  Next ctl
  AddLanguages
  txtString(1).SetFocus
End Sub

Private Sub mDelete_Click()
  X = lstStrings(1).ListIndex
    lstNumbers(1).RemoveItem lstNumbers(1).ListIndex
    lstNumbers(2).RemoveItem lstNumbers(2).ListIndex
    lstStrings(1).RemoveItem lstStrings(1).ListIndex
    lstStrings(2).RemoveItem lstStrings(2).ListIndex
  If X <= lstStrings(1).ListCount - 1 Then
    lstStrings(1).ListIndex = X
  Else
    lstStrings(1).ListIndex = lstStrings(1).ListCount - 1
  End If
End Sub

Private Sub mExit_Click()
  Unload Me
End Sub

Private Sub mInsert_Click()
  cmdButton_Click 0
End Sub

Private Sub mLangAdd_Click()
  CheckSave Me, StringFile, IsSaved
  
  'Check if language exists already
  For X = 0 To cboLanguage(1).ListCount - 1
    If LCase$(cboLanguage(1)) = LCase$(cboLanguage(1).List(X)) Then
      'It exists
      lstLanguages.Visible = True
      Exit Sub
    End If
  Next
  
  'Add it
  If MsgBox("Add '" & cboLanguage(1).Text & "' to Language List", vbYesNo Or vbQuestion, "Confirm Add Language") = vbYes Then
    txt = StrConv(cboLanguage(1).Text, vbProperCase)
    cboLanguage(1).AddItem txt
  
    'Buffer the translation lists
    'to match the primary language
    lstNumbers(2).Clear
    lstStrings(2).Clear
    For X = 0 To lstNumbers(1).ListCount - 1
      lstNumbers(2).AddItem lstNumbers(1).List(X)
      lstStrings(2).AddItem lstStrings(1).List(X)
    Next
      
    'Coordinate the values
    For X = 0 To lstNumbers(1).ListCount - 1
      lstNumbers(2).List(X) = (Val(Left$(lstNumbers(1).List(X), 1)) * cboLanguage(1).ListCount + 1) & Mid$(lstNumbers(1).List(X), 2)
    Next
    
    'Because the msgAnswer might be No
    'Setting saved is required here
    IsSaved = False
  End If

  lstNumbers(2).ListIndex = lstNumbers(2).ListCount - 1
  lstNumbers(1).ListIndex = lstNumbers(1).ListCount - 1
  txtNumber(1).SetFocus
  For X = 0 To lstNumbers(1).ListCount - 1
    If lstStrings(1).List(X) = " " Then
      lstNumbers(2).ListIndex = X
      lstNumbers(1).ListIndex = X
      txtNumber(1).SetFocus
      Exit For
    ElseIf lstStrings(2).List(X) = " " Then
      lstNumbers(1).ListIndex = X
      lstNumbers(2).ListIndex = X
      txtNumber(2).SetFocus
      Exit For
    End If
  Next

End Sub

Private Sub mLangDelete_Click()
  X = cboLanguage(1).ListIndex
  If MsgBox("Delete '" & cboLanguage(1).Text & "' from " & txtTitle, vbYesNo Or vbCritical, "Confirm Delete") = vbYes Then
    cboLanguage(1).RemoveItem X
    cboLanguage(1).ListIndex = X - 1
    IsSaved = False
  End If
  
  lstNumbers(2).ListIndex = lstNumbers(2).ListCount - 1
  lstNumbers(1).ListIndex = lstNumbers(1).ListCount - 1
  txtNumber(1).SetFocus
  For X = 0 To lstNumbers(1).ListCount - 1
    If lstStrings(1).List(X) = " " Then
      lstNumbers(2).ListIndex = X
      lstNumbers(1).ListIndex = X
      txtNumber(1).SetFocus
      Exit For
    ElseIf lstStrings(2).List(X) = " " Then
      lstNumbers(1).ListIndex = X
      lstNumbers(2).ListIndex = X
      txtNumber(2).SetFocus
      Exit For
    End If
  Next

End Sub

Private Sub mMoveDown_Click()
cmdMove_Click 1
End Sub

Private Sub mMoveUp_Click()
cmdMove_Click 0
End Sub


Private Sub mMove_Click(Index As Integer)
On Error Resume Next
If lstNumbers(1).SelCount = 0 Then
  MsgBox "Select String to move", vbOKOnly, "Move String Item"
  Exit Sub
End If

Dim ndx As Integer

Dim txt(3) As String
ReDim arrNum(bytLangCount)
ReDim arrStr(bytLangCount)

ndx = lstNumbers(1).ListIndex

txt(0) = lstNumbers(1).List(lstNumbers(1).ListIndex)
txt(1) = lstNumbers(2).List(lstNumbers(2).ListIndex)
txt(2) = lstStrings(1).List(lstStrings(1).ListIndex)
txt(3) = lstStrings(2).List(lstStrings(2).ListIndex)

  For X = 0 To bytLangCount
    arrNum(X) = arrNumbers(X).List(arrNumbers(X).ListIndex)
    arrStr(X) = arrStrings(X).List(arrStrings(X).ListIndex)
  Next

If Index = 0 Then  'Up
  If ndx = 0 Then
    lstNumbers(1).AddItem txt(0)
    lstNumbers(2).AddItem txt(1)
    lstStrings(1).AddItem txt(2)
    lstStrings(2).AddItem txt(3)
    
    For X = 0 To bytLangCount
      arrNumbers(X).AddItem arrNum(X)
      arrNumbers(X).RemoveItem 0
      arrStrings(X).AddItem arrStr(X)
      arrStrings(X).RemoveItem 0
    Next
    
    lstNumbers(1).RemoveItem 0
    lstNumbers(2).RemoveItem 0
    lstStrings(1).RemoveItem 0
    lstStrings(2).RemoveItem 0
    
    lstNumbers(1).ListIndex = lstNumbers(1).ListCount - 1
  Else
    lstNumbers(1).RemoveItem ndx
    lstNumbers(2).RemoveItem ndx
    lstStrings(1).RemoveItem ndx
    lstStrings(2).RemoveItem ndx
    
    For X = 0 To bytLangCount
      arrNumbers(X).RemoveItem ndx
      arrNumbers(X).AddItem arrNum(X), ndx - 1
      arrStrings(X).RemoveItem ndx
      arrStrings(X).AddItem arrStr(X), ndx - 1
    Next
    
    lstNumbers(1).AddItem txt(0), ndx - 1
    lstNumbers(2).AddItem txt(1), ndx - 1
    lstStrings(1).AddItem txt(2), ndx - 1
    lstStrings(2).AddItem txt(3), ndx - 1
    
    lstNumbers(1).ListIndex = ndx - 1
  End If
  
ElseIf Index = 1 Then 'Down
  
  If ndx = lstNumbers(1).ListCount - 1 Then
    lstNumbers(1).RemoveItem ndx
    lstNumbers(2).RemoveItem ndx
    lstStrings(1).RemoveItem ndx
    lstStrings(2).RemoveItem ndx
    
    For X = 0 To bytLangCount
      arrNumbers(X).RemoveItem ndx
      arrNumbers(X).AddItem arrNum(X), 0
      arrStrings(X).RemoveItem ndx
      arrStrings(X).AddItem arrStr(X), 0
    Next
    
    lstNumbers(1).AddItem txt(0), 0
    lstNumbers(2).AddItem txt(1), 0
    lstStrings(1).AddItem txt(2), 0
    lstStrings(2).AddItem txt(3), 0
    
    lstNumbers(1).ListIndex = 0
  
  Else
    lstNumbers(1).RemoveItem ndx
    lstNumbers(2).RemoveItem ndx
    lstStrings(1).RemoveItem ndx
    lstStrings(2).RemoveItem ndx
    
    For X = 0 To bytLangCount
      arrNumbers(X).RemoveItem ndx
      arrNumbers(X).AddItem arrNum(X), ndx + 1
      arrStrings(X).RemoveItem ndx
      arrStrings(X).AddItem arrStr(X), ndx + 1
    Next
    
    lstNumbers(1).AddItem txt(0), ndx + 1
    lstNumbers(2).AddItem txt(1), ndx + 1
    lstStrings(1).AddItem txt(2), ndx + 1
    lstStrings(2).AddItem txt(3), ndx + 1
    
    lstNumbers(1).ListIndex = ndx + 1
  End If
  
End If
End Sub

Private Sub mOpen_Click()
On Error Resume Next
With CommonDialog1
  .DialogTitle = "Open String xmlResource"
  .filename = "StringResource.xml"
  .Flags = cdlOFNFileMustExist
  .Filter = "XML Files (*.xml)|*.xml"
  .InitDir = App.Path
  .ShowOpen
  txt = .filename
End With
  If Not Len(txt) Then Exit Sub
  
  CheckSave Me, StringFile, IsSaved
  
  StringFile = txt
  LoadXML StringFile
End Sub

Private Sub mPrint_Click()
  txt = binOpen(StringFile)

PrintCode txt
End Sub

Private Sub mReNumber_Click()
  For X = 0 To lstNumbers(1).ListCount - 1
    lstNumbers(1).List(X) = (X) + 1000
    lstNumbers(2).List(X) = (X) + ((cboLanguage(1).ListIndex + 2) * 1000)
  Next

  For X = 0 To lstNumbers(1).ListCount - 1
    For Y = 0 To bytLangCount
      arrNumbers(bytLangCount).List(X) = Right$(lstNumbers(1).List(X), 3)
    Next
  Next
End Sub

Private Sub mSave_Click()
  txt = GetPath & "Templates\" & txtTitle & "\" & txtTitle & ".xml"
'If Not UCase$(txtTitle) = "STRINGS" Then
    Call SaveXML(txt)
'ElseIf Not IsSaved Then
'  MsgBox "Cannot 'Save' to default 'Strings' File   ", vbOKOnly Or vbInformation, "Save Canceled"
'  mSaveAs_Click
'End If
End Sub

Private Sub mSaveAs_Click()
On Error Resume Next

  CheckSave Me, StringFile, IsSaved

txt = InputBox("Enter a New Filename", "Save As - Resource String Editor", txtTitle)
If Not txt = "" Then
  If Right$(txt, 4) = ".xml" Then
    txt = GetPath & "Templates\" & txt & "\" & txt
    txtTitle = Left$(txt, Len(txt) - 4)
  Else
    'Only .xml as extension
    X = InStr(txt, ".")
    If X > 0 Then
      txt = Left$(txt, X - 1)
    End If
    txtTitle = txt
    
    txt = GetPath & "Templates\" & txt & "\" & txt & ".xml"
    If LCase$(txt) = LCase$(GetPath & "Templates\strings\strings.xml") Then
      MsgBox "Cannot 'Save' to default 'Strings' File   ", vbOKOnly Or vbInformation, "Save Canceled"
    Else
      Call SaveXML(txt)
    End If
  End If
End If
End Sub


Private Sub mTran_Click()
  Translate
  mUpdate_Click
End Sub

Private Sub mUpdate_Click()
  If lstStrings(1).ListIndex > -1 Then
    lstNumbers(1).List(lstNumbers(1).ListIndex) = txtNumber(1)
    lstStrings(1).List(lstStrings(1).ListIndex) = txtString(1)
    lstNumbers(2).List(lstNumbers(2).ListIndex) = txtNumber(2)
    lstStrings(2).List(lstStrings(2).ListIndex) = txtString(2)
  End If
End Sub

Private Sub txtNumber_GotFocus(Index As Integer)
On Error Resume Next
  With txtNumber(Index)
    .SelStart = Len(.Text) - 1
    .SelLength = 1
  End With
End Sub


Private Sub txtNumber_KeyPress(Index As Integer, KeyAscii As Integer)
Dim txtClip As String
txt = "The key does not form a valid number."

'Validate for numbers and editing
Select Case KeyAscii
    
    'Editing
  Case vbKeyDelete, vbKeyBack, vbKeyEnd, vbKeyHome, vbKeyRight, vbKeyLeft
  
    'Cut & Copy
  Case 24, 3
    'vbKeyControl and vbKeyX, or vbKeyC
  
    'Paste
  Case 22
    'vbKeyControl and  vbKeyV
    txtClip = Clipboard.GetText
    If Not IsNumeric(txtClip) Then
      KeyAscii = 0
      MsgBox txt
    End If
  
    'Numeric only
  Case 48 To 57
  
    'Scratch the rest
  Case Else
    KeyAscii = 0
    MsgBox txt
End Select

End Sub


Private Sub txtNumber_LostFocus(Index As Integer)
'For the string tables to be of any value
'in a program, they must be coordinated
'for easy access and assignment.

'Restrict the numbers
GoTo hell
  If Not IsNumeric(txtNumber(Index)) Then
    MsgBox "Primary Language Identifier Range is 1000 to 1999."
    txtNumber(Index).SetFocus
    Exit Sub
  End If

If Index = 1 Then
  If Val(txtNumber(1)) < 1000 Or Val(txtNumber(1)) > 1999 Then
    MsgBox "Primary Language Identifier Range is 1000 to 1999."
    txtNumber(1).SetFocus
    Exit Sub
  End If
Else
  ' x accounts for comboBoxes starting at 0
  ' and the primary language being #1
  'So, the listbox is translation = +2
  X = cboLanguage(1).ListIndex + 2
  If Val(txtNumber(2)) < X * 1000 Or Val(txtNumber(2)) > Val(X & "999") Then
    MsgBox "Translated Language Identifier Range is " & X * 1000 & " to " & X * 1000 + 999 & "."
    txtNumber(2).SetFocus
    Exit Sub
  End If
End If

'Force number lists coordination
  If cboLanguage(1).ListIndex = -1 Then
    txtNumber(2) = cboLanguage(1).ListIndex + 3 & Right$(txtNumber(1), 3)
  Else
    txtNumber(2) = cboLanguage(1).ListIndex + 2 & Right$(txtNumber(1), 3)
  End If
  
hell:
End Sub

Private Sub txtString_GotFocus(Index As Integer)
  selTXT
  tBox = Index
End Sub


Private Sub txtString_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Use ascii values to insert characters
'If ascii code is < 3 digits then it must
'begin with a zero (0) - lookup
'Character sets' in helpfile

Static strAsc As String
Dim AltDown, txt
  AltDown = (Shift And vbAltMask) > 0

If AltDown Then
  If IsNumeric(Chr$(KeyCode)) Then
    strAsc = Right$(strAsc, 2) & Val(Chr$(KeyCode))
  End If
End If

If Len(strAsc) = 3 And Val(strAsc) > 31 And Val(strAsc) < 256 Then
  txtString(Index).SelText = Chr$(strAsc)
  strAsc = ""
End If
End Sub


Private Sub txtTitle_GotFocus()
  selTXT
End Sub



Public Sub SaveXML(ByVal tmpFile As String)
Dim strTxt As String

strTxt = "<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " encoding=" & Chr$(34) & "ISO-8859-1" & Chr$(34) & "?>" & vbCrLf
strTxt = strTxt & "<!-- ********** " & txtTitle & " StringsTemplate ********** -->" & vbCrLf
strTxt = strTxt & "<" & txtTitle & ">" & vbCrLf

  strTxt = strTxt & vbTab & "<resDATA id=" & Chr$(34) & cboLanguage(0) & Chr$(34) & ">" & vbCrLf
    For X = 0 To lstNumbers(1).ListCount - 1
      strTxt = strTxt & vbTab & vbTab & "<resNAME>" & Right$(lstNumbers(1).List(X), 3) & "</resNAME>" & vbCrLf
      strTxt = strTxt & vbTab & vbTab & "<resVALUE>" & lstStrings(1).List(X) & "</resVALUE>" & vbCrLf
    Next
  strTxt = strTxt & vbTab & "</resDATA>" & vbCrLf
  
  strTxt = strTxt & vbTab & "<resDATA id=" & Chr$(34) & cboLanguage(1) & Chr$(34) & ">" & vbCrLf
    For X = 0 To lstNumbers(2).ListCount - 1
      strTxt = strTxt & vbTab & vbTab & "<resNAME>" & Right$(lstNumbers(2).List(X), 3) & "</resNAME>" & vbCrLf
      strTxt = strTxt & vbTab & vbTab & "<resVALUE>" & lstStrings(2).List(X) & "</resVALUE>" & vbCrLf
    Next
  strTxt = strTxt & vbTab & "</resDATA>" & vbCrLf

For Y = 0 To bytLangCount
  If arrLang(Y) = cboLanguage(0) Or arrLang(Y) = cboLanguage(1) Then
  Else
  strTxt = strTxt & vbTab & "<resDATA id=" & Chr$(34) & arrLang(Y) & Chr$(34) & ">" & vbCrLf
    For X = 0 To arrNumbers(Y).ListCount - 1
      strTxt = strTxt & vbTab & vbTab & "<resNAME>" & Right$(arrNumbers(Y).List(X), 3) & "</resNAME>" & vbCrLf
      strTxt = strTxt & vbTab & vbTab & "<resVALUE>" & arrStrings(Y).List(X) & "</resVALUE>" & vbCrLf
    Next
  strTxt = strTxt & vbTab & "</resDATA>" & vbCrLf
  End If
Next Y
  
       strTxt = strTxt & "</" & txtTitle & ">"

If LCase$(tmpFile) = LCase$(StringFile) Then
  tmpFile = Left$(StringFile, Len(StringFile) - 3) & "tmp"
End If

On Error GoTo errHandler
Call binSave(tmpFile, strTxt)

If Right$(tmpFile, 3) = "tmp" Then
  DeleteFile StringFile
  Name tmpFile As StringFile
Else
  StringFile = tmpFile
End If

IsSaved = True
Exit Sub
errHandler:
  If Err = 58 Then 'File still exists
    DeleteFile StringFile
    Resume
  ElseIf Err = 76 Then 'Path not found
    'New Template so Create Folder
    MkDir strFilePath(tmpFile)
    Resume
  End If
End Sub


Public Sub Translate()
'Get worldlingo.com's language tags
  strPrime = GetPrimeLang
  strTrans = GetTransLang
If strPrime = "" Then
  MsgBox "Language '" & cboLanguage(0) & "' not supported for translation"
  Exit Sub
ElseIf strTrans = "" Then
  MsgBox "Language '" & cboLanguage(1) & "' not supported for translation"
  Exit Sub
End If

'Using the webBrowser control
'will get the text as shown on the webpage
'which might be different than that in the webpage's
'source code as retrieved by
'an inet control
frmBrowse.WebBrowser1.Navigate "http://www.worldlingo.com/wl/Translate?wl_text=" _
& txtString(1) & "&wl_gloss=1&wl_srclang=" _
& strPrime & "&wl_trglang=" & strTrans

'Mostly, this just keeps the HourGlass
'from flickering.  Otherwise, put the
'DOM parsing routine in the webBrowser's
'DocumentComplete event
Do
  DoEvents
  Screen.MousePointer = vbHourglass
Loop While frmBrowse.WebBrowser1.Busy

'Reference 'MS HTML Object Library'
Dim doc As HTMLDocument
Set doc = frmBrowse.WebBrowser1.Document

On Error Resume Next
Dim i As Long, j As Long, Elements

For X = 0 To doc.Forms.length - 1
For Y = 0 To doc.Forms(X).Elements.length - 1
 If doc.Forms(X).Elements(Y).Type = "textarea" Then
  If doc.Forms(X).Elements(Y).Name = "wl_result" Then
    txtString(2) = doc.Forms(X).Elements(Y).Value
    IsSaved = False
    Exit For
  End If
 End If
Next
Next

Screen.MousePointer = vbDefault
End Sub


Public Function GetPrimeLang()
Select Case cboLanguage(0).Text
 Case "English": GetPrimeLang = "EN"
 Case "French": GetPrimeLang = "FR"
 Case "German": GetPrimeLang = "DE"
 Case "Italian": GetPrimeLang = "IT"
 Case "Portuguese": GetPrimeLang = "PT"
 Case "Spanish": GetPrimeLang = "ES"
End Select
End Function

Public Function GetTransLang()
Select Case cboLanguage(1).Text
 Case "English": GetTransLang = "EN"
 Case "French": GetTransLang = "FR"
 Case "German": GetTransLang = "DE"
 Case "Italian": GetTransLang = "IT"
 Case "Portuguese": GetTransLang = "PT"
 Case "Spanish": GetTransLang = "ES"
End Select
End Function

Public Sub AddLanguages()
'These are the basic languages supported by
'the "Western ISO" character set.

'Some other languages can be included by
'routing to different machine-translators.
 lstLanguages.AddItem "English"
 lstLanguages.AddItem "French"
 lstLanguages.AddItem "German"
 lstLanguages.AddItem "Italian"
 lstLanguages.AddItem "Portuguese"
 lstLanguages.AddItem "Spanish"
End Sub

