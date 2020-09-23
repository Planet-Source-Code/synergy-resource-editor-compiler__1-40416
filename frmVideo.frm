VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmVideo 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   313
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl2.Animation Animation1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      _Version        =   327681
      AutoPlay        =   -1  'True
      FullWidth       =   73
      FullHeight      =   41
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Searching - Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using mci is not the best way to show
'default AVIs found in the shell dll
'because mci, as far as i know, doesn't
'mask the background color.  The Animation
'Control that comes with vb6 does.

Dim tmpFile As String

Private Sub Form_Load()
'This is just a demostration.
'Actually, this avi should be
'extracted from shell32 at runtime,
'then played - not saved in res file.

Dim F As Long
Dim strPath As String
Dim strShortFile As String * 128
Me.Icon = frmEditor.Icon

'Load AVI
vidArray = LoadResData("Search", "VIDEO")
'Save AVI to hd
F = FreeFile
tmpFile = strFilePath(strFile) & "video.avi"
Open tmpFile For Binary As F
Put F, , vidArray
Close F

'In properties - set autoplay = true: backstyle = transparent
Animation1.open tmpFile
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  X = mciSendString("close video", 0&, 0, 0&)
  tmpFile = strFilePath(strFile) & "video.avi"
  DeleteFile tmpFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmVideo = Nothing
End Sub


