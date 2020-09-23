VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmChar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Map"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3300
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5821
      _Version        =   393216
      Rows            =   13
      Cols            =   20
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      HighLight       =   2
      FillStyle       =   1
      GridLinesFixed  =   0
      ScrollBars      =   0
      MergeCells      =   4
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1530
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Insert"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label labZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Index           =   0
      Left            =   5970
      TabIndex        =   3
      Top             =   30
      Width           =   735
   End
   Begin VB.Label labZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdButton_Click(Index As Integer)
Select Case Index
Case 0
  frmStrings.txtString(tBox).SelText = Grid1.Text
  IsSaved = False
  Unload Me
Case 1
  Unload Me
End Select
End Sub

Private Sub Form_Load()
Me.Icon = frmEditor.Icon
On Error GoTo errHandler
Dim X, Y, z

   X = Grid1.Width / 20
   Y = Grid1.Height / 13
   
For z = 0 To 19
   Grid1.ColWidth(z) = X
Next
Grid1.Width = Grid1.ColWidth(0) * 20

For z = 0 To 12
   Grid1.RowHeight(z) = Y
Next
Grid1.Height = Grid1.RowHeight(0) * 12

Y = 0
For X = 0 To 12
  Grid1.Row = X
  
  Y = Y + 20
  For z = 1 To 20
    Grid1.Col = z - 1
    Grid1.Text = Chr$(z + Y)
  Next
   
Next

errHandler:
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set frmChar = Nothing
End Sub


Private Sub Grid1_Click()
  labZoom(0).Caption = Grid1.Text
  labZoom(1).Caption = Grid1.Text
End Sub


