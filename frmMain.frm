VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Testplan and bom version compare (soft version 6.8 great_guo)"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10275
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "And testplan compare"
      Height          =   1815
      Left            =   4800
      TabIndex        =   27
      Top             =   3840
      Width           =   2775
      Begin VB.CheckBox Check2 
         Caption         =   "Not initializel testplan"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Panel Boards"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CommandButton cmdBoards 
         Caption         =   "&CreateTestplan"
         Height          =   375
         Left            =   1440
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Go>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   """analog\"" is null in testplan file"
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5175
      Left            =   7680
      TabIndex        =   24
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Lib Tool (bom to 3070 board)"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   4680
         Width           =   2295
      End
      Begin VB.CommandButton cmdToVerBoard 
         Caption         =   "Read board value tool"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   4200
         Width           =   2295
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Multiple boms merge one bom(one device appear once)"
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton cmdBom8GeVer 
         Caption         =   "Multiple boms merge one bom(one device appear once)  &GO>>"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   36
         Top             =   3360
         Width           =   2295
      End
      Begin VB.OptionButton Option3 
         Caption         =   "3070 Board Format Files To Version Board "
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdBom8Ge 
         Caption         =   "Board Format To Version Board &GO>>"
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CommandButton cmdBomAndBom 
         Caption         =   "Bom 1 and Bom 2       &GO>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bom and bom compare (only two bom file)"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bom and testplan compare"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtVer_8 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   47
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox txtVer_7 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   46
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtVer_6 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   45
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtVer_5 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   44
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtVer_4 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   43
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtVer_3 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   42
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtVer_2 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   41
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtVer_1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6360
         TabIndex        =   40
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtBom5 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Please open bom5 file!(DblClick me open file!"")"
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtBom6 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Please open bom6 file!(DblClick me open file!"")"
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox txtBom7 
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Please open bom7 file!(DblClick me open file!"")"
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox txtBom8 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Please open bom8 file!(DblClick me open file!"")"
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox txtBom1 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Please open bom1 file!(DblClick me open file!"")"
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox txtBom2 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Please open bom2 file!(DblClick me open file!"")"
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtBom3 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Please open bom3 file!(DblClick me open file!"")"
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtBom4 
         BackColor       =   &H0080FFFF&
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Please open bom4 file!(DblClick me open file!"")"
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txtTestplan 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Please open testplan file!(DblClick me open file!"")"
         Top             =   240
         Width           =   7215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2040
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label l5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label l6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   22
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label l7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   21
         Top             =   2880
         Width           =   1800
      End
      Begin VB.Label l8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   3240
         Width           =   1800
      End
      Begin VB.Label l1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label l2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label l3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label l4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   1800
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   4575
      Begin VB.Label Msg4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Msg3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Msg2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Msg1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Menu File_ 
      Caption         =   "File"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim PrmPath As String
Dim strBom1Path As String
Dim bRunBom1 As Boolean
 Dim strTestplanPath         As String
 Dim bRunTestplan As Boolean
 Dim strBom2Path As String
Dim bRunBom2 As Boolean
 Dim strBom3Path As String
Dim bRunBom3 As Boolean
Dim strBom4Path As String
Dim bRunBom4 As Boolean

Dim strBom5Path As String
Dim bRunBom5 As Boolean
Dim strBom6Path As String
Dim bRunBom6 As Boolean
Dim strBom7Path As String
Dim bRunBom7 As Boolean
Dim strBom8Path As String
Dim bRunBom8 As Boolean
Dim Boards As Boolean
Dim strBoardsNumber As String
Dim Not_initializel_testplan As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
   cmdBoards.Enabled = True
   Boards = True
   Else
   cmdBoards.Enabled = False
   Boards = False
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
  Not_initializel_testplan = True
  Else
  Not_initializel_testplan = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
     
     strAnalog_ = ""
   Else
    
    strAnalog_ = "analog/"
    
End If
End Sub

Private Sub cmdBoards_Click()
If Check3.Value = 1 Then
     
     strAnalog_ = ""
   Else
    
    strAnalog_ = "analog/"
    
End If
frmCreateTestplan.Show
End Sub

Private Sub cmdBom8Ge_Click()
Dim bAllVer As Boolean
Dim a

'MkDir PrmPath & "BomCompare\Testplan_Tmp_Analog"
'MkDir PrmPath & "BomCompare\Bom_1"
'MkDir PrmPath & "BomCompare\Bom_2"
'MkDir PrmPath & "BomCompare\Bom_3"
'MkDir PrmPath & "BomCompare\Bom_4"
'MkDir PrmPath & "BomCompare\Bom_5"
'MkDir PrmPath & "BomCompare\Bom_6"
'MkDir PrmPath & "BomCompare\Bom_7"
'MkDir PrmPath & "BomCompare\Bom_8"

 On Error Resume Next
   bAllVer = bRunBom1 Or bRunBom2 Or bRunBom3 Or bRunBom4 Or bRunBom5 Or bRunBom6 Or bRunBom7 Or bRunBom8







 If bAllVer = True Then
           Option1.Enabled = False
     Option2.Enabled = False
    ' Option3.Enabled = False
     Option4.Enabled = False
     txtBom1.Enabled = False
     txtBom2.Enabled = False
    txtBom3.Enabled = False
     txtBom4.Enabled = False
     txtBom5.Enabled = False
     txtBom6.Enabled = False
     txtBom7.Enabled = False
     txtBom8.Enabled = False
     txtTestplan.Enabled = False
    
   
   
   
 '  Call Open_Testplan_Read
   
   
   
   Else
     If bAllVer = False Then
       MsgBox "Please check bom file!", vbCritical
         txtBom1.SetFocus
         Exit Sub
     End If
     
     
 End If
 
If bRunBom1 = True Then
   strVerName_1 = "VERSION " & """" & Trim(txtVer_1.Text) & """"
    If Trim(txtVer_1.Text) = "" Then
       txtVer_1.SetFocus
       MsgBox "Please input Bom1 version name!", vbInformation
       Exit Sub
    End If
End If

If bRunBom2 = True Then
   strVerName_2 = "VERSION " & """" & Trim(txtVer_2.Text) & """"
    If Trim(txtVer_2.Text) = "" Then
       txtVer_2.SetFocus
       MsgBox "Please input Bom2 version name!", vbInformation
       Exit Sub
    End If
End If
 
 If bRunBom3 = True Then
   strVerName_3 = "VERSION " & """" & Trim(txtVer_3.Text) & """"
    If Trim(txtVer_3.Text) = "" Then
       txtVer_3.SetFocus
       MsgBox "Please input Bom3 version name!", vbInformation
       Exit Sub
    End If
End If
 
 If bRunBom4 = True Then
   strVerName_4 = "VERSION " & """" & Trim(txtVer_4.Text) & """"
    If Trim(txtVer_4.Text) = "" Then
       txtVer_4.SetFocus
       MsgBox "Please input Bom4 version name!", vbInformation
       Exit Sub
    End If
End If
 
 If bRunBom5 = True Then
   strVerName_5 = "VERSION " & """" & Trim(txtVer_5.Text) & """"
    If Trim(txtVer_5.Text) = "" Then
       txtVer_5.SetFocus
       MsgBox "Please input Bom5 version name!", vbInformation
       Exit Sub
    End If
End If
 
 If bRunBom6 = True Then
   strVerName_6 = "VERSION " & """" & Trim(txtVer_6.Text) & """"
    If Trim(txtVer_6.Text) = "" Then
       txtVer_6.SetFocus
       MsgBox "Please input Bom6 version name!", vbInformation
       Exit Sub
    End If
End If

 
 If bRunBom7 = True Then
   strVerName_7 = "VERSION " & """" & Trim(txtVer_7.Text) & """"
    If Trim(txtVer_7.Text) = "" Then
       txtVer_7.SetFocus
       MsgBox "Please input Bom7 version name!", vbInformation
       Exit Sub
    End If
End If

 If bRunBom8 = True Then
   strVerName_8 = "VERSION " & """" & Trim(txtVer_8.Text) & """"
    If Trim(txtVer_8.Text) = "" Then
       txtVer_8.SetFocus
       MsgBox "Please input Bom8 version name!", vbInformation
       Exit Sub
    End If
End If




 
' If bRunTestplan = False Then
'     Call Kill_File
'     Call Kill_Device
'     comOK.Enabled = True
'     txtBom1.Enabled = True
'     txtBom2.Enabled = True
'     txtBom3.Enabled = True
'     txtBom4.Enabled = True
'     txtBom5.Enabled = True
'     txtBom6.Enabled = True
'     txtBom7.Enabled = True
'     txtBom8.Enabled = True
'     cmdBoards.Enabled = True
'     Check1.Enabled = True
'     txtTestplan.Enabled = True
'
' End If
 
   strMsg = MsgBox("Do you want to continue ?", 52, "Warning!")
If strMsg = vbYes Then
      GoTo Start
   ElseIf strMsg = vbNo Then
    Exit Sub
End If

           Option1.Enabled = False
     Option2.Enabled = False
     Option3.Enabled = False
     txtBom1.Enabled = False
     txtBom2.Enabled = False
     txtBom3.Enabled = False
     txtBom4.Enabled = False
     txtBom5.Enabled = False
     txtBom6.Enabled = False
     txtBom7.Enabled = False
     txtBom8.Enabled = False
     txtTestplan.Enabled = False
    cmdToVerBoard.Enabled = False
cmdBom8GeVer.Enabled = False

cmdToVerBoard.Enabled = False

Start:


'Open PrmPath & "BomCompare\8_Bom_basic_device.dll" For Output As #99
 
'frmBomVerOption.Show
  cmdBom8Ge.Enabled = False
  Command1.Enabled = False
  cmdToVerBoard.Enabled = False
  Frame4.Enabled = False
  cmdBomAndBom.Enabled = False
 Frame2.Enabled = False
 Option1.Enabled = False
 Option2.Enabled = False
  MkDir PrmPath & "BomCompare\BomAndBom_Comp"
 MkDir PrmPath & "BomCompare\Basic_Tmp"
  MkDir PrmPath & "BomCompare\Basic_All_Bom"
  Kill PrmPath & "BomCompare\Basic_Tmp\*.*"
 
  Kill PrmPath & "BomCompare\Basic_All_Bom\*.*"
 Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
If bRunBom1 = True Then
  MkDir PrmPath & "BomCompare\Bom_1"
  Call Bom8Comp_Bom1
End If
cmdBom8GeVer.Enabled = False
If bRunBom2 = True Then
 MkDir PrmPath & "BomCompare\Bom_2"
 Call Bom8Comp_Bom2
End If
If bRunBom3 = True Then
  MkDir PrmPath & "BomCompare\Bom_3"
    Call Bom8Comp_Bom3
End If
If bRunBom4 = True Then
  MkDir PrmPath & "BomCompare\Bom_4"
   Call Bom8Comp_Bom4
End If

If bRunBom5 = True Then
  MkDir PrmPath & "BomCompare\Bom_5"
    Call Bom8Comp_Bom5
End If
If bRunBom6 = True Then
  MkDir PrmPath & "BomCompare\Bom_6"
  Call Bom8Comp_Bom6
End If
If bRunBom7 = True Then
   MkDir PrmPath & "BomCompare\Bom_7"
  Call Bom8Comp_Bom7
End If
If bRunBom8 = True Then
   MkDir PrmPath & "BomCompare\Bom_8"
  Call Bom8Comp_Bom8
End If
Close #3

Msg1.Caption = "Start compare..."
Msg2.Caption = "Please wait..."
Msg3.Caption = ""
Msg4.Caption = ""

Call FileCompStart_8_NoTestplan





 
 
Kill PrmPath & "BomCompare\BomAndBom_Comp\*.*"
Kill PrmPath & "BomCompare\Bom_1\*.*"
Kill PrmPath & "BomCompare\Bom_2\*.*"
Kill PrmPath & "BomCompare\Bom_3\*.*"
Kill PrmPath & "BomCompare\Bom_4\*.*"
Kill PrmPath & "BomCompare\Bom_5\*.*"
Kill PrmPath & "BomCompare\Bom_6\*.*"
Kill PrmPath & "BomCompare\Bom_7\*.*"
Kill PrmPath & "BomCompare\Bom_8\*.*"
RmDir PrmPath & "BomCompare\Bom_1\"
RmDir PrmPath & "BomCompare\Bom_2\"
RmDir PrmPath & "BomCompare\Bom_3\"
RmDir PrmPath & "BomCompare\Bom_4\"
RmDir PrmPath & "BomCompare\Bom_5\"
RmDir PrmPath & "BomCompare\Bom_6\"
RmDir PrmPath & "BomCompare\Bom_7\"
RmDir PrmPath & "BomCompare\Bom_8\"
RmDir PrmPath & "BomCompare\BomAndBom_Comp"



 


'Msg3.Caption = "Runing create BomDiff.txt file..."
'Msg4.Caption = "Please wait..."

 Call FileHeBin
 
Msg4.Caption = "Compare file end!"
       cmdBom8Ge.Enabled = True
       cmdToVerBoard.Enabled = True
      Option1.Enabled = True
      cmdToVerBoard.Enabled = True
      Command1.Enabled = True
     Option2.Enabled = True
     Option3.Enabled = True
cmdBom8GeVer.Enabled = True
     txtBom1.Enabled = True
     txtBom2.Enabled = True
     txtBom3.Enabled = True
     txtBom4.Enabled = True
     txtBom5.Enabled = True
     txtBom6.Enabled = True
     txtBom7.Enabled = True
     txtBom8.Enabled = True
 cmdToVerBoard.Enabled = True
 Option1.Enabled = True
 Option2.Enabled = True
 Frame2.Enabled = True
 'Close #99
 MsgBox "3070 Board Format Compare ok!,Output board.txt file!,please check!", vbInformation
End Sub

Private Sub cmdBom8GeVer_Click()
Dim bAllVer As Boolean
Dim a
 On Error Resume Next
'MkDir PrmPath & "BomCompare\Testplan_Tmp_Analog"
'MkDir PrmPath & "BomCompare\Bom_1"
'MkDir PrmPath & "BomCompare\Bom_2"
'MkDir PrmPath & "BomCompare\Bom_3"
'MkDir PrmPath & "BomCompare\Bom_4"
'MkDir PrmPath & "BomCompare\Bom_5"
'MkDir PrmPath & "BomCompare\Bom_6"
'MkDir PrmPath & "BomCompare\Bom_7"
'MkDir PrmPath & "BomCompare\Bom_8"
    strMsg = MsgBox("Do you want to continue ?", 52, "Warning!")
If strMsg = vbYes Then
      GoTo Start
   ElseIf strMsg = vbNo Then
   
   ' GoTo GOend
   
    Exit Sub
End If



Start:

   bAllVer = bRunBom1 Or bRunBom2 Or bRunBom3 Or bRunBom4 Or bRunBom5 Or bRunBom6 Or bRunBom7 Or bRunBom8







 If bAllVer = True Then
           Option1.Enabled = False
     Option2.Enabled = False
     Option3.Enabled = False
     Option4.Enabled = False
     txtBom1.Enabled = False
     txtBom2.Enabled = False
     txtBom3.Enabled = False
     txtBom4.Enabled = False
     txtBom5.Enabled = False
     txtBom6.Enabled = False
     txtBom7.Enabled = False
     txtBom8.Enabled = False
     txtTestplan.Enabled = False
    
   
   
   
 '  Call Open_Testplan_Read
   
   
   
   Else
     If bAllVer = False Then
       MsgBox "Please check bom file!", vbCritical
         txtBom1.SetFocus
         Exit Sub
     End If
     
     
 End If
 

 
 
' If bRunTestplan = False Then
'     Call Kill_File
'     Call Kill_Device
'     comOK.Enabled = True
'     txtBom1.Enabled = True
'     txtBom2.Enabled = True
'     txtBom3.Enabled = True
'     txtBom4.Enabled = True
'     txtBom5.Enabled = True
'     txtBom6.Enabled = True
'     txtBom7.Enabled = True
'     txtBom8.Enabled = True
'     cmdBoards.Enabled = True
'     Check1.Enabled = True
'     txtTestplan.Enabled = True
'
' End If



Kill PrmPath & "BomCompare\Bom_1\*.*"
Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Output As #58
Close #58

  cmdBom8GeVer.Enabled = False
  Frame4.Enabled = False
  cmdBomAndBom.Enabled = False
 Option1.Enabled = False
 Option2.Enabled = False
 Option3.Enabled = False
 Command1.Enabled = False
 cmdToVerBoard.Enabled = False
 ' MkDir PrmPath & "BomCompare\BomAndBom_Comp"
 intBomGeShu = 0
 
If bRunBom1 = True Then
  MkDir PrmPath & "BomCompare\Bom_1"
  'MkDir PrmPath & "BomCompare\Bom_2"
  intBomGeShu = intBomGeShu + 1
  
  Call ReadBom1_Ver_Out_Dir
End If

If bRunBom2 = True Then
  MkDir PrmPath & "BomCompare\Bom_1"
  intBomGeShu = intBomGeShu + 1
   Call ReadBom2_Ver_Out_Dir
End If
If bRunBom3 = True Then
  MkDir PrmPath & "BomCompare\Bom_1"
   intBomGeShu = intBomGeShu + 1
  Call ReadBom3_Ver_Out_Dir '
 
End If
If bRunBom4 = True Then
 MkDir PrmPath & "BomCompare\Bom_1"
 intBomGeShu = intBomGeShu + 1
 Call ReadBom4_Ver_Out_Dir
End If
If bRunBom5 = True Then
intBomGeShu = intBomGeShu + 1
  MkDir PrmPath & "BomCompare\Bom_1"
    Call ReadBom5_Ver_Out_Dir
End If
If bRunBom6 = True Then
intBomGeShu = intBomGeShu + 1
  MkDir PrmPath & "BomCompare\Bom_1"
   Call ReadBom6_Ver_Out_Dir
End If

If bRunBom7 = True Then
intBomGeShu = intBomGeShu + 1
  MkDir PrmPath & "BomCompare\Bom_1"
    Call ReadBom7_Ver_Out_Dir
End If
 
If bRunBom8 = True Then
intBomGeShu = intBomGeShu + 1
   MkDir PrmPath & "BomCompare\Bom_1"
  Call ReadBom8_Ver_Out_Dir
End If
'If bRunBom8 = True Then
'   MkDir PrmPath & "BomCompare\Bom_8"
'  Call Bom8Comp_Bom8
'End If
Msg1.Caption = "Start compare..."
Msg2.Caption = "Please wait..."
'Call Start_8Ge_Ver_Bom_Comp

Msg3.Caption = ""
Msg4.Caption = ""
'Call FileCompStart_8_NoTestplan



'�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I
Kill PrmPath & "BomCompare\Bom8GeVer_Comp.txt"
 '�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I�I
 
Kill PrmPath & "BomCompare\BomAndBom_Comp\*.*"
Kill PrmPath & "BomCompare\Bom_1\*.*"
Kill PrmPath & "BomCompare\Bom_2\*.*"
Kill PrmPath & "BomCompare\Bom_3\*.*"
Kill PrmPath & "BomCompare\Bom_4\*.*"
Kill PrmPath & "BomCompare\Bom_5\*.*"
Kill PrmPath & "BomCompare\Bom_6\*.*"
Kill PrmPath & "BomCompare\Bom_7\*.*"
Kill PrmPath & "BomCompare\Bom_8\*.*"
RmDir PrmPath & "BomCompare\Bom_1\"
RmDir PrmPath & "BomCompare\Bom_2\"
RmDir PrmPath & "BomCompare\Bom_3\"
RmDir PrmPath & "BomCompare\Bom_4\"
RmDir PrmPath & "BomCompare\Bom_5\"
RmDir PrmPath & "BomCompare\Bom_6\"
RmDir PrmPath & "BomCompare\Bom_7\"
RmDir PrmPath & "BomCompare\Bom_8\"
RmDir PrmPath & "BomCompare\BomAndBom_Comp"



 


Msg3.Caption = "Runing create BomDiff.txt file..."
Msg4.Caption = "Please wait..."

' Call FileHeBin
 
Msg4.Caption = "Compare file end!"

GOend:
       cmdBom8GeVer.Enabled = True
      Option1.Enabled = True
     Option2.Enabled = True
     Option3.Enabled = True
     Option4.Enabled = True
     txtBom1.Enabled = True
     txtBom2.Enabled = True
     txtBom3.Enabled = True
     txtBom4.Enabled = True
     txtBom5.Enabled = True
     txtBom6.Enabled = True
     txtBom7.Enabled = True
     txtBom8.Enabled = True
 Command1.Enabled = True
 Option1.Enabled = True
 Option2.Enabled = True
 cmdToVerBoard.Enabled = True
 MsgBox "ok"

End Sub

Private Sub cmdBomAndBom_Click()
On Error Resume Next
Dim bTwoBom As Boolean
bTwoBom = bRunBom2 Or bRunBom1
If bRunBom1 = False Then
    MsgBox "Please open bom1 file!", vbCritical
    Exit Sub
End If
If bRunBom2 = False Then
    MsgBox "Please open bom2 file!", vbCritical
    Exit Sub
End If

strMsg = MsgBox(l1.Caption & " file and " & l2.Caption & " file compare,Do you want to continue ?", 52, "Warning!")
If strMsg = vbYes Then
      GoTo Start
   ElseIf strMsg = vbNo Then
    Exit Sub
End If

Start:
cmdToVerBoard.Enabled = -False
Frame3.Enabled = False
cmdToVerBoard.Enabled = False
cmdBomAndBom.Enabled = False
Frame2.Enabled = False
     Option1.Enabled = False
     Option2.Enabled = False
     cmdBom8GeVer.Enabled = False
     Option3.Enabled = False
     Option4.Enabled = False
     Command1.Enabled = False
 
If bTwoBom = True Then
    Open PrmPath & "BomCompare\Bom1_and_Bom2_Compare.txt" For Output As #54
    Close #54
    Call BomAndBomComp_Bom1
    Call BomAndBomComp_Bom2

   
    Call Bom1AndBom2_Dir_Bom2
      Msg1.Caption = l1.Caption & "_"
      Msg2.Caption = "_ and _"
      Msg3.Caption = "_" & l2.Caption
      Msg4.Caption = "_ compare ok!"
      Frame3.Enabled = True
        Call Kill_File
    MsgBox l1.Caption & " and " & l2.Caption & " compare ok!", vbInformation
     Option1.Enabled = True
     Option2.Enabled = True
     cmdToVerBoard.Enabled = True
     Option3.Enabled = True
     Option4.Enabled = True
     cmdToVerBoard.Enabled = True
     Command1.Enabled = True
     cmdBom8GeVer.Enabled = True
     cmdBomAndBom.Enabled = True
     'Option2.Enabled = False
     cmdBom8GeVer.Enabled = True
     Frame2.Enabled = True
End If

End Sub
Private Sub BomAndBomComp_Bom2()
 Dim strBom2_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_2"
intFile_Line = 0
intDevice_Ge = 0
strBom2Path = Trim(txtBom2.Text)
If Dir(strBom2Path) = "" Then
   txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
   bRunBom2 = False
   txtBom2.SetFocus
   strBom2Path = ""
   MsgBox "Bom2 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
   Kill PrmPath & "BomCompare\Bom_2\*.*"
      Open PrmPath & "BomCompare\Bom1_and_Bom2_Compare.txt" For Output As #54
        Print #54, "!============================" & l2.Caption & " file  =True ," & l1.Caption & " file not find devices============================="
   Open strBom2Path For Input As #52
           Do Until EOF(52)
             Line Input #52, strBom2_DeviceName
               Msg1.Caption = "Reading bom2 file..."
               Mystr = LCase(Trim(strBom2_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                 '   strTmp = Split(Mystr, " ")
                    
                               MyStr1 = Replace(Mystr, """", "'")
                             MyStr1 = Replace(MyStr1, "\", "^")
                             MyStr1 = Replace(MyStr1, "/", "^")
                             MyStr1 = Replace(MyStr1, "*", "^")
                             MyStr1 = Replace(MyStr1, "?", "^")
                             MyStr1 = Replace(MyStr1, "<", "[")
                             MyStr1 = Replace(MyStr1, ">", "]")
                    
                    
                    
                    
 '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                     
                               MyStr1 = Replace(MyStr1, " ", ",")
                    strTmp = Split(Replace(MyStr1, Chr(9), ""), ",")
  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    
                    
                    
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                              strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ";", "")
                             If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0)) = "" Then
                                Print #54, strBom2_DeviceName
                             End If
                             
                              Open PrmPath & "BomCompare\Bom_2\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   Print #22, strBom2_DeviceName
                              Close #22
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "AnalogDevice:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read & l2.Caption & file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #52
        Print #54, "!============================" & l1.Caption & " file  =True ," & l2.Caption & " file not find devices============================="
 Close #54
        Msg1.Caption = l2.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom2 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub

Private Sub BomAndBomComp_Bom1()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom1Path = Trim(txtBom1.Text)
If Dir(strBom1Path) = "" Then
   txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
   bRunBom1 = False
   txtBom1.SetFocus
   strBom1Path = ""
   MsgBox "Bom1 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
   Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom1Path For Input As #50
   Open PrmPath & "BomCompare\Bom_1\tmpCompare.dll" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom1 file..."
               Mystr = LCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                  
 '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                              MyStr1 = Replace(Mystr, """", "'")
                             MyStr1 = Replace(MyStr1, "\", "^")
                             MyStr1 = Replace(MyStr1, "/", "^")
                             MyStr1 = Replace(MyStr1, "*", "^")
                             MyStr1 = Replace(MyStr1, "?", "^")
                             MyStr1 = Replace(MyStr1, "<", "[")
                             MyStr1 = Replace(MyStr1, ">", "]")
                            ' strFileNameTmp = Replace(strFileNameTmp, ".", ",")
 
 
                     
                               MyStr1 = Replace(MyStr1, " ", ",")
                               
                    strTmp = Split(Replace(MyStr1, Chr(9), ""), ",")
  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    
                    
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                              strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ";", "")
                              Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   Print #22, strBom1_DeviceName
                              Close #22
                                   'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)   & ".sh"
                                   
                                   Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "AnalogDevice:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read & l1.Caption & file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
        Msg1.Caption = l1.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom1 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub
Private Sub Bom1AndBom2_Dir_Bom2()
Dim MyStr1 As String
    Open PrmPath & "BomCompare\Bom1_and_Bom2_Compare.txt" For Append As #54
 
 Open PrmPath & "BomCompare\Bom_1\tmpCompare.dll" For Input As #59
            Do Until EOF(59)
             Line Input #59, MyStr1
              If MyStr1 <> "" Then
               If Dir(PrmPath & "BomCompare\Bom_2\" & MyStr1) = "" Then
                  Open PrmPath & "BomCompare\Bom_1\" & MyStr1 For Input As #60
                     Line Input #60, TempShit
                  Close #60
                  Print #54, TempShit
               End If
              End If
              DoEvents
            Loop
 
 
 
       Msg2.Caption = "file compare runing..."
       Msg3.Caption = "please wait..."
      Msg4.Caption = l1.Caption & " and " & l2.Caption & " compare ok!"
      
 Close #59
Close #54
End Sub


Private Sub cmdExit_Click()


'Unload frmHelp
Unload frmCreateTestplan
Unload Me
End
End Sub


 

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdOK_Click()
Dim bAllVer As Boolean
Dim a

'MkDir PrmPath & "BomCompare\Testplan_Tmp_Analog"
'MkDir PrmPath & "BomCompare\Bom_1"
'MkDir PrmPath & "BomCompare\Bom_2"
'MkDir PrmPath & "BomCompare\Bom_3"
'MkDir PrmPath & "BomCompare\Bom_4"
'MkDir PrmPath & "BomCompare\Bom_5"
'MkDir PrmPath & "BomCompare\Bom_6"
'MkDir PrmPath & "BomCompare\Bom_7"
'MkDir PrmPath & "BomCompare\Bom_8"

 On Error Resume Next
   bAllVer = bRunBom1 Or bRunBom2 Or bRunBom3 Or bRunBom4 Or bRunBom5 Or bRunBom6 Or bRunBom7 Or bRunBom8
  strTestplanPath = txtTestplan.Text
  If Dir(strTestplanPath) = "" Then ' "Please open testplan file!(DblClick me open file!" Then
     bRunTestplan = False
  End If
  If Check1.Value = 1 Then
    a = True
    Else
     a = False
  End If
  If Check1.Value = 1 Then
   cmdBoards.Enabled = True
   Boards = True
   Else
   cmdBoards.Enabled = False
   Boards = False
End If
  strMsg = MsgBox("Panel Boards=" & a & " ,Do you want to continue ?", 52, "Warning!")
If strMsg = vbYes Then
      GoTo Start
   ElseIf strMsg = vbNo Then
    Exit Sub
End If

Start:
 
 If bRunTestplan = True And bAllVer = True Then
     cmdOk.Enabled = False
     cmdBoards.Enabled = False
     cmdOk.Enabled = False
     Check3.Enabled = False
     Check1.Enabled = False
     Check2.Enabled = False
     Option1.Enabled = False
     Option2.Enabled = False
     Option3.Enabled = False
     txtBom1.Enabled = False
     txtBom2.Enabled = False
     txtBom3.Enabled = False
     txtBom4.Enabled = False
     txtBom5.Enabled = False
     txtBom6.Enabled = False
     txtBom7.Enabled = False
     txtBom8.Enabled = False
     cmdToVerBoard.Enabled = False
     txtTestplan.Enabled = False
     MkDir PrmPath & "BomCompare\Testplan_Tmp_Analog"
   Call Open_Testplan_Read
   Else
    If bRunTestplan = False Then
        MsgBox "Please check testplan path!", vbCritical
        txtTestplan.SetFocus
        Exit Sub
     End If
     If bAllVer = False Then
       MsgBox "Please check bom file!", vbCritical
         txtBom1.SetFocus
         Exit Sub
     End If
     
     
 End If
 If bRunTestplan = False Then
     Call Kill_File
     Call Kill_Device
      Check3.Enabled = True
     cmdOk.Enabled = True
     txtBom1.Enabled = True
     txtBom2.Enabled = True
     txtBom3.Enabled = True
     txtBom4.Enabled = True
     txtBom5.Enabled = True
     txtBom6.Enabled = True
     txtBom7.Enabled = True
     txtBom8.Enabled = True
     cmdBoards.Enabled = True
     Check1.Enabled = True
     txtTestplan.Enabled = True
    Option1.Enabled = True
     Option2.Enabled = True
     Option3.Enabled = True
     
 End If
 
If bRunBom1 = True And bRunTestplan = True Then
  MkDir PrmPath & "BomCompare\Bom_1"
  Call Open_Bom1_Read
End If
If bRunBom2 = True And bRunTestplan = True Then
 MkDir PrmPath & "BomCompare\Bom_2"
  Call Open_Bom2_Read
End If
If bRunBom3 = True And bRunTestplan = True Then
  MkDir PrmPath & "BomCompare\Bom_3"
  Call Open_Bom3_Read
End If
If bRunBom4 = True And bRunTestplan = True Then
  MkDir PrmPath & "BomCompare\Bom_4"
  Call Open_Bom4_Read
End If

If bRunBom5 = True And bRunTestplan = True Then
  MkDir PrmPath & "BomCompare\Bom_5"
  Call Open_Bom5_Read
End If
If bRunBom6 = True And bRunTestplan = True Then
  MkDir PrmPath & "BomCompare\Bom_6"
  Call Open_Bom6_Read
End If
If bRunBom7 = True And bRunTestplan = True Then
   MkDir PrmPath & "BomCompare\Bom_7"
  Call Open_Bom7_Read
End If
If bRunBom8 = True And bRunTestplan = True Then
   MkDir PrmPath & "BomCompare\Bom_8"
  Call Open_Bom8_Read
End If

Call Kill_Device



Msg1.Caption = "Start compare..."
Msg2.Caption = ""
Msg3.Caption = ""
Msg4.Caption = ""

 Kill PrmPath & "BomCompare\Bom_1_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_2_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_3_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_4_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_5_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_6_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_7_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_8_Exist.txt"
 Kill PrmPath & "BomCompare\Comm_Device_Exist.txt"
 Kill PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt"
 Kill PrmPath & "BomCompare\Testplan_Analog.txt"


Call File_Compare


Msg3.Caption = "Runing create sub_analog.txt file..."
Msg4.Caption = "Please wait..."

Call Create_Sub_analog_file
Call Kill_File
 

' Kill PrmPath & "BomCompare\Testplan_Analog.txt"
'  Kill PrmPath & "BomCompare\Bom_1_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_2_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_3_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_4_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_5_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_6_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_7_Exist.txt"
' Kill PrmPath & "BomCompare\Bom_8_Exist.txt"
' Kill PrmPath & "BomCompare\Comm_Device_Exist.txt"
' Kill PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt"
 
Msg4.Caption = "Compare analog file end!"
cmdToVerBoard.Enabled = True
     cmdOk.Enabled = True
     Check3.Enabled = True
     cmdBoards.Enabled = True
  Option1.Enabled = True
     Option2.Enabled = True
     Option3.Enabled = True
     txtBom1.Enabled = True
     txtBom2.Enabled = True
     txtBom3.Enabled = True
     txtBom4.Enabled = True
     txtBom5.Enabled = True
     txtBom6.Enabled = True
     txtBom7.Enabled = True
     txtBom8.Enabled = True
     Check1.Enabled = True
     Check2.Enabled = True
     txtTestplan.Enabled = True
     MsgBox Me.Caption, vbInformation
End Sub
Private Sub File_Compare()
Dim MyName
Dim Mystr As String
Dim bBom1OK As Boolean
Dim bBom2OK As Boolean
Dim bBom3OK As Boolean
Dim bBom4OK As Boolean
Dim bBom5OK As Boolean
Dim bBom6OK As Boolean
Dim bBom7OK As Boolean
Dim bBom8OK As Boolean
Dim bTestorder As Boolean
Dim bUnString As Boolean
Dim strTmpWait() As String
Dim intI As Integer
'dim MyWaitStr as string
 On Error Resume Next
If bRunBom1 = True Then
    Open PrmPath & "BomCompare\Bom_1_Exist.txt" For Output As #5
    Close #5
End If
If bRunBom2 = True Then

   Open PrmPath & "BomCompare\Bom_2_Exist.txt" For Output As #5
   Close #5
End If
If bRunBom3 = True Then

   Open PrmPath & "BomCompare\Bom_3_Exist.txt" For Output As #5
   Close #5
End If
If bRunBom4 = True Then

   Open PrmPath & "BomCompare\Bom_4_Exist.txt" For Output As #5
   Close #5
End If
If bRunBom5 = True Then

   Open PrmPath & "BomCompare\Bom_5_Exist.txt" For Output As #5
   Close #5
End If
If bRunBom6 = True Then

   Open PrmPath & "BomCompare\Bom_6_Exist.txt" For Output As #5
   Close #5
End If
If bRunBom7 = True Then

   Open PrmPath & "BomCompare\Bom_7_Exist.txt" For Output As #5
   Close #5
End If
If bRunBom8 = True Then
   Open PrmPath & "BomCompare\Bom_8_Exist.txt" For Output As #5
   Close #5
End If
   Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Output As #5
   Close #5
'   Open PrmPath & "BomCompare\sub_analog.txt" For Output As #5
'   Close #5
   Open PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt" For Append As #5
   Close #5

   
 MyName = Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\*.*")
'If MyName <> "" Then
Open PrmPath & "BomCompare\Testplan_Analog.txt" For Output As #7
  Do While MyName <> ""   ' ��ʼѭ��
     tmpfile = Trim(Replace(MyName, ".sh", ""))
        If Not_initializel_testplan = False Then
           If InStr(LCase(Mystr), ".unknownstring") <> 0 Then
               tmpfile = ""
           End If
        End If

     If tmpfile <> "" Then
       Print #7, tmpfile
       Msg2.Caption = "Read:" & Testplan_Analog.txt & " file"
       Msg3.Caption = tmpfile
     End If
      Debug.Print MyName
      
    MyName = Dir   ' ������һ��Ŀ¼��
    DoEvents
  Loop
Close #7





  Kill PrmPath & "BomCompare\Testplan_Tmp_Analog\*.sh"
 Kill PrmPath & "BomCompare\Testplan_Tmp_Analog\*.UnknownString"
  Open PrmPath & "BomCompare\Testplan_Analog.txt" For Input As #7
    Do Until EOF(7)
      Line Input #7, Mystr
      Mystr = Trim(LCase(Mystr))
        If Mystr <> "" Then
           tmpFile2 = Mystr
            If InStr(Mystr, ".%") <> 0 Then
               tmpFile2 = Trim(Replace(Mystr, ".%", ""))
                Mystr = Trim(Replace(Mystr, ".%", ""))
               tmpFile2 = Left(Mystr, Len(Mystr) - (Len(Mystr) - InStr(Mystr, "%")))
                tmpFile2 = Trim(Replace(tmpFile2, "%", ""))
              'baobei 2
               tmpFile2 = Trim(tmpFile2)
               Else
                If InStr(Mystr, "._") <> 0 Then
                     Mystr = Trim(Replace(Mystr, "._", ""))
                    tmpFile2 = Trim(Replace(Mystr, "._", ""))
                    tmpFile2 = Left(Mystr, Len(Mystr) - (Len(Mystr) - InStr(Mystr, "_")))
                    tmpFile2 = Trim(Replace(tmpFile2, "_", ""))
                    'baobei
                    tmpFile2 = Trim(tmpFile2)
                End If
                If InStr(Mystr, ".testcommentedintestorder") <> 0 Then
                   Mystr = Trim(Replace(Mystr, ".testcommentedintestorder", ""))
                    tmpFile2 = Mystr
                    tmpFile2 = Trim(tmpFile2)
                    bTestorder = True
                End If
                If Not_initializel_testplan = True Then
                  
                    If InStr(Mystr, ".unknownstring") <> 0 Then
                         tmpFile2 = ""
                       Mystr = Trim(Replace(Mystr, ".unknownstring", ""))
                          Open PrmPath & "BomCompare\WaitDevice.txt" For Input As #25
                             Do Until EOF(25)
                                Line Input #25, MyWaitStr
                                  MyWaitStr = Trim(LCase(MyWaitStr))
                                    If MyWaitStr <> "" Then
                                       strTmpWait = Split(MyWaitStr, ",")
                                         If strTmpWait(0) = Mystr Then
                                            tmpFile2 = strTmpWait(0)
                                            tmpFile2 = Trim(tmpFile2)

                                            
                                            Exit Do
                                         End If
                                    End If
                                    DoEvents
                             Loop
                           Close #25
                           
                             If tmpFile2 <> "" Then
                               bUnString = True
                            End If
                           
                    End If



                End If
            End If
               
        
        
        
            If Dir(PrmPath & "BomCompare\Bom_1\" & tmpFile2) <> "" Then
                bBom1OK = True
               Else
                bBom1OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_2\" & tmpFile2) <> "" Then
               bBom2OK = True
               Else
               bBom2OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_3\" & tmpFile2) <> "" Then
               bBom3OK = True
               Else
               bBom3OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_4\" & tmpFile2) <> "" Then
               bBom4OK = True
               Else
               bBom4OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_5\" & tmpFile2) <> "" Then
               bBom5OK = True
               Else
               bBom5OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_6\" & tmpFile2) <> "" Then
               bBom6OK = True
               Else
               bBom6OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_7\" & tmpFile2) <> "" Then
               bBom7OK = True
               Else
               bBom7OK = False
            End If
            If Dir(PrmPath & "BomCompare\Bom_8\" & tmpFile2) <> "" Then
               bBom8OK = True
               Else
               bBom8OK = False
            End If
            If bRunBom1 = False Then bBom1OK = True
            If bRunBom2 = False Then bBom2OK = True
            If bRunBom3 = False Then bBom3OK = True
            If bRunBom4 = False Then bBom4OK = True
            If bRunBom5 = False Then bBom5OK = True
            If bRunBom6 = False Then bBom6OK = True
            If bRunBom7 = False Then bBom7OK = True
            If bRunBom8 = False Then bBom8OK = True



    '  'start com
             If (bBom1OK And bRunBom1) = True And (bBom2OK And bBom3OK And bBom4OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                             If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_1_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_1_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                          Else
                                Open PrmPath & "BomCompare\Bom_1_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_1_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5
                 End If
             End If
    
             If (bBom2OK And bRunBom2) = True And (bBom1OK And bBom3OK And bBom4OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_2_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_2_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                            
                          Else
                                Open PrmPath & "BomCompare\Bom_2_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_2_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5
                 End If
             End If
    
             If (bBom3OK And bRunBom3) = True And (bBom1OK And bBom2OK And bBom4OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_3_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_3_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                          Else
                                Open PrmPath & "BomCompare\Bom_3_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                          
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_3_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5

                 End If
             End If
    
             If (bBom4OK And bRunBom4) = True And (bBom1OK And bBom3OK And bBom2OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_4_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_4_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                         Else
                                Open PrmPath & "BomCompare\Bom_4_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            
                            
                            
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_4_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5

                 End If
             End If
             If (bBom5OK And bRunBom5) = True And (bBom1OK And bBom3OK And bBom4OK And bBom2OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_5_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_5_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                           Else
                                Open PrmPath & "BomCompare\Bom_5_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_5_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5

                 End If
             End If
             If (bBom6OK And bRunBom6) = True And (bBom1OK And bBom3OK And bBom4OK And bBom2OK And bBom5OK And bBom7OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_6_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                                
                                
                                
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_6_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                                
                                
                            End If
                          Else
                                Open PrmPath & "BomCompare\Bom_6_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_6_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5

                 End If
             End If
             If (bBom7OK And bRunBom7) = True And (bBom1OK And bBom3OK And bBom4OK And bBom2OK And bBom6OK And bBom5OK And bBom8OK) = False Then
                 If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_7_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_7_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                            
                           Else
                                Open PrmPath & "BomCompare\Bom_7_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_7_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5

                 End If
             End If
             If (bBom8OK And bRunBom8) = True And (bBom1OK And bBom3OK And bBom4OK And bBom2OK And bBom6OK And bBom5OK And bBom7OK) = False Then
                If bTestorder <> True Then
                        If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\Bom_8_Exist.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    " & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\Bom_8_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                          Else
                                Open PrmPath & "BomCompare\Bom_8_Exist.txt" For Append As #5
                                  Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\Bom_8_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5
                 End If
             End If
    
    
             If (bBom1OK And bBom2OK And bBom3OK And bBom4OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = True Then
                
                If bTestorder <> True Then
                
                        
                        If Not_initializel_testplan = True Then
                             If bUnString = True Then
                                 Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Append As #5
                                      If strTmpWait(1) <> "" Then
                                           Print #5, "    " & strTmpWait(1)
                                        Else
                                           Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                           
                                               Print #6, MyWaitStr & "  " & Now
                                           Close #6
                                      End If
                                 Close #5
                               Else
                            
                                 Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Append As #5
                                   Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                 Close #5
                             End If
                           Else
                                 Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Append As #5
                                   Print #5, "    test " & """" & strAnalog_ & Mystr & """"
                                 Close #5
                            
                            
                        End If
'
'
'                     Else
'                                Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Append As #5
'                                  If Boards = True Then
'                                   'strBoardsNumber
'                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
'                                    Else
'                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
'                                  End If
'                                Close #5
                End If
                
             End If
             
             If (bBom1OK And bBom2OK And bBom3OK And bBom4OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                If bTestorder <> True Then

                         If Not_initializel_testplan = True Then
                            If bUnString = True Then
                                Open PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt" For Append As #5
                                     If strTmpWait(1) <> "" Then
                                          Print #5, "    !N" & strTmpWait(1)
                                       Else
                                          Open PrmPath & "BomCompare\Error.txt" For Append As #6
                                          
                                              Print #6, MyWaitStr & "  " & Now
                                          Close #6
                                     End If
                                Close #5
                              Else
                           
                                Open PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt" For Append As #5
                                 Print #5, "    !N test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            End If
                           Else
                                Open PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt" For Append As #5
                                 Print #5, "    !N test " & """" & strAnalog_ & Mystr & """"
                                Close #5
                            
                        End If
        
                           
                     Else
                                Open PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   !N test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   !N test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5
                  End If 'testorder
             End If
             
             
            Msg2.Caption = "Current device:" & Mystr
         End If
             bTestorder = False
             strTmpWait(1) = ""
             bUnString = False
             strTmpWait(0) = ""
             strTmpWait(1) = ""
        DoEvents
        
        
    Loop
  Close #7
    
    Msg4.Caption = "Compare ok!"

End Sub

Private Sub Open_Testplan_Read()
Dim Mystr As String
Dim strAnalogName As String
Dim bSubAnalog As Boolean
Dim i 'As Integer
Dim t 'As Integer
Dim BoardSetNomber() As String
 On Error Resume Next
Kill PrmPath & "BomCompare\Testplan_Tmp_Analog\*.*"

strTestplanPath = Trim(txtTestplan.Text)
If Dir(strTestplanPath) = "" Then
   txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
   bRunTestplan = False
   strTestplanPath = ""
   txtTestplan.SetFocus
   MsgBox "Testplan not find!,please check path!", vbCritical
   Exit Sub
End If
i = 0
t = 0
'Open PrmPath & "BomCompare\EspeciallyDevice.txt" For Output As #2
'Close #2
'open testplan file
Msg1.Caption = "Reading testplan file..."

'  Open PrmPath & "BomCompare\TestorderClose.txt" For Output As #23
'  Close #23
  Open PrmPath & "BomCompare\UnsettledDevice.txt" For Output As #23
  Close #23
  Open PrmPath & "BomCompare\Kill_Device" For Output As #4
  Close #4
  If Not_initializel_testplan = True Then
      Open PrmPath & "BomCompare\WaitDevice.txt" For Output As #23
     
      Close #23
  End If
  
  
strBoardsNumber = ""
  Open strTestplanPath For Input As #1
   Do Until EOF(1)
       Line Input #1, Mystr
         Mystr = Trim(Mystr)
       If Mystr <> "" Then
              If bSubAnalog = True And Trim(LCase(Mystr)) = "subend" Then
                                      Msg1.Caption = "Testplan file read ok!"
                                      Msg2.Caption = "Find subend ok"
                                      Msg3.Caption = "Device(+  % rX , % crX)  :" & t
                                      t = 0
                                      Msg4.Caption = ""
                                      
                                    Exit Do
              End If
                     
              If bSubAnalog = True And Left(Trim(LCase(Mystr)), 6) = "subend" Then
                                      Msg1.Caption = "Testplan file read ok!"
                                      Msg2.Caption = "Find subend ok"
                                      Msg3.Caption = "Device(+  % rX , % crX)  :" & t
                                      t = 0
                                      Msg4.Caption = ""
                                      
                                    Exit Do
              End If
                     
                     
             
                If Boards = True Then
                   If InStr(Replace(LCase(Mystr), " ", ""), "subanalog_tests") <> 0 Then
                       'Mystr = Replace(Mystr, " ", "")
                       'Mystr = Replace(LCase(Mystr), "*", "$")
                        BoardSetNomber = Split(Mystr, ",")
                           If BoardSetNomber(0) <> "" Then
                                bSubAnalog = True
                           End If
                           If BoardSetNomber(1) <> "" Then
                              strBoardsNumber = BoardSetNomber(1)
                           End If
                                          
                   
                                          
                   
                           If strBoardsNumber = "" Then
                                   MsgBox "The testplan not vb create boards testplan!", vbCritical
                                   bSubAnalog = False
                                   bRunTestplan = False
                                Exit Do
                              
                           End If 'strBoardsNumber = ""
                     End If 'InStr(Replace(LCase(Mystr), " ", ""), "subanalog_tests") <> 0
                 End If 'Boards = True
                 If Left(Mystr, 1) = "!" And InStr(Mystr, "test ") <> 0 And bSubAnalog = True And InStr(Mystr, strAnalog_) <> 0 And InStr(Replace(Mystr, " ", ""), "testcommentedintestorder") <> 0 Then
            '                    Open PrmPath & "BomCompare\TestorderClose.txt" For Append As #23
            '                       Print #23, Mystr
            '                    Close #23
                            strAnalogName = ""
                            strAnalogName = Replace(Mystr, " ", "")
                            strAnalogName = Replace(strAnalogName, "testcommentedintestorder", "")
                            strAnalogName = Replace(strAnalogName, Left(strAnalogName, (InStr(strAnalogName, "test") - 1)), "")
                            strAnalogName = Trim(Replace(strAnalogName, "test", ""))
                            strAnalogName = Right(strAnalogName, Len(strAnalogName) - 1)
                            strAnalogName = Trim(Replace(strAnalogName, strAnalog_, ""))
                            strAnalogName = Left(strAnalogName, InStr(strAnalogName, """"))
                            strAnalogName = Trim(LCase(Trim(Replace(strAnalogName, """", ""))))
                            'baobei weiwei3007
                             ' strAnalogName = LCase(Trim(Replace(strAnalogName, "_", "")))
                              ' strAnalogName = LCase(Trim(Replace(strAnalogName, "%", "")))


                             'create analog device file
                             

                             
                         If strAnalogName = "" Then
                              Open PrmPath & "BomCompare\UnsettledDevice.txt" For Append As #23
                                    Print #23, Mystr
                              Close #23
                            Else
                             Open PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strAnalogName & ".testcommentedintestorder" For Output As #2
                             Close #2
                              t = t + 1
                              Msg4.Caption = strAnalogName
                             
                             
                         End If

                         
                         strAnalogName = ""
                         
                       Else
                          If Not_initializel_testplan = False And Left(Mystr, 1) = "!" And bSubAnalog = True And InStr(Mystr, strAnalog_) <> 0 Then
             
                                    Open PrmPath & "BomCompare\UnsettledDevice.txt" For Append As #23
                                              Print #23, Mystr
                                     Close #23
                       
                          End If
                         If Left(Mystr, 1) = "!" And bSubAnalog = True And InStr(Mystr, strAnalog_) <> 0 Then
                                 
                                 
                            If Not_initializel_testplan = True Then
                                 strAnalogName = Replace(Mystr, " ", "")
                                 strAnalogName = Replace(strAnalogName, Left(strAnalogName, (InStr(strAnalogName, "test") - 1)), "")
                                 strAnalogName = Trim(Replace(strAnalogName, "test", ""))
                                 strAnalogName = Right(strAnalogName, Len(strAnalogName) - 1)
                                 strAnalogName = Trim(Replace(strAnalogName, strAnalog_, ""))
                                 strAnalogName = Trim(Left(strAnalogName, InStr(strAnalogName, """")))
                                 strAnalogName = Trim(LCase(Trim(Replace(strAnalogName, """", ""))))
                                  strAnalogName = LCase(Trim(Replace(strAnalogName, "%", "")))
                                  'baobei
                                    If strAnalogName <> "" Then
                                            Open PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strAnalogName & ".UnknownString" For Output As #2
                                            Close #2
                                             t = t + 1
                                             Msg4.Caption = strAnalogName
                                            
                                             Open PrmPath & "BomCompare\WaitDevice.txt" For Append As #23
                                                Print #23, strAnalogName & "," & Mystr
                                             Close #23
                                             
                                             strAnalogName = ""
                                        Else
                                              Open PrmPath & "BomCompare\UnsettledDevice.txt" For Append As #23
                                                    Print #23, Mystr
                                             Close #23
                                             
                                         
                                    End If 'strAnalogName <> ""

                           End If 'Not_initializel_testplan = True
                        Else
                          If Left(Mystr, 1) = "!" And Not_initializel_testplan = True And bSubAnalog = True And InStr(Mystr, strAnalog_) <> 0 Then
                            Open PrmPath & "BomCompare\UnsettledDevice.txt" For Append As #23
                                      Print #23, Mystr
                             Close #23
                          End If
                           
                     End If 'Left(Mystr, 1) = "!" And bSubAnalog = True And InStr(Mystr, strAnalog_) <> 0
                  End If
                  If Left(Mystr, 1) <> "!" Then
                        If InStr(Replace(LCase(Mystr), " ", ""), "subanalog_tests") <> 0 Then
                           bSubAnalog = True
                        End If
                        If bSubAnalog = True Then
                                 If Trim(LCase(Mystr)) = "subend" Then
                                      Msg1.Caption = "Testplan file read ok!"
                                      Msg2.Caption = "Find subend ok"
                                      Msg3.Caption = "Device(+  % rX , % crX)  :" & t
                                      t = 0
                                      Msg4.Caption = ""
                                      
                                    Exit Do
                                 End If
                                Msg3.Caption = "sub Analog_Tests find ok!"
                               If Left(LCase(Mystr), 5) = "test " And InStr(LCase(Mystr), strAnalog_) <> 0 Then
                                    ' strAnalogName = Replace(MyStr, " ", "")
                                        strAnalogName = ""
                                        strAnalogName = Replace(Mystr, " ", "")
                                        strAnalogName = Trim(Replace(strAnalogName, "test", ""))
                                        strAnalogName = Right(strAnalogName, Len(strAnalogName) - 1)
                                        strAnalogName = Trim(Replace(strAnalogName, strAnalog_, ""))
                                        strAnalogName = Left(strAnalogName, InStr(strAnalogName, """"))
                                        strAnalogName = Trim(LCase(Trim(Replace(strAnalogName, """", ""))))
                                       'strAnalogName = Mid(MyStr, InStr(MyStr, strAnalog_) + 8, InStr(InStr(MyStr, ""strAnalog_), MyStr, """"))
                                     If strAnalogName <> "" Then
                                              If InStr(strAnalogName, "%") <> 0 Then
                                                          Open PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strAnalogName & ".%" For Output As #4
                                                          Close #4
                    '                                      Open PrmPath & "BomCompare\EspeciallyDevice.txt" For Append As #2
                    '                                         Print #2, strAnalogName
                    '                                      Close #2
                                                          
                                                         strAnalogName = Replace(Left(strAnalogName, Len(strAnalogName) - (Len(strAnalogName) - InStr(strAnalogName, "%"))), "%", "")
                                                         Open PrmPath & "BomCompare\Kill_Device" For Append As #4
                                                            Print #4, strAnalogName & ".sh"
                                                         Close #4
                                                   Else
                                                      If InStr(strAnalogName, "_") <> 0 Then
                                                              Open PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strAnalogName & "._" For Output As #5
                                                               Close #5
                                                               
                    '                                            Open PrmPath & "BomCompare\EspeciallyDevice.txt" For Append As #2
                    '                                               Print #2, strAnalogName
                    '                                            Close #2
                                                              
                                                              strAnalogName = Replace(Left(strAnalogName, Len(strAnalogName) - (Len(strAnalogName) - InStr(strAnalogName, "_"))), "_", "")
                                                                Open PrmPath & "BomCompare\Kill_Device" For Append As #4
                                                                   Print #4, strAnalogName & ".sh"
                                                                Close #4

                                                      End If 'InStr(strAnalogName, "_") <> 0
                                               End If '<>"%"
                                           strAnalogName = LCase(Trim(Replace(strAnalogName, "_", "")))
                                              strAnalogName = LCase(Trim(Replace(strAnalogName, "%", "")))
                                            '  baobei
                                        
                                             'create analog device file
                                             Open PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strAnalogName & ".sh" For Output As #2
                                             Close #2
                                              t = t + 1
                                              Msg4.Caption = strAnalogName
                                              strAnalogName = ""
'                                         Else
'                                            Open PrmPath & "BomCompare\UnsettledDevice.txt" For Append As #23
'                                                    Print #23, Mystr
'                                            Close #23
                                    End If ' strAnalogName <> ""
                                End If 'Left(LCase(Mystr), 5) = "test " And InStr(LCase(Mystr), ""strAnalog_)
                         End If  ' bRunTestplan
                
                 
                      strAnalogName = ""
            End If '<> !
   End If '<>""
      i = i + 1
      strAnalogName = ""
      Msg2.Caption = "Read file line:" & i
    DoEvents
Loop
  Close #1
     If bSubAnalog = False Then
        Msg1.Caption = "The testplan is bad file!"
        Msg2.Caption = "The testplan not find sub analog!"
        Msg3.Caption = "shit!"
        Msg4.Caption = ""
        bRunTestplan = False
        
     End If
   Msg4.Caption = "Testplan closed!"
   bSubAnalog = False
End Sub


Private Sub Open_Bom1_Read()
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
On Error Resume Next
strBom1Path = Trim(txtBom1.Text)
If Dir(strBom1Path) = "" Then
   txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
   bRunBom1 = False
   txtBom1.SetFocus
   strBom1Path = ""
   MsgBox "Bom1 not find!,please check path!", vbCritical
   Exit Sub
End If
'List1.Clear
'open bom1 file
   'open PrmPath & "BomCompare\Bom_1"
   Kill PrmPath & "BomCompare\Bom_1\*.*"
Open strBom1Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading bom1 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create bom1 analog file
                      Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                   '  Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read bom1 file line: " & i
       DoEvents
   Loop
Close #1
Msg1.Caption = "Bom1 file closed!"
If i = 0 Then
  MsgBox "Shit ,the bom1 file is null!", vbCritical
  Exit Sub
End If
 
 
End Sub
Private Sub Open_Bom2_Read()
On Error Resume Next
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
strBom2Path = Trim(txtBom2.Text)
If Dir(strBom2Path) = "" Then
   txtBom2.Text = " Please open Bom2 file!(DblClick me open file!)"
   bRunBom2 = False
   txtBom2.SetFocus
   strBom2Path = ""
   MsgBox "Bom2 not find!,please check path!", vbCritical
   Exit Sub
End If
   Kill PrmPath & "BomCompare\Bom_2\*.*"

'List2.Clear
'open Bom2 file
   'open PrmPath & "BomCompare\Bom_1"
Open strBom2Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom2 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom2 analog file
                      Open PrmPath & "BomCompare\Bom_2\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                     Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom2 file line: " & i
       DoEvents
   Loop
Close #1
 Msg1.Caption = "Bom2 file closed!"
 If i = 0 Then
  MsgBox "Shit ,the bom2 file is null!", vbCritical
  Exit Sub
End If
End Sub
Private Sub Open_Bom3_Read()
On Error Resume Next
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
strBom3Path = Trim(txtBom3.Text)
If Dir(strBom3Path) = "" Then
   txtBom3.Text = " Please open Bom3 file!(DblClick me open file!)"
   bRunBom3 = False
   txtBom3.SetFocus
   strBom3Path = ""
   MsgBox "Bom3 not find!,please check path!", vbCritical
   Exit Sub
End If
'List2.Clear
'open Bom3 file
   'open PrmPath & "BomCompare\Bom_1"
      Kill PrmPath & "BomCompare\Bom_3\*.*"

Open strBom3Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom3 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom3 analog file
                      Open PrmPath & "BomCompare\Bom_3\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                     Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom3 file line: " & i
       DoEvents
   Loop
Close #1
 Msg1.Caption = "Bom3 file closed!"
 If i = 0 Then
  MsgBox "Shit ,the bom3 file is null!", vbCritical
  Exit Sub
End If
 
 
End Sub
Private Sub Open_Bom4_Read()
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
On Error Resume Next
strBom4Path = Trim(txtBom4.Text)
If Dir(strBom4Path) = "" Then
   txtBom4.Text = " Please open Bom4 file!(DblClick me open file!)"
   bRunBom4 = False
   txtBom4.SetFocus
   strBom4Path = ""
   MsgBox "Bom4 not find!,please check path!", vbCritical
   Exit Sub
End If
'List2.Clear
'open Bom4 file
   'open PrmPath & "BomCompare\Bom_1"
      Kill PrmPath & "BomCompare\Bom_4\*.*"

Open strBom4Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom4 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom4 analog file
                      Open PrmPath & "BomCompare\Bom_4\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                     Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom4 file line: " & i
       DoEvents
   Loop
Close #1
 Msg1.Caption = "Bom4 file closed!"
 If i = 0 Then
  MsgBox "Shit ,the bom4 file is null!", vbCritical
  Exit Sub
End If
End Sub


Private Sub Open_Bom5_Read()
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
On Error Resume Next
strBom5Path = Trim(txtBom5.Text)
If Dir(strBom5Path) = "" Then
   txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
   bRunBom5 = False
   txtBom5.SetFocus
   strBom5Path = ""
   MsgBox "Bom5 not find!,please check path!", vbCritical
   Exit Sub
End If
'List1.Clear
'open Bom5 file
   'open PrmPath & "BomCompare\Bom_1"
   Kill PrmPath & "BomCompare\Bom_5\*.*"
Open strBom5Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom5 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom5 analog file
                      Open PrmPath & "BomCompare\Bom_5\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                   '  Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom5 file line: " & i
       DoEvents
   Loop
Close #1
Msg1.Caption = "Bom5 file closed!"
If i = 0 Then
  MsgBox "Shit ,the Bom5 file is null!", vbCritical
  Exit Sub
End If
 
 
End Sub
Private Sub Open_Bom6_Read()
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
On Error Resume Next
strBom6Path = Trim(txtBom6.Text)
If Dir(strBom6Path) = "" Then
   txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
   bRunBom6 = False
   txtBom6.SetFocus
   strBom6Path = ""
   MsgBox "Bom6 not find!,please check path!", vbCritical
   Exit Sub
End If
'List1.Clear
'open Bom6 file
   'open PrmPath & "BomCompare\Bom_1"
   Kill PrmPath & "BomCompare\Bom_6\*.*"
Open strBom6Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom6 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom6 analog file
                      Open PrmPath & "BomCompare\Bom_6\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                   '  Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom6 file line: " & i
       DoEvents
   Loop
Close #1
Msg1.Caption = "Bom6 file closed!"
If i = 0 Then
  MsgBox "Shit ,the Bom6 file is null!", vbCritical
  Exit Sub
End If
 
 
End Sub
Private Sub Open_Bom7_Read()
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
On Error Resume Next
strBom7Path = Trim(txtBom7.Text)
If Dir(strBom7Path) = "" Then
   txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
   bRunBom7 = False
   txtBom7.SetFocus
   strBom7Path = ""
   MsgBox "Bom7 not find!,please check path!", vbCritical
   Exit Sub
End If
'List1.Clear
'open Bom7 file
   'open PrmPath & "BomCompare\Bom_1"
   Kill PrmPath & "BomCompare\Bom_7\*.*"
Open strBom7Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom7 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom7 analog file
                      Open PrmPath & "BomCompare\Bom_7\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                   '  Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom7 file line: " & i
       DoEvents
   Loop
Close #1
Msg1.Caption = "Bom7 file closed!"
If i = 0 Then
  MsgBox "Shit ,the Bom7 file is null!", vbCritical
  Exit Sub
End If
 
 
End Sub

Private Sub Open_Bom8_Read()
Dim Mystr As String
Dim strTmp() As String
Dim i
Dim t
On Error Resume Next
strBom8Path = Trim(txtBom8.Text)
If Dir(strBom8Path) = "" Then
   txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
   bRunBom8 = False
   txtBom8.SetFocus
   strBom8Path = ""
   MsgBox "Bom8 not find!,please check path!", vbCritical
   Exit Sub
End If
'List1.Clear
'open Bom8 file
   'open PrmPath & "BomCompare\Bom_1"
   Kill PrmPath & "BomCompare\Bom_8\*.*"
Open strBom8Path For Input As #1
   Do Until EOF(1)
     Line Input #1, Mystr
       Msg1.Caption = "Reading Bom8 file..."
       Mystr = LCase(Trim(Mystr))
       If Mystr <> "" Then
          If Left(Mystr, 1) <> "-" Then
            strTmp = Split(Mystr, " ")
              If Trim(strTmp(UBound(strTmp))) <> "" Then
                 If Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strTmp(UBound(strTmp))) & ".*" <> "" Then
                     'create Bom8 analog file
                      Open PrmPath & "BomCompare\Bom_8\" & strTmp(UBound(strTmp)) For Output As #2
                           
                      Close #2
                     t = t + 1
                     Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                     Msg4.Caption = "AnalogDevice:" & t
                   '  Debug.Print strTmp(UBound(strTmp))
                 End If
                 

                 
              End If
          End If '<>-
          
       End If '<>""
       i = i + 1
       Msg2.Caption = "Read Bom8 file line: " & i
       DoEvents
   Loop
Close #1
Msg1.Caption = "Bom8 file closed!"
If i = 0 Then
  MsgBox "Shit ,the Bom8 file is null!", vbCritical
  Exit Sub
End If
End Sub

 

Private Sub cmdToVerBoard_Click()
frmBomValue.Show
End Sub

Private Sub Command1_Click()

frmLibEdit.Show
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
 
  
   If App.PrevInstance Then
     MsgBox "The application is already open", vbInformation, "Error"

   End
   Exit Sub
   
  End If

PrmPath = App.Path
If Right(PrmPath, 1) <> "\" Then PrmPath = PrmPath & "\"
MkDir PrmPath & "BomCompare"


If Option1.Value = True Then
   Frame4.Enabled = True
   Else
   Frame4.Enabled = False
End If
Open PrmPath & "BomCompare\NotDelete.sys" For Output As #77
If Check2.Value = 1 Then
  Not_initializel_testplan = True
  Else
  Not_initializel_testplan = False
End If
If Check3.Value = 1 Then
     
     strAnalog_ = ""
   Else
    
    strAnalog_ = "analog/"
    
End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
'txtBom1.Width = Me.Width - 380
' txtbom2.Width = txtBom1.Width
' txtBom3.Width = txtBom1.Width
' txtBom4.Width = txtBom1.Width
' txtTestplan.Width = txtBom1.Width
'  txtBomTxt1.Width = Me.Width / 2 ' + 2000
' txtBomTxt1.Height = Me.ScaleHeight - txtBomTxt1.Top - 150
' Dim tmpLeft
' tmpLeft = Me.Width
' tmpLeft = (tmpLeft - 120 * 5) / 4
' List1.Width = tmpLeft
' List2.Width = List1.Width
' List3.Width = List1.Width
' List4.Width = List1.Width
' List2.Left = List1.Width + 240
' List3.Left = List1.Width + List2.Width + 360
' List4.Left = List1.Width + List2.Width + List3.Width + 480
' l1.Width = List1.Width
' l1.Left = List1.Left
'  l2.Width = List2.Width
' l2.Left = List2.Left
'
'  l3.Width = List3.Width
' l3.Left = List3.Left
'  l4.Width = List4.Width
' l4.Left = List4.Left
' comOK.Left = txtBomTxt1.Width + 240
' Msg1.Left = comOK.Left
' Msg2.Left = comOK.Left
' Msg3.Left = comOK.Left
' Msg4.Left = comOK.Left
' comOK.Width = Me.Width - txtBomTxt1.Width - 240 - 200
'  Msg1.Width = comOK.Width
' Msg2.Width = comOK.Width
' Msg3.Width = comOK.Width
'  Msg4.Width = comOK.Width
' comOK.Height = txtBomTxt1.Height / 2 + 100
''
'     txtBom2.Enabled = False
'    txtBom3.Enabled = False
'    txtBom4.Enabled = False
'    txtBom5.Enabled = False
'    txtBom6.Enabled = False
'    txtBom7.Enabled = False
'    txtBom8.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Close #77
Call Kill_File




RmDir PrmPath & "BomCompare\Testplan_Tmp_Analog"
RmDir PrmPath & "BomCompare\Bom_1"
RmDir PrmPath & "BomCompare\Bom_2"
RmDir PrmPath & "BomCompare\Bom_3"
RmDir PrmPath & "BomCompare\Bom_4"
RmDir PrmPath & "BomCompare\Bom_5"
RmDir PrmPath & "BomCompare\Bom_6"
RmDir PrmPath & "BomCompare\Bom_7"
RmDir PrmPath & "BomCompare\Bom_8"

 Kill PrmPath & "BomCompare\NotDelete.sys"
Unload frmCreateTestplan
Unload frmHelp
Unload Me
End
End Sub

Private Sub Kill_Device()
On Error Resume Next
Dim Mystr
  Open PrmPath & "BomCompare\Kill_Device" For Input As #44
     Do Until EOF(44)
        Line Input #44, Mystr
        Mystr = Trim(Mystr)
        If Mystr <> "" Then
          Kill PrmPath & "BomCompare\Testplan_Tmp_Analog\" & Mystr
        End If
        DoEvents
     Loop
  Close #44
  Kill PrmPath & "BomCompare\Kill_Device"
  ' Testplan_Analog.txt
End Sub

Private Sub Kill_File()
On Error Resume Next
a = Dir(PrmPath & "BomCompare\Bom_1\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_2\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_3\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_4\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_5\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_6\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_7\fuckyou")
a = Dir(PrmPath & "BomCompare\Bom_8\fuckyou")
a = Dir(PrmPath & "BomCompare\Testplan_Tmp_Analog\fuckyou")

Kill PrmPath & "BomCompare\WaitDevice.txt"
'Kill PrmPath & "BomCompare\NotDelete.sys"
   Kill PrmPath & "BomCompare\Bom_1\*.*"
   Kill PrmPath & "BomCompare\Bom_2\*.*"
   Kill PrmPath & "BomCompare\Bom_3\*.*"
   Kill PrmPath & "BomCompare\Bom_4\*.*"
   Kill PrmPath & "BomCompare\Bom_5\*.*"
   Kill PrmPath & "BomCompare\Bom_6\*.*"
   Kill PrmPath & "BomCompare\Bom_7\*.*"
   Kill PrmPath & "BomCompare\Bom_8\*.*"
   Kill PrmPath & "BomCompare\Basic_Tmp\*.*"
   Kill PrmPath & "BomCompare\Basic_All_Bom\*.*"
   Kill PrmPath & "BomCompare\Testplan_Tmp_Analog\*.*"
 Kill PrmPath & "BomCompare\Testplan_Analog.txt"
  Kill PrmPath & "BomCompare\Bom_1_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_2_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_3_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_4_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_5_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_6_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_7_Exist.txt"
 Kill PrmPath & "BomCompare\Bom_8_Exist.txt"
 Kill PrmPath & "BomCompare\Comm_Device_Exist.txt"
 Kill PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt"
 RmDir PrmPath & "BomCompare\Testplan_Tmp_Analog"
 
 RmDir PrmPath & "BomCompare\Basic_All_Bom"
RmDir PrmPath & "BomCompare\Basic_Tmp"

RmDir PrmPath & "BomCompare\Bom_1"
RmDir PrmPath & "BomCompare\Bom_2"
RmDir PrmPath & "BomCompare\Bom_3"
RmDir PrmPath & "BomCompare\Bom_4"
RmDir PrmPath & "BomCompare\Bom_5"
RmDir PrmPath & "BomCompare\Bom_6"
RmDir PrmPath & "BomCompare\Bom_7"
RmDir PrmPath & "BomCompare\Bom_8"
 
 
 
End Sub

Private Sub Option1_Click()


If Option1.Value = True Then
   Frame4.Enabled = True
   txtTestplan.Enabled = True
   cmdBom8Ge.Enabled = False
   cmdBomAndBom.Enabled = False
   cmdBom8GeVer.Enabled = False
    txtBom1.Enabled = True
    txtBom2.Enabled = True
    txtBom3.Enabled = True
    txtBom4.Enabled = True
    txtBom5.Enabled = True
    txtBom6.Enabled = True
    txtBom7.Enabled = True
    txtBom8.Enabled = True
End If
End Sub

Private Sub Option2_Click()

If Option2.Value = True Then
   Frame4.Enabled = False
   txtBom3.Enabled = False
   txtBom2.Enabled = True
   txtBom1.Enabled = True
   txtBom4.Enabled = False
   txtBom5.Enabled = False
   txtBom6.Enabled = False
   txtBom7.Enabled = False
   txtBom8.Enabled = False
   txtTestplan.Enabled = False
   cmdBomAndBom.Enabled = True
   cmdBom8Ge.Enabled = False
   cmdBom8GeVer.Enabled = False

 
End If
End Sub

Private Sub Option3_Click()
If Option3.Enabled = True Then

    cmdBom8Ge.Enabled = True
    Frame4.Enabled = False
    txtTestplan.Enabled = False
    cmdBomAndBom.Enabled = False
    cmdBom8GeVer.Enabled = False
    txtBom1.Enabled = True
    txtBom2.Enabled = True
    txtBom3.Enabled = True
    txtBom4.Enabled = True
    txtBom5.Enabled = True
    txtBom6.Enabled = True
    txtBom7.Enabled = True
    txtBom8.Enabled = True
End If
End Sub

Private Sub Option4_Click()
If Option4.Enabled = True Then

    cmdBom8Ge.Enabled = False
    Frame4.Enabled = False
    txtTestplan.Enabled = False
    cmdBomAndBom.Enabled = False
    cmdBom8GeVer.Enabled = True
    txtBom1.Enabled = True
    txtBom2.Enabled = True
    txtBom3.Enabled = True
    txtBom4.Enabled = True
    txtBom5.Enabled = True
    txtBom6.Enabled = True
    txtBom7.Enabled = True
    txtBom8.Enabled = True
End If
End Sub

Private Sub txtBom1_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom1.Text = Me.CommonDialog1.FileName
     bRunBom1 = True
     l1.Caption = Me.CommonDialog1.FileTitle
     
    If Dir(txtBom1.Text) = "" Then
      txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strBom1Path = ""
         l1.Caption = ""
        bRunBom1 = False
      Exit Sub
      Else
        If txtBom1.Text = txtTestplan.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom2.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom3.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom4.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom5.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom6.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom7.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
        If txtBom1.Text = txtBom8.Text Then
            txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom1Path = ""
               l1.Caption = ""
              bRunBom1 = False
        End If
      
    End If
    'strBom1Path
End With




Exit Sub
errh:
      txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
 l1.Caption = ""
        strBom1Path = ""
        bRunBom1 = False
       
        
        
        
        
MsgBox Err.Description, vbCritical
End Sub

Private Sub txtBom2_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom2.Text = Me.CommonDialog1.FileName
     l2.Caption = Me.CommonDialog1.FileTitle
     bRunBom2 = True
    If Dir(txtBom2.Text) = "" Then
      txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
         l2.Caption = ""
        strBom2Path = ""
        bRunBom2 = False
        strBom1Path = ""
         

        
        
      Exit Sub
      Else
        If txtBom2.Text = txtTestplan.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False
        End If
        If txtBom2.Text = txtBom1.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
        If txtBom2.Text = txtBom3.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
        If txtBom2.Text = txtBom4.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
        If txtBom2.Text = txtBom5.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
        If txtBom2.Text = txtBom6.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
              MsgBox "File not find!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
        If txtBom2.Text = txtBom7.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
        If txtBom2.Text = txtBom8.Text Then
            txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l2.Caption = ""
              strBom2Path = ""
              bRunBom2 = False

        End If
      
    End If
    'strBom1Path
End With


Exit Sub
errh:
      txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
 l2.Caption = ""
        strBom2Path = ""
        bRunBom2 = False
         
MsgBox Err.Description, vbCritical
End Sub

Private Sub txtBom3_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom3.Text = Me.CommonDialog1.FileName
     bRunBom3 = True
      l3.Caption = Me.CommonDialog1.FileTitle
    If Dir(txtBom3.Text) = "" Then
      txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
         l3.Caption = ""
        strBom3Path = ""
          
        bRunBom3 = False
        
        
        
      Exit Sub
      Else
        If txtBom3.Text = txtTestplan.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False
        End If
        If txtBom3.Text = txtBom1.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
        If txtBom3.Text = txtBom2.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
        If txtBom3.Text = txtBom4.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
              MsgBox "File not find!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
        If txtBom3.Text = txtBom5.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
        If txtBom3.Text = txtBom6.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
        If txtBom3.Text = txtBom7.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
        If txtBom3.Text = txtBom8.Text Then
            txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l3.Caption = ""
              strBom3Path = ""
              bRunBom3 = False

        End If
      
    End If
    'strBom1Path
End With
Exit Sub
errh:
      txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
 l3.Caption = ""
        strBom3Path = ""
        bRunBom3 = False
        
MsgBox Err.Description, vbCritical

End Sub

Private Sub txtBom4_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom4.Text = Me.CommonDialog1.FileName
     l4.Caption = Me.CommonDialog1.FileTitle
    bRunBom4 = True
    If Dir(txtBom4.Text) = "" Then
      txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
         MsgBox "File not find!", vbCritical
         l4.Caption = ""
        strBom4Path = ""
        bRunBom4 = False
        
        
      Exit Sub
      Else
        If txtBom4.Text = txtTestplan.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False
        End If
        If txtBom4.Text = txtBom1.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False


        End If
        If txtBom4.Text = txtBom3.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False


        End If
        If txtBom4.Text = txtBom2.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False


        End If
        If txtBom4.Text = txtBom5.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False


        End If
        If txtBom4.Text = txtBom6.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False

        End If
        If txtBom4.Text = txtBom7.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False


        End If
        If txtBom4.Text = txtBom8.Text Then
            txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
               l4.Caption = ""
              strBom4Path = ""
              bRunBom4 = False


        End If
      
    End If
    'strBom1Path
End With

Exit Sub
errh:
      txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
 l4.Caption = ""
        strBom4Path = ""
        bRunBom4 = False
MsgBox Err.Description, vbCritical

End Sub


Private Sub txtBom5_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom5.Text = Me.CommonDialog1.FileName
     bRunBom5 = True
     l5.Caption = Me.CommonDialog1.FileTitle
     
    If Dir(txtBom5.Text) = "" Then
      txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strBom5Path = ""
         l5.Caption = ""
        bRunBom5 = False
        
        
      Exit Sub
      Else
        If txtBom5.Text = txtTestplan.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom2.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom3.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom4.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom1.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom6.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom7.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
        If txtBom5.Text = txtBom8.Text Then
            txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom5Path = ""
               l5.Caption = ""
              bRunBom5 = False
        End If
      
    End If
    'strBom1Path
End With

Exit Sub
errh:
      txtBom5.Text = " Please open bom5 file!(DblClick me open file!)"
 l5.Caption = ""
        strBom5Path = ""
        bRunBom5 = False
        
        
        
MsgBox Err.Description, vbCritical

End Sub

Private Sub txtBom6_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom6.Text = Me.CommonDialog1.FileName
     bRunBom6 = True
     l6.Caption = Me.CommonDialog1.FileTitle
     
    If Dir(txtBom6.Text) = "" Then
      txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strBom6Path = ""
         l6.Caption = ""
        bRunBom6 = False
        
        
        
      Exit Sub
      Else
        If txtBom6.Text = txtTestplan.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom2.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom3.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom4.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom1.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom5.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom7.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
        If txtBom6.Text = txtBom8.Text Then
            txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom6Path = ""
               l6.Caption = ""
              bRunBom6 = False
        End If
      
    End If
    'strBom1Path
End With

Exit Sub
errh:
      txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
 l6.Caption = ""
        strBom6Path = ""
        bRunBom6 = False
        
        
        
MsgBox Err.Description, vbCritical



End Sub

Private Sub txtBom7_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom7.Text = Me.CommonDialog1.FileName
     bRunBom7 = True
     l7.Caption = Me.CommonDialog1.FileTitle
     
    If Dir(txtBom7.Text) = "" Then
      txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strBom7Path = ""
         l7.Caption = ""
        bRunBom7 = False
        
        
      Exit Sub
      Else
        If txtBom7.Text = txtTestplan.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom2.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom3.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom4.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom1.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom5.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom6.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
        If txtBom7.Text = txtBom8.Text Then
            txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom7Path = ""
               l7.Caption = ""
              bRunBom7 = False
        End If
      
    End If
    'strBom1Path
End With
Exit Sub
errh:
      txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
 l7.Caption = ""
        strBom7Path = ""
        bRunBom7 = False
          If Option4.Value = True Then
 
        bRunBom8 = False
 
        l8.Caption = ""
        strBom1Path = ""
         
 
        strBom8Path = ""
 
        txtBom8.Text = " Please open bom8 file!(DblClick me open file!)"
         
 
        txtBom8.Enabled = False
      End If
        
        
        
MsgBox Err.Description, vbCritical



End Sub

Private Sub txtBom8_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = ""
    .Filter = "*.txt|*.txt|*.*|*.*"
    .ShowOpen
    txtBom8.Text = Me.CommonDialog1.FileName
     bRunBom8 = True
     l8.Caption = Me.CommonDialog1.FileTitle
     
    If Dir(txtBom8.Text) = "" Then
      txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strBom8Path = ""
         l8.Caption = ""
        bRunBom8 = False
      Exit Sub
      Else
        If txtBom8.Text = txtTestplan.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom2.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom3.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom4.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom1.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom5.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom6.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
               MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
        If txtBom8.Text = txtBom7.Text Then
            txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strBom8Path = ""
               l8.Caption = ""
              bRunBom8 = False
        End If
      
    End If
    'strBom1Path
End With
 
Exit Sub
errh:
      txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
 l8.Caption = ""
        strBom8Path = ""
        bRunBom8 = False
MsgBox Err.Description, vbCritical



End Sub

Private Sub txtTestplan_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = "testplan"
     If Boards = True Then
         .Filter = "testplan vb Create file|*testplan*.vb_Create|*.*|*.*"
       Else
       .Filter = "testplan file|*testplan*.*|*.txt|*.txt|*.*|*.*"
     End If
    .ShowOpen
    txtTestplan.Text = Me.CommonDialog1.FileName
      bRunTestplan = True
    If Dir(txtTestplan.Text) = "" Then
      txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strTestplanPath = ""
        bRunTestplan = False
      Exit Sub
      
      Else
        If txtTestplan.Text = txtBom1.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom2.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom3.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom4.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom5.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom6.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom7.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
        If txtTestplan.Text = txtBom8.Text Then
             txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
              MsgBox "File reload!", vbCritical
              strTestplanPath = ""
              bRunTestplan = False
        End If
      
      
      
      
      
      
    End If
    'strBom1Path
End With
Exit Sub
errh:
      txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
 
        strTestplanPath = ""
        bRunTestplan = False
MsgBox Err.Description, vbCritical

End Sub

Private Sub Create_Sub_analog_file()
 On Error Resume Next
 Dim Mystr As String
 Dim strTmp As String
    Open PrmPath & "BomCompare\sub_analog.txt" For Output As #19
If bRunBom1 = True Then
       Print #19, "!#############" & Replace(l1.Caption, ".txt", "") & " bom1 Start#############"
       Print #19,
    Open PrmPath & "BomCompare\Bom_1_Exist.txt" For Input As #11
       Do Until EOF(11)
          Line Input #11, Mystr
            'Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
              End If
               strTmp = ""
               DoEvents
       Loop
       Print #19,
       Print #19, "!#############" & Replace(l1.Caption, ".txt", "") & " bom1 End#############"
       Print #19,
       Print #19,
    Close #11
End If
If bRunBom2 = True Then
       Print #19, "!#############" & Replace(l2.Caption, ".txt", "") & " bom2 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
    Open PrmPath & "BomCompare\Bom_2_Exist.txt" For Input As #12
        Do Until EOF(12)
          Line Input #12, Mystr
            'Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  
                  Else
                   Print #19, Mystr
                End If
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l2.Caption, ".txt", "") & " bom2 End#############"
       Print #19,
       Print #19,
    Close #12
End If
If bRunBom3 = True Then
       Print #19, "!#############" & Replace(l3.Caption, ".txt", "") & " bom3 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
        
    Open PrmPath & "BomCompare\Bom_3_Exist.txt" For Input As #13
        Do Until EOF(13)
          Line Input #13, Mystr
            'Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l3.Caption, ".txt", "") & " bom3 End#############"
       Print #19,
       Print #19,
    Close #13
End If
If bRunBom4 = True Then
       Print #19, "!#############" & Replace(l4.Caption, ".txt", "") & " bom4 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
 
    Open PrmPath & "BomCompare\Bom_4_Exist.txt" For Input As #14
        Do Until EOF(14)
          Line Input #14, Mystr
            'Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
                 
                 
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l4.Caption, ".txt", "") & " bom4 End#############"
       Print #19,
       Print #19,
    Close #14
 End If
If bRunBom5 = True Then
       Print #19, "!#############" & Replace(l5.Caption, ".txt", "") & " bom5 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
    
    Open PrmPath & "BomCompare\Bom_5_Exist.txt" For Input As #15
        Do Until EOF(15)
          Line Input #15, Mystr
           ' Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
                 
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l5.Caption, ".txt", "") & " bom5 End#############"
       Print #19,
       Print #19,
    Close #15
End If
If bRunBom6 = True Then
       Print #19, "!#############" & Replace(l6.Caption, ".txt", "") & " bom6 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
        
    Open PrmPath & "BomCompare\Bom_6_Exist.txt" For Input As #16
    
        Do Until EOF(16)
          Line Input #16, Mystr
           ' Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
                 
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l6.Caption, ".txt", "") & " bom6 End#############"
       Print #19,
       Print #19,
    Close #16
End If
If bRunBom7 = True Then
        Print #19, "!#############" & Replace(l7.Caption, ".txt", "") & " bom7 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
    
    
    
    Open PrmPath & "BomCompare\Bom_7_Exist.txt" For Input As #17
        Do Until EOF(17)
          Line Input #17, Mystr
            'Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
                 
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l7.Caption, ".txt", "") & " bom7 End#############"
       Print #19,
       Print #19,
    Close #17
End If
If bRunBom8 = True Then
        Print #19, "!#############" & Replace(l8.Caption, ".txt", "") & " bom8 Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
    Open PrmPath & "BomCompare\Bom_8_Exist.txt" For Input As #18
        Do Until EOF(18)
          Line Input #18, Mystr
           ' Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then

                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If

              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############" & Replace(l8.Caption, ".txt", "") & " bom8 End#############"
       Print #19,
       Print #19,
    Close #18
End If
       Print #19, "!#############Comm device Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
         
    Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Input As #20
        Do Until EOF(20)
          Line Input #20, Mystr
           ' Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
              End If
                 strTmp = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############Comm device End#############"
       Print #19,
       Print #19,
       
    Close #20
    
       Print #19, "!#############Testplan=True  All version not find! Start#############"
       Print #19,
       Mystr = ""
       strTmp = ""
             
    Open PrmPath & "BomCompare\NotTest_in_Curr_Ver.txt" For Input As #21
    
         Do Until EOF(21)
          Line Input #21, Mystr
           ' Mystr = LCase(Mystr)
            strTmp = Trim(LCase(Mystr))
              If strTmp <> "" Then
                If Boards = True Then
                    If InStr(Replace(strTmp, " ", ""), "!testcommentedintestorder") = 0 Then
                       Print #19, Mystr & " " & strBoardsNumber
                      Else
                       Print #19, Mystr
                    End If
                  Else
                   Print #19, Mystr
                End If
                 
              End If
                 strTmp = ""
                 Mystr = ""
                 DoEvents
        Loop
       Print #19,
       Print #19, "!#############Testplan=True  All version not find! End#############"
    Close #21

    
 Close #19
 If Boards = True Then
 
    FileCopy PrmPath & "BomCompare\sub_analog.txt", PrmPath & "BomCompare\" & Replace(strBoardsNumber, "*", "x") & "_sub_analog.txt"
 
    Kill PrmPath & "BomCompare\sub_analog.txt"
    Me.Caption = "Save file: " & PrmPath & "BomCompare\" & Replace(strBoardsNumber, "*", "x") & "_sub_analog.txt"
    Exit Sub
    
 End If
 
 Msg3.Caption = "..\BomCompare\sub_analog.txt"
 Me.Caption = "Save file: " & PrmPath & "BomCompare\sub_analog.txt"
End Sub



Private Sub Bom8Comp_Bom1()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom1Path = Trim(txtBom1.Text)
If Dir(strBom1Path) = "" Then
   txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
   bRunBom1 = False
   txtBom1.SetFocus
   strBom1Path = ""
   MsgBox "Bom1 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
   Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom1Path For Input As #50
  '  Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
    
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l1.Caption & "_Bom1.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom1 file..."
               Mystr = LCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  '   strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If
               
  
               
               
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then
 
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                             strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom1_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                              Open PrmPath & "BomCompare\Bom_1\" & strFileNameTmp & strDeviceType For Output As #22
                                   Print #22, strBom1_DeviceName
                              Close #22
                              
                               Open PrmPath & "BomCompare\Bom_1\" & strTmp(0) For Output As #22
                                   Print #22, strBom1_DeviceName & strDeviceType
                              Close #22
                              
                              
                              
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l1.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
'Close #3

 Close #50
 Close #58
 ' intYY = intYY + intFile_Line
  
        Msg1.Caption = l1.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom1 file is null!", vbCritical
          Exit Sub
        End If
  ' intFile_Line = 0
 
End Sub
Private Sub Bom8Comp_Bom2()
 Dim strBom2_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
  '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 
 
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_2"
intFile_Line = 0
intDevice_Ge = 0
strBom2Path = Trim(txtBom2.Text)
If Dir(strBom2Path) = "" Then
   txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
   bRunBom2 = False
   txtBom2.SetFocus
   strBom2Path = ""
   MsgBox "Bom2 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom2 file
   Kill PrmPath & "BomCompare\Bom_2\*.*"

   Open strBom2Path For Input As #50
   
    ' Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l2.Caption & "_Bom2.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom2_DeviceName
               Msg1.Caption = "Reading Bom2 file..."
               Mystr = LCase(Trim(strBom2_DeviceName))
               If Mystr <> "" Then
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
                 '  strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If

               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next
                             
                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                            strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp) & strDeviceType = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom2_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                             
                             
                             Open PrmPath & "BomCompare\Bom_2\" & strTmp(0) For Output As #22
                                   Print #22, strBom2_DeviceName & strDeviceType
                              Close #22
                             
                             
                              Open PrmPath & "BomCompare\Bom_2\" & strFileNameTmp & strDeviceType For Output As #22
                              
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                              
                                   Print #22, strBom2_DeviceName
                              Close #22
                              
                              
                              
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l2.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 'Close #3
' intYY = intYY + intFile_Line
        Msg1.Caption = l2.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom2 file is null!", vbCritical
          Exit Sub
        End If
   
' intFile_Line = 0
End Sub

Private Sub Bom8Comp_Bom3()
 Dim strBom3_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 
 
   '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_3"
intFile_Line = 0
intDevice_Ge = 0
strBom3Path = Trim(txtBom3.Text)
If Dir(strBom3Path) = "" Then
   txtBom3.Text = " Please open Bom3 file!(DblClick me open file!)"
   bRunBom3 = False
   txtBom3.SetFocus
   strBom3Path = ""
   MsgBox "Bom3 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom3 file
   Kill PrmPath & "BomCompare\Bom_3\*.*"

   Open strBom3Path For Input As #50
   
   'Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l3.Caption & "_Bom3.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom3_DeviceName
               Msg1.Caption = "Reading Bom3 file..."
               Mystr = LCase(Trim(strBom3_DeviceName))
               If Mystr <> "" Then
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
                 '  strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If


               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                            strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom3_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                               Open PrmPath & "BomCompare\Bom_3\" & strTmp(0) For Output As #22
                                   Print #22, strBom3_DeviceName & strDeviceType
                              Close #22
                             
                              Open PrmPath & "BomCompare\Bom_3\" & strFileNameTmp & strDeviceType For Output As #22
                              
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                              
                                   Print #22, strBom3_DeviceName
                              Close #22
                              
                              
                              
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                      
                                   
                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l3.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 'Close #3
        Msg1.Caption = l3.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom3 file is null!", vbCritical
          Exit Sub
        End If
'   intYY = inytt + intFile_Line
'   intFile_Line = 0
'
End Sub

Private Sub Bom8Comp_Bom4()
 Dim strBom4_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 
 
 
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_4"
intFile_Line = 0
intDevice_Ge = 0
strBom4Path = Trim(txtBom4.Text)
If Dir(strBom4Path) = "" Then
   txtBom4.Text = " Please open Bom4 file!(DblClick me open file!)"
   bRunBom4 = False
   txtBom4.SetFocus
   strBom4Path = ""
   MsgBox "Bom4 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom4 file
   Kill PrmPath & "BomCompare\Bom_4\*.*"

   Open strBom4Path For Input As #50
   
   ' Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l4.Caption & "_Bom4.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom4_DeviceName
               Msg1.Caption = "Reading Bom4 file..."
               Mystr = LCase(Trim(strBom4_DeviceName))
               If Mystr <> "" Then
                  '  strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If
               
               
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                           strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom4_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                               Open PrmPath & "BomCompare\Bom_4\" & strTmp(0) For Output As #22
                                   Print #22, strBom4_DeviceName & strDeviceType
                              Close #22
                             
                              Open PrmPath & "BomCompare\Bom_4\" & strFileNameTmp & strDeviceType For Output As #22
                              
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                              
                                   Print #22, strBom4_DeviceName
                              Close #22
                              
                              
                              
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
               
                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l4.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
' Close #3
        Msg1.Caption = l4.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom4 file is null!", vbCritical
          Exit Sub
        End If
      intYY = inytt + intFile_Line
      intFile_Line = 0
 
End Sub


Private Sub Bom8Comp_Bom5()
 Dim strBom5_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
     '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 
 
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_5"
intFile_Line = 0
intDevice_Ge = 0
strBom5Path = Trim(txtBom5.Text)
If Dir(strBom5Path) = "" Then
   txtBom5.Text = " Please open Bom5 file!(DblClick me open file!)"
   bRunBom5 = False
   txtBom5.SetFocus
   strBom5Path = ""
   MsgBox "Bom5 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom5 file
   Kill PrmPath & "BomCompare\Bom_5\*.*"

   Open strBom5Path For Input As #50
   
   ' Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l5.Caption & "_Bom5.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom5_DeviceName
               Msg1.Caption = "Reading Bom5 file..."
               Mystr = LCase(Trim(strBom5_DeviceName))
               If Mystr <> "" Then
                    '  strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If
               
               
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                              strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom5_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                                 Open PrmPath & "BomCompare\Bom_5\" & strTmp(0) For Output As #22
                                   Print #22, strBom5_DeviceName & strDeviceType
                              Close #22
                             
                             
                              Open PrmPath & "BomCompare\Bom_5\" & strFileNameTmp & strDeviceType For Output As #22
                              
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                              
                                   Print #22, strBom5_DeviceName
                              Close #22
                              
                              
                              
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l5.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 'Close #3
        Msg1.Caption = l5.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom5 file is null!", vbCritical
          Exit Sub
        End If
       intYY = inytt + intFile_Line
      intFile_Line = 0
 
End Sub

Private Sub Bom8Comp_Bom6()
 Dim strBom6_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 
      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_6"
intFile_Line = 0
intDevice_Ge = 0
strBom6Path = Trim(txtBom6.Text)
If Dir(strBom6Path) = "" Then
   txtBom6.Text = " Please open Bom6 file!(DblClick me open file!)"
   bRunBom6 = False
   txtBom6.SetFocus
   strBom6Path = ""
   MsgBox "Bom6 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom6 file
   Kill PrmPath & "BomCompare\Bom_6\*.*"

   Open strBom6Path For Input As #50
   
    'Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
   
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l6.Caption & "_Bom6.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom6_DeviceName
               Msg1.Caption = "Reading Bom6 file..."
               Mystr = LCase(Trim(strBom6_DeviceName))
              
               If Mystr <> "" Then
                    ' strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If
               
               
               
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then

                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file

'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                          strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom6_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                               Open PrmPath & "BomCompare\Bom_6\" & strTmp(0) For Output As #22
                                   Print #22, strBom6_DeviceName & strDeviceType
                              Close #22
                             
                              Open PrmPath & "BomCompare\Bom_6\" & strFileNameTmp & strDeviceType For Output As #22

                              
                                   Print #22, strBom6_DeviceName
                              Close #22

                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  
                                   
                                   
                                   
                                   
                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l6.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
' Close #3
        Msg1.Caption = l6.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom6 file is null!", vbCritical
          Exit Sub
        End If
   
       intYY = inytt + intFile_Line
      intFile_Line = 0
End Sub

Private Sub Bom8Comp_Bom7()
 Dim strBom7_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
       '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_7"
intFile_Line = 0
intDevice_Ge = 0
strBom7Path = Trim(txtBom7.Text)
If Dir(strBom7Path) = "" Then
   txtBom7.Text = " Please open Bom7 file!(DblClick me open file!)"
   bRunBom7 = False
   txtBom7.SetFocus
   strBom7Path = ""
   MsgBox "Bom7 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom7 file
   Kill PrmPath & "BomCompare\Bom_7\*.*"

   Open strBom7Path For Input As #50
   
   ' Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
   
   Open PrmPath & "BomCompare\BomAndBom_Comp\" & l7.Caption & "_Bom7.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom7_DeviceName
               Msg1.Caption = "Reading Bom7 file..."
               Mystr = LCase(Trim(strBom7_DeviceName))
               If Mystr <> "" Then
                 '  strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If
               
               
               
               
               
               
               
               
               
               
               
               
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then

                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file

'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                             strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom7_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                                Open PrmPath & "BomCompare\Bom_7\" & strTmp(0) For Output As #22
                                   Print #22, strBom7_DeviceName & strDeviceType
                              Close #22
                             
                             
                             
                              Open PrmPath & "BomCompare\Bom_7\" & strFileNameTmp & strDeviceType For Output As #22

                              
                                   Print #22, strBom7_DeviceName
                              Close #22

                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l7.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 'Close #3
        Msg1.Caption = l7.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom7 file is null!", vbCritical
          Exit Sub
        End If
         intYY = inytt + intFile_Line
      intFile_Line = 0
 
End Sub


Private Sub Bom8Comp_Bom8()
 Dim strBom8_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
        '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 Dim strFileNameTmp As String
 '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2
 
 
 On Error Resume Next
 MkDir PrmPath & "BomCompare\Bom_8"
intFile_Line = 0
intDevice_Ge = 0
strBom8Path = Trim(txtBom8.Text)
If Dir(strBom8Path) = "" Then
   txtBom8.Text = " Please open Bom8 file!(DblClick me open file!)"
   bRunBom8 = False
   txtBom8.SetFocus
   strBom8Path = ""
   MsgBox "Bom8 not find!,please check path!", vbCritical
   Exit Sub
End If

'open Bom8 file
   Kill PrmPath & "BomCompare\Bom_8\*.*"

   Open strBom8Path For Input As #50
   
   
' Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Output As #3
Open PrmPath & "BomCompare\BomAndBom_Comp\" & l8.Caption & "_Bom8.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom8_DeviceName
               Msg1.Caption = "Reading Bom8 file..."
               Mystr = LCase(Trim(strBom8_DeviceName))
               
                   
               If Mystr <> "" Then
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@2

                    'strDeviceType = ".UnDevice"
                   If Mystr = "capacitor" Then
                       strDeviceType = ".capacitor"
                   End If
                   If Mystr = "pin library" Then
                       strDeviceType = ".pin_library"
                   End If
                   If Mystr = "jumper" Then
                       strDeviceType = ".jumper"
                   End If
                   If Mystr = "resistor" Then
                       strDeviceType = ".resistor"
                   End If
                   If Mystr = "connector" Then
                       strDeviceType = ".connector"
                   End If
                   If Mystr = "diode" Then
                       strDeviceType = ".diode"
                   End If


               If Left(Mystr, 1) <> "-" And Mystr <> "capacitor" And Mystr <> "pin library" And Mystr <> "jumper" And Mystr <> "resistor" And Mystr <> "diode" And Mystr <> "connector" Then

                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(0)) <> "" Then
                         ' Print #99, UCase(strTmp(UBound(strTmp))) & "," & UCase(strTmp(0)) & "," & l1.Caption & "," & UCase(Mystr)
                             'create bom1 analog file

'                             Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strTmp(0)) For Output As #22
'                             Close #22
                             strFileNameTmp = ""
                             For t = 0 To UBound(strTmp)
                               If strTmp(t) <> "" Then
                                strFileNameTmp = strFileNameTmp & "$" & strTmp(t)
                               End If
                             Next

                             
                             strFileNameTmp = Replace(strFileNameTmp, """", "'")
                             strFileNameTmp = Replace(strFileNameTmp, "\", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "/", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "*", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "?", "^")
                             strFileNameTmp = Replace(strFileNameTmp, "<", "[")
                             strFileNameTmp = Replace(strFileNameTmp, ">", "]")
                           strFileNameTmp = Replace(strFileNameTmp, ".", ",")
                             
                             
                             If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType) = "" Then
                                 Open PrmPath & "BomCompare\Basic_All_Bom\" & strFileNameTmp & strDeviceType For Output As #22
                                      Print #22, strBom8_DeviceName
                                 Close #22
                                 Print #3, strFileNameTmp & strDeviceType
                             End If
                             
                                Open PrmPath & "BomCompare\Bom_8\" & strTmp(0) For Output As #22
                                   Print #22, strBom8_DeviceName & strDeviceType
                              Close #22
                             
                             
                             
                              Open PrmPath & "BomCompare\Bom_8\" & strFileNameTmp & strDeviceType For Output As #22

                              
                                   Print #22, strBom8_DeviceName
                              Close #22

                                   Print #58, strFileNameTmp & strDeviceType
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                                   
                                   
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                         
        
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l8.Caption & " file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 'Close #3
        Msg1.Caption = l8.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the Bom8 file is null!", vbCritical
          Exit Sub
        End If
         intYY = inytt + intFile_Line
      intFile_Line = 0
 
End Sub
Private Sub FileCompStart_8_NoTestplan()
Dim MyStr1 As String
Dim TempShit As String
Dim inBom1 As Boolean
Dim inBom2 As Boolean
Dim inBom3 As Boolean
Dim inBom4 As Boolean
Dim inBom5 As Boolean
Dim inBom6 As Boolean
Dim inBom7 As Boolean
Dim inBom8 As Boolean
Dim bBASIC As Boolean
Dim AllTure As Boolean
Dim strDevice_Type() As String
Dim strBasicDev() As String



Dim i As Integer
i = 0

Open PrmPath & "BomCompare\Capacitor.txt" For Output As #4
Open PrmPath & "BomCompare\Diode.txt" For Output As #7
Open PrmPath & "BomCompare\Jumper.txt" For Output As #9
Open PrmPath & "BomCompare\Connector.txt" For Output As #5
Open PrmPath & "BomCompare\Pin_Library.txt" For Output As #8
Open PrmPath & "BomCompare\Resistor.txt" For Output As #6
 
' Open PrmPath & "BomCompare\Basic.txt" For Output As #61
  
   ' If bRunBom1 = True Then
    '    Open PrmPath & "BomCompare\Bom1.txt" For Output As #57
    '    Close #57
       
     '  Open PrmPath & "BomCompare\Bom1.txt" For Append As #57
     
         ' Print #57, "!$$$$$" & "," & l1.Caption
         
         
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
      '   inBom1 = True
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
          Open PrmPath & "BomCompare\BomAndBom_Comp\Basic_All_Bom.dll" For Input As #58
         ' Open PrmPath & "BomCompare\BomAndBom_Comp\" & l1.Caption & "_Bom1.txt" For Input As #58
            Do Until EOF(58)
             Line Input #58, MyStr1
             MyStr1 = Trim(MyStr1)
If MyStr1 <> "" Then

   strBasicDev = Split(UCase(MyStr1), "$")


      If Dir(PrmPath & "BomCompare\Basic_All_Bom\" & Trim(strBasicDev(1))) <> "" Then
            MyStr1 = ""
      End If
  
  
End If


             
If MyStr1 <> "" Then
                         
                         
                            If bRunBom1 = True Then 'bom1
                                   If Dir(PrmPath & "BomCompare\Bom_1\" & MyStr1) <> "" Then
                                      inBom1 = True
                                      Else
                                       inBom1 = False
                                   End If
                               Else
                                inBom1 = True 'bRunBom1 = True
                            End If 'bom1
                        
                             
                         
                         
                        If bRunBom2 = True Then 'bom2
                               If Dir(PrmPath & "BomCompare\Bom_2\" & MyStr1) <> "" Then
                                  inBom2 = True
                                  Else
                                   inBom2 = False
                               End If

                           Else
                            inBom2 = True 'bRunBom2 = True
                        End If 'bom2
                    
                    If bRunBom3 = True Then 'bom3
                            If Dir(PrmPath & "BomCompare\Bom_3\" & MyStr1) <> "" Then
                              inBom3 = True
                              Else
                              inBom3 = False
                            End If
                       Else
                       inBom3 = True
                    End If 'bom3
                    If bRunBom4 = True Then 'bom4
                       If Dir(PrmPath & "BomCompare\Bom_4\" & MyStr1) <> "" Then
                          inBom4 = True
                          Else
                          inBom4 = False
                       End If
                       Else
                       inBom4 = True
                    End If 'bom4
                    If bRunBom5 = True Then 'bom5
                       If Dir(PrmPath & "BomCompare\Bom_5\" & MyStr1) <> "" Then
                          inBom5 = True
                          Else
                          inBom5 = False
                       End If
                       Else
                       inBom5 = True
                    End If 'bom5
                     If bRunBom6 = True Then 'bom6
                   
                       If Dir(PrmPath & "BomCompare\Bom_6\" & MyStr1) <> "" Then
                          inBom6 = True
                          Else
                          inBom6 = False
                          
                       End If
                       Else
                       inBom6 = True
                    End If 'bom6
                    If bRunBom7 = True Then 'bom7
                   
                       If Dir(PrmPath & "BomCompare\Bom_7\" & MyStr1) <> "" Then
                          inBom7 = True
                          Else
                          inBom7 = False
                          
                       End If
                       Else
                       inBom7 = True
                    End If 'bom5
                   If bRunBom8 = True Then 'bom8
                   
                         If Dir(PrmPath & "BomCompare\Bom_8\" & MyStr1) <> "" Then
                            inBom8 = True
                               Else
                            inBom8 = False
                         End If
                        Else
                       inBom8 = True
                    End If 'bom8

'all =true
                
                  AllTure = inBom1 And inBom2 And inBom3 And inBom4 And inBom5 And inBom6 And inBom7 And inBom8
                  bBASIC = False
'+++++++++++++++++++=+++++++++++++++++++++++++++++++++++++
        If AllTure = False Then
    
                             '   Dim strBasicDev() As String
                        
                        
                                strBasicDev = Split(UCase(MyStr1), "$")
                              
'                                If UCase(Trim(strBasicDev(1))) = "R161" Then
'                                   Debug.Print
'                                End If
                                 
                              
                                If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) = "" Then 'basic

                                           Open PrmPath & "BomCompare\Basic_All_Bom\" & Trim(strBasicDev(1)) For Output As #60
                                               
                                           Close #60




                                           Open PrmPath & "BomCompare\Basic_All_Bom\" & MyStr1 For Input As #60
                                               Line Input #60, TempShit
                                           Close #60

                                           '  Print #61, TempShit ' & " " & strVerName_1 & ";"
                                                  strDevice_Type = Split(LCase(MyStr1), ".")
                                                    If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, TempShit
                                                    End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, TempShit
                                                   
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, TempShit
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, TempShit
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, TempShit
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, TempShit
                                                   End If
                                             
                                             
                                     '  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                                         Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1)) For Output As #80
                                            Print #80, TempShit
                                         Close #80
                                         TempShit = ""
                                      bBASIC = True
                                End If 'basic
                                             
''�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
'                                If AllTure = False And inBom1 = True And bRunBom1 = True Then 'BOM1 TRUE
'                                             Open PrmPath & "BomCompare\Bom_1\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_1 & ";")
'
'                                                    strDevice_Type = Split(LCase(MyStr1), ".")
'                                                '   Print #61, TempShit ' & " " & strVerName_1 & ";"
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'                                                   TempShit = ""
'
'                                End If 'BOM1 TRUE
'
'  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
                                
                               
                                
                                If AllTure = False And inBom1 = False And bRunBom1 = True Then 'BOM1  False
                                               
                                               
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_1\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_1\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_1 & ";"
                                                           TempShit = UCase(TempShit)
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_1 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_1 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_1 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_1 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_1 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_1 & ";"
                                                   End If
                                                   
                                         End If 'JJ

                               End If 'BOM1 False


''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
'                                If AllTure = False And inBom2 = True And bRunBom2 = True Then 'BOM2 TRUE
'                                             Open PrmPath & "BomCompare\Bom_2\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_2 & ";")
'                                                  ' Print #61, TempShit '& " " & strVerName_2 & ";"
'
'                                                   strDevice_Type = Split(LCase(MyStr1), ".")
'
'
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM2 TRUE
'
'  '  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                If AllTure = False And inBom2 = False And bRunBom2 = True Then   'BOM2  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_2\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_2\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_2 & ";"
                                                           TempShit = UCase(TempShit)
                                                           
                                                           
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else

                                                   
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_2 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_2 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_2 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_2 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_2 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_2 & ";"
                                                   End If
                                        End If 'JJ
 

                                End If 'BOM2 False
                                
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
'                                If AllTure = False And inBom3 = True And bRunBom3 = True Then 'BOM3 TRUE
'                                             Open PrmPath & "BomCompare\Bom_3\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_3 & ";")
'                                                '   Print #61, TempShit & " " & strVerName_3 & ";"
'                                                  strDevice_Type = Split(LCase(MyStr1), ".")
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM3 TRUE
'
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                
                                If AllTure = False And inBom3 = False And bRunBom3 = True Then      'BOM3  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_3\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_3\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_3 & ";"
                                                           TempShit = UCase(TempShit)
                                                           
                                                           
                                                           
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                           
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_3 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_3 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_3 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_3 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_3 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_3 & ";"
                                                   End If
                                                   
                                         End If 'JJ

                                End If 'BOM3 False
                        
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
'                                 If AllTure = False And inBom4 = True And bRunBom4 = True Then 'BOM4 TRUE
'                                             Open PrmPath & "BomCompare\Bom_4\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_4 & ";")
'                                                    strDevice_Type = Split(LCase(MyStr1), ".")
'                                                  ' Print #61, TempShit & " " & strVerName_4 & ";"
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'
'
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM4 TRUE
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                If AllTure = False And inBom4 = False And bRunBom4 = True Then      'BOM4  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_4\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_4\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_4 & ";"
                                                           TempShit = UCase(TempShit)
                                                           
                                                           
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                           
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_4 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_4 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_4 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_4 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_4 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_4 & ";"
                                                   End If
                                         End If 'JJ
                                End If 'BOM4 False
                                
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
'                                 If AllTure = False And inBom5 = True And bRunBom5 = True Then 'BOM5 TRUE
'                                             Open PrmPath & "BomCompare\Bom_5\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_5 & ";")
'                                                 '  Print #61, TempShit & " " & strVerName_5 & ";"
'                                                  strDevice_Type = Split(LCase(MyStr1), ".")
'
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'
'
'
'
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM5 TRUE
'
' '  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                If AllTure = False And inBom5 = False And bRunBom5 = True Then      'BOM5  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_5\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_5\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_5 & ";"
                                                           TempShit = UCase(TempShit)
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_5 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_5 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_5 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_5 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_5 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_5 & ";"
                                                   End If
                                                   
                                         End If 'JJ
                                End If 'BOM5 False
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'                                 If AllTure = False And inBom6 = True And bRunBom6 = True Then 'BOM6 TRUE
'                                             Open PrmPath & "BomCompare\Bom_6\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_6 & ";")
'                                                   'Print #61, TempShit & " " & strVerName_6 & ";"
'
'                                                    strDevice_Type = Split(LCase(MyStr1), ".")
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM6 TRUE
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                If AllTure = False And inBom6 = False And bRunBom6 = True Then    'BOM6  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_6\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_6\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_6 & ";"
                                                           TempShit = UCase(TempShit)
                                                           
                                                           
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                           
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_6 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_6 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_6 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_6 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_6 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_6 & ";"
                                                   End If
                                                   
                                      End If 'JJ

                                End If 'BOM6 False
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'
'
'                                 If AllTure = False And inBom7 = True And bRunBom7 = True Then 'BOM7 TRUE
'                                             Open PrmPath & "BomCompare\Bom_7\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_7 & ";")
'                                                '   Print #61, TempShit & " " & strVerName_7 & ";"
'                                                 strDevice_Type = Split(LCase(MyStr1), ".")
'
'                                                    If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM7 TRUE
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                
                                If AllTure = False And inBom7 = False And bRunBom7 = True Then    'BOM7  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_7\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_7\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_7 & ";"
                                                           TempShit = UCase(TempShit)
                                                           
                                                           
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                           
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p   20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_7 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_7 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_7 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_7 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_7 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_7 & ";"
                                                   End If
                                                   
                                       End If 'JJ

                                End If 'BOM7 False
                                
                                
                                
                                
                                
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
'                                 If AllTure = False And inBom8 = True And bRunBom8 = True Then 'BOM8 TRUE
'                                             Open PrmPath & "BomCompare\Bom_8\" & MyStr1 For Input As #60
'                                                   Line Input #60, TempShit
'                                             Close #60
'                                                   TempShit = Replace(TempShit, ";", " " & strVerName_8 & ";")
'                                                 '  Print #61, TempShit & " " & strVerName_8 & ";"
'                                                  strDevice_Type = Split(LCase(MyStr1), ".")
'                                                     If strDevice_Type(1) = "capacitor" Then
'                                                        Print #4, TempShit
'                                                    End If
'                                                   'CONNECT
'                                                   If strDevice_Type(1) = "connector" Then
'                                                        Print #5, TempShit
'
'                                                   End If
'                                                   'RES
'                                                   If strDevice_Type(1) = "resistor" Then
'                                                        Print #6, TempShit
'                                                   End If
'                                                   'Diode
'                                                    If strDevice_Type(1) = "diode" Then
'                                                        Print #7, TempShit
'                                                   End If
'                                                   'pin lib
'                                                   If strDevice_Type(1) = "pin_library" Then
'                                                        Print #8, TempShit
'                                                   End If
'                                                   'jumper
'                                                   If strDevice_Type(1) = "jumper" Then
'                                                        Print #9, TempShit
'                                                   End If
'
'
'                                                   TempShit = ""
'
'                                End If 'BOM8 TRUE
'
''  '�ѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡѡ�
                                
                                
                                If AllTure = False And inBom8 = False And bRunBom8 = True Then    'BOM8  False
                                          If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) <> "" And Dir(PrmPath & "BomCompare\Bom_8\" & Trim(strBasicDev(1))) <> "" Then   'JJ
                                                     Open PrmPath & "BomCompare\Bom_8\" & Trim(strBasicDev(1)) For Input As #60
                                                           Line Input #60, TempShit
                                                     Close #60
                                                          strDevice_Type = Split(LCase(TempShit), ";")
                                                           TempShit = strDevice_Type(0) & " " & strVerName_8 & ";"
                                                           TempShit = UCase(TempShit)
                                                           
                                                           
                                                           strDevice_Type(1) = Replace(strDevice_Type(1), ".", "")
                                                           
                                                        If strDevice_Type(1) = "capacitor" Then
                                                            Print #4, TempShit
                                                        End If
                                                       'CONNECT
                                                       If strDevice_Type(1) = "connector" Then
                                                            Print #5, TempShit
    
                                                       End If
                                                       'RES
                                                       If strDevice_Type(1) = "resistor" Then
                                                            Print #6, TempShit
                                                       End If
                                                       'Diode
                                                        If strDevice_Type(1) = "diode" Then
                                                            Print #7, TempShit
                                                       End If
                                                       'pin lib
                                                       If strDevice_Type(1) = "pin_library" Then
                                                            Print #8, TempShit
                                                       End If
                                                       'jumper
                                                       If strDevice_Type(1) = "jumper" Then
                                                            Print #9, TempShit
                                                       End If
    
                                                       TempShit = ""
                                                           
                                                           
                                                   Else
                                                   
                                                   'CAP
                                                   If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, Trim(strBasicDev(1)); Tab(30); "11.1p    20    20 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_8 & ";"
                                                   End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_8 & ";"
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, Trim(strBasicDev(1)); Tab(30); "8.88M    10    10 f NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_8 & ";"
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, Trim(strBasicDev(1)); Tab(30); "0.8    0.2  NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_8 & ";"
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, Trim(strBasicDev(1)); Tab(30); "NP PN " & """" & Trim(strBasicDev(1)) & """" & strVerName_8 & ";"
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, Trim(strBasicDev(1)); Tab(30); "OPEN    " & strVerName_8 & ";"
                                                   End If
                                                        
                                        End If 'JJ
                                End If 'BOM8 False
                        
                        

                        
                       Else
                              strBasicDev = Split(UCase(MyStr1), "$")
                                
                                If UCase(Trim(strBasicDev(1))) = "TC8" Then
                                   Debug.Print
                                End If
                              
                              
                              
                              
                                If Dir(PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1))) = "" Then 'basic
                                

                                           Open PrmPath & "BomCompare\Basic_All_Bom\" & MyStr1 For Input As #60
                                               Line Input #60, TempShit
                                           Close #60
                                             
                                              Open PrmPath & "BomCompare\Basic_All_Bom\" & Trim(strBasicDev(1)) For Output As #60
                                               
                                           Close #60
                                             
                                             
                                             
                                              strDevice_Type = Split(LCase(MyStr1), ".")
                                             
                                             
                                            ' Print #61, TempShit ' & " " & strVerName_1 & ";"
                                                      If strDevice_Type(1) = "capacitor" Then
                                                        Print #4, TempShit
                                                    End If
                                                   'CONNECT
                                                   If strDevice_Type(1) = "connector" Then
                                                        Print #5, TempShit
                                                   
                                                   End If
                                                   'RES
                                                   If strDevice_Type(1) = "resistor" Then
                                                        Print #6, TempShit
                                                   End If
                                                   'Diode
                                                    If strDevice_Type(1) = "diode" Then
                                                        Print #7, TempShit
                                                   End If
                                                   'pin lib
                                                   If strDevice_Type(1) = "pin_library" Then
                                                        Print #8, TempShit
                                                   End If
                                                   'jumper
                                                   If strDevice_Type(1) = "jumper" Then
                                                        Print #9, TempShit
                                                   End If
                                             
                                             
                                             
                                             
                                             
                                             
                                     '  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                                         Open PrmPath & "BomCompare\Basic_Tmp\" & Trim(strBasicDev(1)) For Output As #80
                                            Print #80, TempShit
                                         Close #80
                                         TempShit = ""
                                      '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  
                                       bBASIC = True
                                      
                                      
                                      
                                End If 'basic




                    End If
                        
'+++++++++++++++++++=+++++++++++++++++++++++++++++++++++++
                        
                        
                        
                        
                        
'++++++++++++++++++++++++++++++++++++++++
                        
'                         If bRunBom1 = True Then
'                            Kill PrmPath & "BomCompare\Bom_1\" & MyStr1
'                         End If
'                         If bRunBom2 = True Then
'                            Kill PrmPath & "BomCompare\Bom_2\" & MyStr1
'                         End If
'                         If bRunBom3 = True Then
'                              Kill PrmPath & "BomCompare\Bom_3\" & MyStr1
'                         End If
'                         If bRunBom4 = True Then
'                             Kill PrmPath & "BomCompare\Bom_4\" & MyStr1
'                         End If
'                         If bRunBom5 = True Then
'                             Kill PrmPath & "BomCompare\Bom_5\" & MyStr1
'                         End If
'                         If bRunBom6 = True Then
'                            Kill PrmPath & "BomCompare\Bom_6\" & MyStr1
'                         End If
'                         If bRunBom7 = True Then
'                            Kill PrmPath & "BomCompare\Bom_7\" & MyStr1
'                         End If
'                         If bRunBom8 = True Then
'                            Kill PrmPath & "BomCompare\Bom_8\" & MyStr1
'                         End If
                    

                         
                         
End If

              i = i + 1
              Msg3.Caption = "Runing... " & i
              DoEvents
            Loop
               TempShit = ""
             AllTure = False
             bBASIC = False
             MyStr1 = ""
             inBom1 = False
             inBom2 = False
             inBom3 = False
             inBom4 = False
             inBom5 = False
             inBom6 = False
             inBom7 = False
             inBom8 = False
        ' Close #57
         Close #58
         Close #4
         Close #5
         Close #6
         Close #7
         Close #8
         Close #9

       '  Close #61
  'End If '1
  intYY = i
  
  Open PrmPath & "BomCompare\Board_To_Version_Board_Check_Error.txt" For Output As #3
   Print #3, "All Device is:" & intYY
Close #3
  
  
  Exit Sub
  

End Sub
Private Sub FileHeBin()
Dim MyStr1 As String
Open PrmPath & "BomCompare\board.txt" For Output As #57

      Print #57, "CAPACITOR"
         If Dir(PrmPath & "BomCompare\Capacitor.txt") <> "" Then
             Open PrmPath & "BomCompare\Capacitor.txt" For Input As #58
                 Do Until EOF(58)
                 Line Input #58, MyStr1
                   If MyStr1 <> "" Then
                     Print #57, MyStr1
                   End If
                 Loop
                 DoEvents
                 
             Close #58
                Kill PrmPath & "BomCompare\Capacitor.txt"
                MyStr1 = ""
            Else
              MsgBox "Are you delete file " & PrmPath & "BomCompare\Capacitor.txt" & "?", vbCritical
          End If
      
      Print #57,

      Print #57, "RESISTOR"
         If Dir(PrmPath & "BomCompare\Resistor.txt") <> "" Then
             Open PrmPath & "BomCompare\Resistor.txt" For Input As #58
                 Do Until EOF(58)
                 Line Input #58, MyStr1
                   If MyStr1 <> "" Then
                     Print #57, MyStr1
                   End If
                 Loop
                 DoEvents
                 
             Close #58
                Kill PrmPath & "BomCompare\Resistor.txt"
                MyStr1 = ""
            Else
              MsgBox "Are you delete file " & PrmPath & "BomCompare\Resistor.txt" & "?", vbCritical
          End If
      
      
      Print #57,

      Print #57, "DIODE"
         If Dir(PrmPath & "BomCompare\Diode.txt") <> "" Then
             Open PrmPath & "BomCompare\Diode.txt" For Input As #58
                 Do Until EOF(58)
                 Line Input #58, MyStr1
                   If MyStr1 <> "" Then
                     Print #57, MyStr1
                   End If
                 Loop
                 DoEvents
                 
             Close #58
                Kill PrmPath & "BomCompare\Diode.txt"
                MyStr1 = ""
            Else
              MsgBox "Are you delete file " & PrmPath & "BomCompare\Diode.txt" & "?", vbCritical
          End If
      
      Print #57,

      Print #57, "JUMPER"
         If Dir(PrmPath & "BomCompare\Jumper.txt") <> "" Then
             Open PrmPath & "BomCompare\Jumper.txt" For Input As #58
                 Do Until EOF(58)
                 Line Input #58, MyStr1
                   If MyStr1 <> "" Then
                     Print #57, MyStr1
                   End If
                 Loop
                 DoEvents
                 
             Close #58
                Kill PrmPath & "BomCompare\Jumper.txt"
                MyStr1 = ""
            Else
              MsgBox "Are you delete file " & PrmPath & "BomCompare\Jumper.txt" & "?", vbCritical
          End If
      
      Print #57,
    

      Print #57, "CONNECTOR"
         If Dir(PrmPath & "BomCompare\Connector.txt") <> "" Then
             Open PrmPath & "BomCompare\Connector.txt" For Input As #58
                 Do Until EOF(58)
                 Line Input #58, MyStr1
                   If MyStr1 <> "" Then
                     Print #57, MyStr1
                   End If
                 Loop
                 DoEvents
                 
             Close #58
                Kill PrmPath & "BomCompare\Connector.txt"
                MyStr1 = ""
            Else
              MsgBox "Are you delete file " & PrmPath & "BomCompare\Connector.txt" & "?", vbCritical
          End If
      
      Print #57,
    

      Print #57, "PIN LIBRARY"
         If Dir(PrmPath & "BomCompare\Pin_Library.txt") <> "" Then
             Open PrmPath & "BomCompare\Pin_Library.txt" For Input As #58
                 Do Until EOF(58)
                 Line Input #58, MyStr1
                   If MyStr1 <> "" Then
                     Print #57, MyStr1
                   End If
                 Loop
                 DoEvents
                 
             Close #58
                Kill PrmPath & "BomCompare\Pin_Library.txt"
                MyStr1 = ""
            Else
              MsgBox "Are you delete file " & PrmPath & "BomCompare\Pin_Library.txt" & "?", vbCritical
          End If
      
    ' Print #57,
    
Close #57
End Sub
Private Sub ReadBom1_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
  Dim strTmp1() As String
   Dim strTmp2() As String
 Dim strBoardver() As String
 
 On Error Resume Next
 strBoardver = Split(l1.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom1Path = Trim(txtBom1.Text)
If Dir(strBom1Path) = "" Then
   txtBom1.Text = " Please open bom1 file!(DblClick me open file!)"
   bRunBom1 = False
   txtBom1.SetFocus
   strBom1Path = ""
   MsgBox "Bom1 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  Kill PrmPath & "BomCompare\Bom_1\*.*"
  'Kill PrmPath & "BomCompare\Bom_2\*.*"

   Open strBom1Path For Input As #50
   Open PrmPath & "BomCompare\Basic.txt" For Output As #59
   Open PrmPath & "BomCompare\Basic_Space_String.txt" For Output As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Output As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom1 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                         '       strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                              
                              
'
'                               Open PrmPath & "BomCompare\Bom_2\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #23
'                                   tmptmpStr = Replace(Mystr, strTmp(0), "")
'                                   tmptmpStr = Trim(Replace(tmptmpStr, strTmp(UBound(strTmp)), ""))
'                                   Print #23, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardVer(0)
'                               Close #23
                              Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
                                   tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1
                                   'tmptmpStr = Trim(Replace(tmptmpStr, strTmp(UBound(strTmp)), ""))
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                   
                                  ' Print #58, strTmp(UBound(strTmp))
                               Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                               Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                               Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                           tmptmpStr = ""
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l1.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #59
 Close #58
 Close #61
        Msg1.Caption = l1.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom1 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub

Private Sub ReadBom2_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 Dim strBoardver() As String
 Dim strTmp1() As String
 On Error Resume Next
 strBoardver = Split(l2.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom2Path = Trim(txtBom2.Text)
If Dir(strBom2Path) = "" Then
   txtBom2.Text = " Please open bom2 file!(DblClick me open file!)"
   bRunBom2 = False
   txtBom2.SetFocus
   strBom2Path = ""
   MsgBox "Bom2 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom2Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
    Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom2 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                              '  strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                   tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1
                              
                              
                              
                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                                   
                                   Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                               Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName

                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                  
                                tmptmpStr = ""
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l2.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
 Close #61
        Msg1.Caption = l2.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom2 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub


Private Sub ReadBom3_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 Dim strBoardver() As String
  Dim strTmp1() As String
 On Error Resume Next
 strBoardver = Split(l3.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom3Path = Trim(txtBom3.Text)
If Dir(strBom3Path) = "" Then
   txtBom3.Text = " Please open bom3 file!(DblClick me open file!)"
   bRunBom3 = False
   txtBom3.SetFocus
   strBom3Path = ""
   MsgBox "Bom3 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom3Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
       Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom3 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                 strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                            '    strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                    tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1
                             
                              
                              
                              
                              
                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   
                                   Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                               Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                              
                              
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
 
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                 tmptmpStr = ""
                                  
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l3.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
 Close #61
        Msg1.Caption = l3.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom3 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub

Private Sub ReadBom4_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
  Dim strTmp1() As String
 Dim strBoardver() As String
 
 On Error Resume Next
 strBoardver = Split(l4.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom4Path = Trim(txtBom4.Text)
If Dir(strBom4Path) = "" Then
   txtBom4.Text = " Please open bom4 file!(DblClick me open file!)"
   bRunBom4 = False
   txtBom4.SetFocus
   strBom4Path = ""
   MsgBox "Bom4 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom4Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
    Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom4 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                              '  strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                    tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1

                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                                   
                                   Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                               Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                                   
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
 
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                               tmptmpStr = ""
                                  
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l4.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
 Close #61
        Msg1.Caption = l4.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom4 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub
Private Sub ReadBom5_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
  Dim strTmp1() As String
 Dim strBoardver() As String
 
 On Error Resume Next
 strBoardver = Split(l5.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom5Path = Trim(txtBom5.Text)
If Dir(strBom5Path) = "" Then
   txtBom5.Text = " Please open bom5 file!(DblClick me open file!)"
   bRunBom5 = False
   txtBom5.SetFocus
   strBom5Path = ""
   MsgBox "Bom5 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom5Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
    Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom5 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                              '  strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                    tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1
                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                                   
                                   Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                               Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
 
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                 tmptmpStr = ""
                                  
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l5.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
 Close #61
        Msg1.Caption = l5.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom5 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub
Private Sub ReadBom6_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
  Dim strTmp1() As String
 Dim strBoardver() As String
 
 On Error Resume Next
 strBoardver = Split(l6.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom6Path = Trim(txtBom6.Text)
If Dir(strBom6Path) = "" Then
   txtBom6.Text = " Please open bom6 file!(DblClick me open file!)"
   bRunBom6 = False
   txtBom6.SetFocus
   strBom6Path = ""
   MsgBox "Bom6 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom6Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
    Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom1 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                              '  strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                    tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1

                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                                   Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                                   Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                                   
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
 
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                  
                               mptmpStr = ""
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l6.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
  Close #61
        Msg1.Caption = l6.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom6 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub
Private Sub ReadBom7_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
 Dim strTmp1() As String
 Dim strBoardver() As String
 
 On Error Resume Next
 strBoardver = Split(l7.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom7Path = Trim(txtBom7.Text)
If Dir(strBom7Path) = "" Then
   txtBom7.Text = " Please open bom7 file!(DblClick me open file!)"
   bRunBom7 = False
   txtBom7.SetFocus
   strBom7Path = ""
   MsgBox "Bom7 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom7Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
   Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom1 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                              '  strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                    tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1

                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                                   
                                  Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                                   Print #61, strTmp(0) & Chr(9) & Replace(tmptmpStr, ",", Chr(9)) & Chr(9) & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
 
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                  
                                mptmpStr = ""
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l7.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
 Close #61
        Msg1.Caption = l7.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom7 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub

Private Sub ReadBom8_Ver_Out_Dir()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim Mystr As String
 Dim strTmp() As String
  Dim strTmp1() As String
 Dim strBoardver() As String
 
 On Error Resume Next
 strBoardver = Split(l8.Caption, ".")
 MkDir PrmPath & "BomCompare\Bom_1"
intFile_Line = 0
intDevice_Ge = 0
strBom8Path = Trim(txtBom8.Text)
If Dir(strBom8Path) = "" Then
   txtBom8.Text = " Please open bom8 file!(DblClick me open file!)"
   bRunBom8 = False
   txtBom8.SetFocus
   strBom8Path = ""
   MsgBox "Bom8 not find!,please check path!", vbCritical
   Exit Sub
End If

'open bom1 file
  'Kill PrmPath & "BomCompare\Bom_1\*.*"

   Open strBom8Path For Input As #50
    Open PrmPath & "BomCompare\Basic.txt" For Append As #59
 Open PrmPath & "BomCompare\Basic_Space_String.txt" For Append As #61
   Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Append As #58
           Do Until EOF(50)
             Line Input #50, strBom1_DeviceName
               Msg1.Caption = "Reading bom1 file..."
               Mystr = UCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "/", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "\", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "*", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "|", "$$$$$$$$$$")
                                strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), "?", "$$$$$$$$$$")
                              '  strTmp(UBound(strTmp)) = Replace(strTmp(UBound(strTmp)), ".", "$$$$$$$$$$")
                             
                             
                              strTmp(UBound(strTmp)) = Trim(strTmp(UBound(strTmp)))
                                    tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1) - 1
                                       
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                       
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   strTmp1 = Split(tmptmpStr, " ")
                                   tmptmpStr = ""
                                   For y = 0 To UBound(strTmp1)
                                     If y = UBound(strTmp1) Then
                                         tmptmpStr = tmptmpStr & "," & strTmp1(UBound(strTmp1))
                                       Else
                                       tmptmpStr = tmptmpStr & " " & strTmp1(y)
                                     End If
                                   Next
                                   tmptmpStr = Trim(tmptmpStr)
                                   Erase strTmp1

                              If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp))) = "" Then
                                   Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Output As #22
                                   Print #58, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                                   Print #61, strTmp(0) & "    " & Replace(tmptmpStr, ",", " ") & "  " & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                                   
                                   Print #59, strTmp(0) & "," & tmptmpStr & "," & strTmp(UBound(strTmp)) ' & "," & strBoardver(0)
                                   
                              Else
                              
                                  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) For Append As #22
                              End If
                            '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
                                   'Print #22, strBom1_DeviceName
                                   tmptmpStr = Replace(Mystr, strTmp(0), "")
                                   tmptmpStr = Trim(Replace(tmptmpStr, strTmp(UBound(strTmp)), ""))
                                   Print #22, strTmp(UBound(strTmp)) & "," & strTmp(0) & "," & tmptmpStr & "," & strBoardver(0)
                              Close #22
                                  'Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0)
                                   
                                  
                                  mptmpStr = ""
                             intDevice_Ge = intDevice_Ge + 1
                             Msg3.Caption = Trim(strTmp(UBound(strTmp)))
                             Msg4.Caption = "Device:" & intDevice_Ge
                           '  Debug.Print strTmp(UBound(strTmp))
                         
                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1
               Msg2.Caption = "Read " & l8.Caption & "file line: " & intFile_Line
               DoEvents
               
              
           Loop
 Close #50
 Close #58
 Close #59
 Close #61
        Msg1.Caption = l8.Caption & " file closed!"
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom8 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub

Private Sub Start_8Ge_Ver_Bom_Comp()
Dim Mystr As String
Dim strMyStr As String

Dim intI As Integer
Dim intJ As Integer
Dim bJumper As Boolean
Dim bPinLib As Boolean
Dim bConnect As Boolean
Dim bDiode As Boolean
Dim bCap As Boolean
Dim bRes As Boolean
Dim strBoardPN As String
Dim bBom1Find As Boolean
Dim bBom2Find As Boolean
Dim bBom3Find As Boolean
Dim bBom4Find As Boolean
Dim bBom5Find As Boolean
Dim bBom6Find As Boolean
Dim bBom7Find As Boolean
Dim bBom8Find As Boolean




intJ = 0
intI = 0
'Open PrmPath & "BomCompare\8GeBom_Jumper.txt" For Output As #74
'Open PrmPath & "BomCompare\8GeBom_Diode.txt" For Output As #75
'Open PrmPath & "BomCompare\8GeBom_Connect.txt" For Output As #76
'Open PrmPath & "BomCompare\8GeBom_PinLib.txt" For Output As #74

Open PrmPath & "BomCompare\8GeBom_Basic_Device.txt" For Output As #73

Open PrmPath & "BomCompare\8GeVer_Bom.txt" For Output As #72


Open PrmPath & "BomCompare\Bom8GeVer_Comp.txt" For Input As #58
    Do Until EOF(58)
      Line Input #58, Mystr

  
      
      Mystr = Trim(Mystr)
      If Trim(Mystr) <> "" Then
          If Mystr = "Q20.1" Then
             Debug.Print
          End If
           Mystr = Replace(Mystr, "/", "$$$$$$$$$$")
           Mystr = Replace(Mystr, "\", "$$$$$$$$$$")
           Mystr = Replace(Mystr, "*", "$$$$$$$$$$")
           Mystr = Replace(Mystr, "|", "$$$$$$$$$$")
           Mystr = Replace(Mystr, "?", "$$$$$$$$$$")
         '  Mystr = Replace(Mystr, ".", "$$$$$$$$$$")
           Open PrmPath & "BomCompare\Bom_1\" & Mystr For Input As #71
             Dim strLineStr_0() As String
             Dim strBomVerstr(10) As String
             Dim strLineStr() As String
                   For t = 0 To 10
                     strBomVerstr(t) = ""
                   Next
               
                Do Until EOF(71)
                  Line Input #71, strMyStr
                  

                  
                  
                   strMyStr = UCase(Trim(strMyStr))
                 If strMyStr <> "" Then
                 intJ = intJ + 1
                    strBomVerstr(intJ) = strMyStr
                    
                 End If
                 If intJ >= 8 Then Exit Do
                Loop
                
bBom1Find = False
bBom2Find = False
bBom3Find = False
bBom4Find = False
bBom5Find = False
bBom6Find = False
bBom7Find = False
bBom8Find = False
                
                
                
                Print #73, strBomVerstr(1)
                
                        If strBomVerstr(1) <> "" Then
                           strLineStr_0 = Split(strBomVerstr(1), ",")
                        End If
                        Print #72, strLineStr_0(0) & "," & strLineStr_0(1) & "," & strLineStr_0(2)
                                 strTmpCaption = Replace(Trim(UCase(l1.Caption)), ".TXT", "")
                               If bRunBom1 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom1Find = True
                               End If
                                 strTmpCaption = Replace(Trim(UCase(l2.Caption)), ".TXT", "")
                                If bRunBom2 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom2Find = True
                               End If
                               
                                  strTmpCaption = Replace(Trim(UCase(l3.Caption)), ".TXT", "")
                               If bRunBom3 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom3Find = True
                               End If
                                 strTmpCaption = Replace(Trim(UCase(l4.Caption)), ".TXT", "")
                                If bRunBom4 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom4Find = True
                               End If
                               
                               
                                 strTmpCaption = Replace(Trim(UCase(l5.Caption)), ".TXT", "")
                               If bRunBom5 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom5Find = True
                               End If
                                 strTmpCaption = Replace(Trim(UCase(l6.Caption)), ".TXT", "")
                                If UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom6Find = True
                               End If
                               
                                  strTmpCaption = Replace(Trim(UCase(l7.Caption)), ".TXT", "")
                               If bRunBom7 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom7Find = True
                               End If
                                 strTmpCaption = Replace(Trim(UCase(l8.Caption)), ".TXT", "")
                                If bRunBom8 = True And UCase(Trim(strLineStr_0(3))) = strTmpCaption Then
                                  bBom8Find = True
                               End If
                   
                   
                   
                   
                       If intBomGeShu = intJ Then
                                 For g = 1 To intJ
                               '  If strBomVerstr(g) <> "" Then
                                    strLineStr_0 = Split(strBomVerstr(g), ",")
                                      strLineStr = Split(strBomVerstr(g + 1), ",")
                    
                                      If strLineStr(1) <> strLineStr_0(1) And strLineStr(0) = strLineStr_0(0) Then
        
                                             Print #72, strBomVerstr(g)
                                      End If
                                                      strTmpCaption = Replace(Trim(UCase(l1.Caption)), ".TXT", "")
                                                    If bRunBom1 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom1Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l2.Caption)), ".TXT", "")
                                                     If bRunBom2 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom2Find = True
                                                    End If
                                                    
                                                       strTmpCaption = Replace(Trim(UCase(l3.Caption)), ".TXT", "")
                                                    If bRunBom3 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom3Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l4.Caption)), ".TXT", "")
                                                     If bRunBom4 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom4Find = True
                                                    End If
                                                    
                                                    
                                                      strTmpCaption = Replace(Trim(UCase(l5.Caption)), ".TXT", "")
                                                    If bRunBom5 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom5Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l6.Caption)), ".TXT", "")
                                                     If bRunBom6 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom6Find = True
                                                    End If
                                                    
                                                       strTmpCaption = Replace(Trim(UCase(l7.Caption)), ".TXT", "")
                                                    If bRunBom7 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom7Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l8.Caption)), ".TXT", "")
                                                     If bRunBom8 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom8Find = True
                                                    End If
                                      
                                      
                                      
                                      
                                      
                                      
                                      
                                '  End If
                              Next
                            Else
                                For g = 2 To intBomGeShu
                               
                                  If strBomVerstr(g) = "" Then
                                             'strLineStr = Split(strBomVerstr(g), ",")
                                            ' If strLineStr(1) <> strLineStr_0(1) And strLineStr(0) = strLineStr_0(0) Then
                                              '  Print #72, strBomVerstr(g)
                                                 strTmpCaption = Replace(Trim(UCase(l1.Caption)), ".TXT", "")
                                                     
                                                If bBom1Find = False And bRunBom1 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""
                                                
                                                 strTmpCaption = Replace(Trim(UCase(l2.Caption)), ".TXT", "")
                                                     
                                                If bBom2Find = False And bRunBom2 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""

                                                 strTmpCaption = Replace(Trim(UCase(l3.Caption)), ".TXT", "")
                                                     
                                                If bBom3Find = False And bRunBom3 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""
 

                                                 strTmpCaption = Replace(Trim(UCase(l4.Caption)), ".TXT", "")
                                                     
                                                If bBom4Find = False And bRunBom4 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""

                                                 strTmpCaption = Replace(Trim(UCase(l5.Caption)), ".TXT", "")
                                                     
                                                If bBom5Find = False And bRunBom5 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""
                                                
                                                 strTmpCaption = Replace(Trim(UCase(l6.Caption)), ".TXT", "")
                                                     
                                                If bBom6Find = False And bRunBom6 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""

                                                 strTmpCaption = Replace(Trim(UCase(l7.Caption)), ".TXT", "")
                                                     
                                                If bBom7Find = False And bRunBom7 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""
 

                                                 strTmpCaption = Replace(Trim(UCase(l8.Caption)), ".TXT", "")
                                                     
                                                If bBom8Find = False And bRunBom8 = True Then
                                                       Print #72, strLineStr_0(0) & ",Not Find in:,, " & strTmpCaption
                                                End If
                                                strTmpCaption = ""

                                            'End If
                                            
                                            Exit For
                                        Else
                                           strLineStr = Split(strBomVerstr(g), ",")
                                          If strLineStr(1) <> strLineStr_0(1) And strLineStr(0) = strLineStr_0(0) Then
                                              Print #72, strBomVerstr(g)
                                                      strTmpCaption = Replace(Trim(UCase(l1.Caption)), ".TXT", "")
                                                    If bRunBom1 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom1Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l2.Caption)), ".TXT", "")
                                                     If bRunBom2 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom2Find = True
                                                    End If
                                                    
                                                       strTmpCaption = Replace(Trim(UCase(l3.Caption)), ".TXT", "")
                                                    If bRunBom3 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom3Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l4.Caption)), ".TXT", "")
                                                     If bRunBom4 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom4Find = True
                                                    End If
                                                    
                                                    
                                                      strTmpCaption = Replace(Trim(UCase(l5.Caption)), ".TXT", "")
                                                    If bRunBom5 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom5Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l6.Caption)), ".TXT", "")
                                                     If bRunBom6 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom6Find = True
                                                    End If
                                                    
                                                       strTmpCaption = Replace(Trim(UCase(l7.Caption)), ".TXT", "")
                                                    If bRunBom7 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom7Find = True
                                                    End If
                                                      strTmpCaption = Replace(Trim(UCase(l8.Caption)), ".TXT", "")
                                                     If bRunBom8 = True And UCase(Trim(strLineStr(3))) = strTmpCaption Then
                                                       bBom8Find = True
                                                    End If
                                                                   
                                              
                                              
                                              
                                              
                                              
                                          End If
                                             
                                    End If
                                     
                                Next
                         
                         
                         
                        End If
                         
                 
            DoEvents
           Mystr = ""
           intJ = 0
           strMyStr = ""
                
           Close #71

      End If

    Loop

  Close #72
  Close #73
Close #58
Kill PrmPath & "BomCompare\Bom8GeVer_Comp.txt"
End Sub
Private Sub Read_BomFile()
Dim Mystr As String
Dim intI As String
Dim strDeviceName As String
Dim TmpStr() As String
Dim tmpSTR1 As String
Dim DeviceType_ As String
Dim DeviceType_A As String
Dim DeviceType_1 As String
Dim CValue As String
Dim RValue As String
Dim strCAP() As String
Dim strRES() As String
Dim strReadText As String
Dim LowToJumper
Dim strDeviceNomber As String
Dim bListPinLib As Boolean
strCH = txtCH.Text
strCL = txtCL.Text
strRH = txtRH.Text
strRL = txtRH.Text
strDH = txtDH.Text
strDL = txtDL.Text


If CheckC.Value = 1 Then
   bListCatacitor = True
   Else
   bListCatacitor = False
End If

If Checklb.Value = 1 Then
   bListPinLib = True
   Else
   bListPinLib = False
End If

 
If CheckR.Value = 1 Then
   bListResistor = True
   Else
   bListResistor = False
End If
If CheckD.Value = 1 Then
   bListDiode = True
   Else
   bListDiode = False
End If

 


On Error GoTo EX
  ' Open PrmPath & "ReadBomValue\WaitCheck.txt" For Output As #7
  '   Print #7, Now
   '   Print #7,
   Open PrmPath & "ReadBomValue\Jumper.txt" For Output As #6
    ' Print #6, Now
      Print #6,
   If bListCatacitor = True Then
      Open PrmPath & "ReadBomValue\Catacitor.txt" For Output As #2
      'Print #2, Now
      Print #2,
   End If
   If bListResistor = True Then
      Open PrmPath & "ReadBomValue\Resistor.txt" For Output As #4
     ' Print #4, Now
      Print #4,
   End If
   If bListDiode = True Then
      Open PrmPath & "ReadBomValue\Diode.txt" For Output As #8
    '  Print #8, Now
      Print #8,
   
   End If
   
 If bListPinLib = True Then
      Open PrmPath & "ReadBomValue\Pin Linrary.txt" For Output As #9
   '   Print #9, Now
      Print #9,
   
   End If
   
   
   
      Open PrmPath & "ReadBomValue\Unknow.txt" For Output As #5
      '   Print #5, Now
         Print #5,
      Close #5
   Open Trim(txtBomPath.Text) For Input As #1
      Do Until EOF(1)
        Line Input #1, Mystr
         strReadText = Mystr
           Mystr = Trim(UCase(Mystr))
           
             If Mystr <> "" And Left(Mystr, 1) <> "!" Then
                If Left(Mystr, 1) <> "-" Then
                  TmpStr = Split(Mystr, " ")
                  strDeviceNomber = Trim(TmpStr(0))
                  strDeviceName = TmpStr(UBound(TmpStr))
                  tmpSTR1 = Trim(tmpSTR1)
                  tmpSTR1 = Trim(Replace(Mystr, TmpStr(0), ""))
                  TmpStr = Split(tmpSTR1, " ")
                  DeviceType_ = Trim(TmpStr(0))
                  Select Case DeviceType_
                    
                     Case "CONN"
                     Case "SKT"
                     Case "HEAD"
                     Case "EMI"
                     Case "BOSS"
                     Case "2HIP"
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;" '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                     
                     
                     Case "SKT"
                     Case "CHIP"
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         tmpSTR1 = Trim(tmpSTR1)
                         TmpStr = Split(tmpSTR1, " ")
                         DeviceType_A = Trim(TmpStr(0))
                          If Len(DeviceType_A) > 1 Then
                             Select Case Left(DeviceType_A, 3)
                                 
                               Case "CAP"
                                 If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                               Else
                                                 CValue = strCAP(0)
                                               
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                     
                                 End If 'bListCatacitor=true
                                 
                               Case "RES"
                                 If bListResistor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "RES", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strRES = Split(tmpSTR1, " ")
                                   RValue = strRES(0)
                                       RValue1 = Val(RValue)
                                    If Right(RValue, 1) <> "K" And Right(RValue, 1) <> "M" And InStr(RValue, "K") = 0 And InStr(RValue, "M") = 0 Then
                                       If CheckJumper.Value = 1 Then
                                           
                                           LowToJumper = Val(txtJumper.Text)
                                        If RValue1 < LowToJumper Then
                                           Print #6, strDeviceName; Tab(25); "CLOSED;" '; Tab(35); ' "PN""" & strDeviceName & """  ;      !BOM Value: " & RValue
                                           strReadText = "OK"
                                           Else
                                             Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                            strReadText = "OK"
                                        End If
                                      End If
                                      Else
                                         Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                          strReadText = "OK"
                                   End If
                                   
'                                   Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
'                                    strReadText = "OK"
                                 End If
                               Case "LED"
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                       strReadText = "OK"
                                   End If
                                  
                                  
                               Case "FUS"
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;" '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                               Case "NTW"
                                   If bListPinLib = True Then
                                        Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                      strReadText = "OK"
                                   End If
                               Case "BEA"
                                   If CheckIND.Value = 1 Then
                                       Print #6, strDeviceName; Tab(25); "CLOSED;" ' ;Tab(35); "PN""" & strDeviceName & """  ;"
                                       strReadText = "OK"
                                    End If
                               
                               Case "CHO"
                                   If CheckIND.Value = 1 Then
                                       Print #6, strDeviceName; Tab(25); "CLOSED;" '; Tab(35); "PN""" & strDeviceName & """  ;"
                                       strReadText = "OK"
                                    End If
                               
                               Case "IND"
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;" '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                                
                               Case "VAR"
                              Case "0.0"
                                 'If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                               Else
                                                 CValue = strCAP(0)
                                               
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                     
                                ' End If 'bListCatacitor=true
                               
                               
                               
                               
                             End Select
                          Else
                             If Left(DeviceType_A, 1) = "C" Then
                                 If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                            
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                               strReadText = "OK"
                                               
                                               Else
                                                 CValue = strCAP(0)
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                 End If 'bListCatacitor=true
 
                    
                    
                    
                           End If 'Left(DeviceType_A, 1) = "C"
                             
                             
                          End If 'Len(DeviceType_A) > 1
                     Case "IC"
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                       strReadText = "OK"
                                   End If
                     Case "XFORM"
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                       strReadText = "OK"
                                   End If
                     
                     Case "THERM"
                         
  'great 091123
  
                                If bListResistor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "THERM", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strRES = Split(tmpSTR1, " ")
                                   RValue = strRES(0)
                                       RValue1 = Val(RValue)
                                    If Right(RValue, 1) <> "K" And Right(RValue, 1) <> "M" And InStr(RValue, "K") = 0 And InStr(RValue, "M") = 0 Then
                                       If CheckJumper.Value = 1 Then
                                           
                                           LowToJumper = Val(txtJumper.Text)
                                        If RValue1 < LowToJumper Then
                                           Print #6, strDeviceName; Tab(25); "CLOSED;" '; Tab(35); "PN""" & strDeviceName & """  ;      !BOM Value: " & RValue
                                           strReadText = "OK"
                                           Else
                                             Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                            strReadText = "OK"
                                        End If
                                      End If
                                      Else
                                         Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
                                          strReadText = "OK"
                                   End If
                                   
'                                   Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
'                                    strReadText = "OK"
                                 End If
                                  
                                  
  
  
  
  
  
                     Case "IR"
                       
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                      strReadText = "OK"
                                   End If
                     Case "RESO"
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                       strReadText = "OK"
                                   End If
                       
                     Case "XTAL"
                     
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                      strReadText = "OK"
                                   End If
                     
                     Case "DIODE"
                       If bListDiode = True Then
                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """    ;"
                           strReadText = "OK"
                       End If
                     
                     Case "DIODES"
                       If bListDiode = True Then
                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """    ;"
                           strReadText = "OK"
                       End If
                     
                     Case "LED"
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                       strReadText = "OK"
                                   End If
                     
                     
                     Case "XTOR"
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                                   
                                      strReadText = "OK"
                                   
                                   End If
                     Case "FET"
                      If bListPinLib = True Then
                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """  ;"
                           strReadText = "OK"
                      
                      End If
                     Case "STANDOFF"

'                                If CheckIND.Value = 1 Then
'                                   Print #6, strDeviceName; Tab(25); "OPEN"; Tab(35); "PN""" & strDeviceName & """  ;"
'                                   strReadText = "OK"
'                                End If
                     
                  End Select


                    
                End If 'Left(Mystr, 1) <> "-"
             End If 'Mystr <> ""
            If strReadText <> "OK" Then
                Open PrmPath & "ReadBomValue\Unknow.txt" For Append As #3
                   Print #3, strReadText
                Close #3
                
            End If
            strReadText = ""
            DoEvents
            tmpSTR1 = ""
            Mystr = ""
            strDeviceName = ""
            CValue = ""
            RValue = ""
             RValue1 = ""
             strDeviceNomber = ""
      Loop
   Close #1
   If bListCatacitor = True Then
     Close #2
   End If
  
 ' Resistor
  
   If bListResistor = True Then
     Close #4
   End If
     Close #6
   '  Close #7
   If bListDiode = True Then
      Close #8
   End If
   
  If bListPinLib = True Then
      Close #9
   End If
   MsgBox "OK" & Chr(13) & Chr(10) & "File save path:" & PrmPath & "ReadBomValue\", vbInformation
  
Exit Sub
EX:
MsgBox Err.Description, vbCritical

End Sub



Private Sub Creat_Basic_Bom(strPath As String, strPartName As String, Mystr As String, strBoardver As String)
 On Error Resume Next
 Dim tmptmpStr As String
 MkDir PrmPath & "BomCompare\Basic"
          strPath = Trim(strPath)
          If Dir(PrmPath & "BomCompare\Basic\" & strPath) = "" Then
               Open PrmPath & "BomCompare\Basic\" & strPath For Output As #22
               Print #58, strTmp(UBound(strTmp))
          Else
          
              Open PrmPath & "BomCompare\Basic\" & strPath For Append As #22
          End If
        '  Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) For Output As #22
               'Print #22, strBom1_DeviceName
               tmptmpStr = Replace(Mystr, strPath, "")
               tmptmpStr = Trim(Replace(tmptmpStr, strPath, ""))
               Print #22, strPath & "," & strPartName & "," & tmptmpStr & "," & strBoardver
          Close #22
End Sub

Private Sub txtVer_1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom2 = True And Trim(txtVer_1.Text) <> "" Then
  txtVer_2.SetFocus
End If
End Sub
Private Sub txtVer_2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom3 = True And Trim(txtVer_2.Text) <> "" Then
  txtVer_3.SetFocus
End If
End Sub
Private Sub txtVer_3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom4 = True And Trim(txtVer_3.Text) <> "" Then
  txtVer_4.SetFocus
End If
End Sub
Private Sub txtVer_4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom5 = True And Trim(txtVer_4.Text) <> "" Then
  txtVer_5.SetFocus
End If
End Sub

Private Sub txtVer_5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom6 = True And Trim(txtVer_5.Text) <> "" Then
  txtVer_6.SetFocus
End If
End Sub
Private Sub txtVer_6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom7 = True And Trim(txtVer_6.Text) <> "" Then
  txtVer_7.SetFocus
End If
End Sub
Private Sub txtVer_7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bRunBom8 = True And Trim(txtVer_7.Text) <> "" Then
  txtVer_8.SetFocus
End If
End Sub
Private Sub txtVer_1_LostFocus()
txtVer_1.Text = UCase(txtVer_1.Text)
End Sub

Private Sub txtVer_2_LostFocus()
txtVer_2.Text = UCase(txtVer_2.Text)
End Sub

Private Sub txtVer_3_LostFocus()
txtVer_3.Text = UCase(txtVer_3.Text)
End Sub

Private Sub txtVer_4_LostFocus()
txtVer_4.Text = UCase(txtVer_4.Text)
End Sub

Private Sub txtVer_5_LostFocus()
txtVer_5.Text = UCase(txtVer_5.Text)
End Sub

Private Sub txtVer_6_LostFocus()
txtVer_6.Text = UCase(txtVer_6.Text)
End Sub

Private Sub txtVer_7_LostFocus()
txtVer_7.Text = UCase(txtVer_7.Text)
End Sub

Private Sub txtVer_8_LostFocus()
txtVer_8.Text = UCase(txtVer_8.Text)
End Sub
