VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Testplan and bom version compare"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10275
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
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
      Begin VB.CommandButton cmdBomAndBom 
         Caption         =   "Bom 1 and Bom 2       &GO>>"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Bom and bom compare (only two bom file)"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Bom and testplan compare"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
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
         Height          =   255
         Left            =   4680
         TabIndex        =   23
         Top             =   2280
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   2640
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   21
         Top             =   3000
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   3360
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   840
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   1560
         Width           =   2655
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
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   1920
         Width           =   2655
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
Frame3.Enabled = False
If bTwoBom = True Then
    Open PrmPath & "BomCompare\Bom1_and_Bom2_Compare.txt" For Output As #54
    Close #54
    Call BomAndBomComp_Bom1
    Call BomAndBomComp_Bom2

   
    Call Bom1AndBom2_Dir_Bom2
      msg1.Caption = l1.Caption & "_"
      Msg2.Caption = "_ and _"
      Msg3.Caption = "_" & l2.Caption
      Msg4.Caption = "_ compare ok!"
      Frame3.Enabled = True
        Call Kill_File
    MsgBox l1.Caption & " and " & l2.Caption & " compare ok!", vbInformation

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
        Print #54, "!============================" & l1.Caption & " file  =True ," & l2.Caption & " file not find devices============================="
   Open strBom2Path For Input As #52
           Do Until EOF(52)
             Line Input #52, strBom2_DeviceName
               msg1.Caption = "Reading bom2 file..."
               Mystr = LCase(Trim(strBom2_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                             If Dir(PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) & ".shit") = "" Then
                                Print #54, strBom2_DeviceName
                             End If
                             
                              Open PrmPath & "BomCompare\Bom_2\" & strTmp(UBound(strTmp)) & "." & strTmp(0) & ".shit" For Output As #22
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
        Print #54, "!============================" & l2.Caption & " file  =True ," & l1.Caption & " file not find devices============================="
 Close #54
        msg1.Caption = l2.Caption & " file closed!"
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
               msg1.Caption = "Reading bom1 file..."
               Mystr = LCase(Trim(strBom1_DeviceName))
               If Mystr <> "" Then
                  If Left(Mystr, 1) <> "-" Then
                    strTmp = Split(Mystr, " ")
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                             'create bom1 analog file
                              Open PrmPath & "BomCompare\Bom_1\" & strTmp(UBound(strTmp)) & "." & strTmp(0) & ".shit" For Output As #22
                                   Print #22, strBom1_DeviceName
                              Close #22
                                   Print #58, strTmp(UBound(strTmp)) & "." & strTmp(0) & ".shit"
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
        msg1.Caption = l1.Caption & " file closed!"
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
     comOK.Enabled = False
     cmdBoards.Enabled = False
     Check1.Enabled = False
     Check2.Enabled = False
     txtBom1.Enabled = False
     txtBom2.Enabled = False
     txtBom3.Enabled = False
     txtBom4.Enabled = False
     txtBom5.Enabled = False
     txtBom6.Enabled = False
     txtBom7.Enabled = False
     txtBom8.Enabled = False
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
     comOK.Enabled = True
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



msg1.Caption = "Start compare..."
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

     comOK.Enabled = True
     cmdBoards.Enabled = True
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
  Do While MyName <> ""   ' 开始循环
     tmpfile = Trim(Replace(MyName, ".shit", ""))
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
      
    MyName = Dir   ' 查找下一个目录。
    DoEvents
  Loop
Close #7





  Kill PrmPath & "BomCompare\Testplan_Tmp_Analog\*.shit"
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
        
                           
                     Else
                                Open PrmPath & "BomCompare\Comm_Device_Exist.txt" For Append As #5
                                  If Boards = True Then
                                   'strBoardsNumber
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & " " & strBoardsNumber & "  ! test commented in testorder"
                                    Else
                                     Print #5, "   ! test " & """" & strAnalog_ & Mystr & """" & "  ! test commented in testorder"
                                  End If
                                Close #5
                End If
                
             End If
             
             If (bBom1OK And bBom2OK And bBom3OK And bBom4OK And bBom5OK And bBom6OK And bBom7OK And bBom8OK) = False Then
                If bTestorder = True Then

                         
                     
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
msg1.Caption = "Reading testplan file..."

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
                                      msg1.Caption = "Testplan file read ok!"
                                      Msg2.Caption = "Find subend ok"
                                      Msg3.Caption = "Device(+  % rX , % crX)  :" & t
                                      t = 0
                                      Msg4.Caption = ""
                                      
                                    Exit Do
              End If
                     
              If bSubAnalog = True And Left(Trim(LCase(Mystr)), 6) = "subend" Then
                                      msg1.Caption = "Testplan file read ok!"
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
                                      msg1.Caption = "Testplan file read ok!"
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
                                                            Print #4, strAnalogName & ".shit"
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
                                                                   Print #4, strAnalogName & ".shit"
                                                                Close #4

                                                      End If 'InStr(strAnalogName, "_") <> 0
                                               End If '<>"%"
                                           strAnalogName = LCase(Trim(Replace(strAnalogName, "_", "")))
                                              strAnalogName = LCase(Trim(Replace(strAnalogName, "%", "")))
                                            '  baobei
                                        
                                             'create analog device file
                                             Open PrmPath & "BomCompare\Testplan_Tmp_Analog\" & strAnalogName & ".shit" For Output As #2
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
        msg1.Caption = "The testplan is bad file!"
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
       msg1.Caption = "Reading bom1 file..."
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
msg1.Caption = "Bom1 file closed!"
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
       msg1.Caption = "Reading Bom2 file..."
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
 msg1.Caption = "Bom2 file closed!"
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
       msg1.Caption = "Reading Bom3 file..."
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
 msg1.Caption = "Bom3 file closed!"
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
       msg1.Caption = "Reading Bom4 file..."
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
 msg1.Caption = "Bom4 file closed!"
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
       msg1.Caption = "Reading Bom5 file..."
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
msg1.Caption = "Bom5 file closed!"
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
       msg1.Caption = "Reading Bom6 file..."
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
msg1.Caption = "Bom6 file closed!"
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
       msg1.Caption = "Reading Bom7 file..."
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
msg1.Caption = "Bom7 file closed!"
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
       msg1.Caption = "Reading Bom8 file..."
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
msg1.Caption = "Bom8 file closed!"
If i = 0 Then
  MsgBox "Shit ,the Bom8 file is null!", vbCritical
  Exit Sub
End If
End Sub

 

Private Sub Form_Load()
On Error Resume Next
 
 
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
'
 
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
   Else
   Frame4.Enabled = False
End If
If Option2.Value = True Then
   txtBom3.Enabled = False
   txtBom4.Enabled = False
   txtBom5.Enabled = False
   txtBom6.Enabled = False
   txtBom7.Enabled = False
   txtBom8.Enabled = False
   txtTestplan.Enabled = False
   cmdBomAndBom.Enabled = True
   Else
   txtBom2.Enabled = True
   txtBom3.Enabled = True
   txtBom4.Enabled = True
   txtBom5.Enabled = True
   txtBom6.Enabled = True
   txtBom7.Enabled = True
   txtBom8.Enabled = True
   txtTestplan.Enabled = True
   cmdBomAndBom.Enabled = False
 
End If
End Sub

Private Sub Option2_Click()
If Option1.Value = True Then
   Frame4.Enabled = True
   Else
   Frame4.Enabled = False
End If
If Option2.Value = True Then
   txtBom3.Enabled = False
   txtBom4.Enabled = False
   txtBom5.Enabled = False
   txtBom6.Enabled = False
   txtBom7.Enabled = False
   txtBom8.Enabled = False
   txtTestplan.Enabled = False
   cmdBomAndBom.Enabled = True
   Else
   txtBom2.Enabled = True
   txtBom3.Enabled = True
   txtBom4.Enabled = True
   txtBom5.Enabled = True
   txtBom6.Enabled = True
   txtBom7.Enabled = True
   txtBom8.Enabled = True
   txtTestplan.Enabled = True
   cmdBomAndBom.Enabled = False
 
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
