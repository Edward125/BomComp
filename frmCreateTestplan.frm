VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateTestplan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Put forward boards..."
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton comGo 
      Caption         =   "&Go"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTestplan 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Please open testplan file!(DblClick me open file!"")"
      Top             =   120
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      Begin VB.Label msg1 
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmCreateTestplan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTestplanPath As String
Dim bRunTestplan As Boolean

Private Sub cmdCancel_Click()
frmMain.Show
Unload Me
End Sub

Private Sub comGo_Click()
 On Error Resume Next
  Kill PrmPath & "BomCompare\Panel_Boards_Testplan\*.*"
 
 If bRunTestplan = True Then
     comGo.Enabled = False

     txtTestplan.Enabled = False
   Call Open_Testplan_Read
   Else
    If bRunTestplan = False Then
        MsgBox "Please check testplan path!", vbCritical
         comGo.Enabled = True
          txtTestplan.Enabled = True
        txtTestplan.SetFocus
        
        Exit Sub
     End If
 End If
 comGo.Enabled = True
 txtTestplan.Enabled = True
 Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

PrmPath = App.Path
If Right(PrmPath, 1) <> "\" Then PrmPath = PrmPath & "\"
MkDir PrmPath & "BomCompare"
MkDir PrmPath & "BomCompare\Panel_Boards_Testplan"



Open PrmPath & "BomCompare\NotDelete.sys" For Output As #77

End Sub

Private Sub Open_Testplan_Read()
Dim Mystr As String
Dim strAnalogName As String
Dim bSubAnalog As Boolean
Dim i 'As Integer
Dim t 'As Integer
Dim BoardsNumber
On Error Resume Next
Kill PrmPath & "BomCompare\Panel_boards_testplan\*.*"

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
Msg1(0).Caption = "Reading testplan file..."

  Open PrmPath & "BomCompare\UnsettledDevice.txt" For Output As #23
  Close #23
 Open PrmPath & "BomCompare\UnsettledDevice.txt" For Append As #2
  Open strTestplanPath For Input As #1
      Do Until EOF(1)
       Line Input #1, Mystr
         Mystr = LCase(Trim(Mystr))
         If Mystr <> "" Then
         
                  If InStr(Replace(LCase(Mystr), " ", ""), "subanalog_tests(") <> 0 Then
                     bSubAnalog = True
                  End If
                   If LCase(Trim(Mystr)) = "subend" And bSubAnalog = True Then
                       Exit Do
                   End If
                   If Left(LCase(Trim(Mystr)), 6) = "subend" And bSubAnalog = True Then
                       Exit Do
                   End If
                  If InStr(Replace(LCase(Mystr), " ", ""), "subend(") <> 0 And bSubAnalog = True Then
                      Exit Do
                  End If
                    tmptext = Replace(LCase(Mystr), " ", "")
                        If InStr(tmptext, "onboardsboardset_boards_") <> 0 And InStr(Mystr, "test ") <> 0 And bSubAnalog = True And InStr(Mystr, strAnalog_) <> 0 Then
                           'BoardsNumber = Right(Trim(Replace(tmptext, "(*)", "")), 1)
                           tmptext = Right(tmptext, (Len(tmptext) - InStr(tmptext, "onboardsboardset_boards_")) + 1)
                           tmptext = Replace(LCase(tmptext), "onboardsboardset_boards_", "")
                           tmptext = Replace(LCase(tmptext), "to_", "")
                           tmptext = Left(tmptext, InStr(tmptext, "_"))
                           BoardsNumber = Trim(Replace(LCase(tmptext), "_", ""))
                                              
                                If BoardsNumber <> "" Then
                                      If Left(Mystr, 1) = "!" And InStr(Replace(Mystr, " ", ""), "testcommentedintestorder") <> 0 Then
                                         
                                         Mystr = Replace(Mystr, " ", "")
                                         Mystr = Replace(Mystr, "testcommentedintestorder", "")
                                         Mystr = Replace(Mystr, "onboardsboardset_boards_" & BoardsNumber & "_to_" & BoardsNumber & "(*)", "")
                                         Mystr = Replace(Mystr, "!", "")
                                         Mystr = Replace(Mystr, "test", "test ")
                                         Mystr = "!   " & Mystr & "  ! test commented in testorder"
                                           If Dir(PrmPath & "BomCompare\Panel_boards_testplan\Testplan_Board_Number_" & BoardsNumber & ".vb_Create") = "" Then
                                              Open PrmPath & "BomCompare\Panel_boards_testplan\Testplan_Board_Number_" & BoardsNumber & ".vb_Create" For Output As #3
                                                 Print #3, "sub Analog_Tests" & ",on boards BoardSet_boards_" & BoardsNumber & "_to_" & BoardsNumber & "(*)"
                                                 Print #3, "     " & Mystr
                                              Close #3
                                             Else
                                              Open PrmPath & "BomCompare\Panel_boards_testplan\Testplan_Board_Number_" & BoardsNumber & ".vb_Create" For Append As #3
                                                 Print #3, "     " & Mystr
                                              Close #3
                                           End If
                                         BoardsNumber = ""
                                      End If
                                 End If
                               
                           If BoardsNumber <> "" Then
                                    Mystr = Replace(Mystr, " ", "")
                                    Mystr = Replace(Mystr, "onboardsboardset_boards_" & BoardsNumber & "_to_" & BoardsNumber & "(*)", "")
                                    Mystr = Replace(Mystr, "test", "test ")
                                    If Dir(PrmPath & "BomCompare\Panel_boards_testplan\Testplan_Board_Number_" & BoardsNumber & ".vb_Create") = "" Then
                                         Open PrmPath & "BomCompare\Panel_boards_testplan\Testplan_Board_Number_" & BoardsNumber & ".vb_Create" For Output As #3
                                            Print #3, "sub Analog_Tests " & ",on boards BoardSet_boards_" & BoardsNumber & "_to_" & BoardsNumber & "(*)"
                                            Print #3, "     " & Mystr
                                         Close #3
                                        Else
                                         Open PrmPath & "BomCompare\Panel_boards_testplan\Testplan_Board_Number_" & BoardsNumber & ".vb_Create" For Append As #3
                                            Print #3, "     " & Mystr
                                         Close #3
                                    End If
                           End If
                          Else
                           Print #2, Mystr
                    End If

         End If '<>""
          i = i + 1
          tmptext = ""
          BoardsNumber = ""
          Mystr = ""
          Msg1(0).Caption = "Reading testplan file... Read file line:" & i
        DoEvents
      Loop
  Close #1
  Close #2

  
     If bSubAnalog = False Then
        Msg1(0).Caption = "The testplan is bad file!"
        bRunTestplan = False
        
     End If
    Msg1(0).Caption = "Testplan closed! & boards testplan create ok!"
    
   bSubAnalog = False
   MsgBox "Boards testplan create ok! In " & PrmPath & "BomCompare\Panel_boards_testplan\", vbInformation
End Sub




Private Sub txtTestplan_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = "testplan"
    .Filter = "testplan file|*testplan*.*|*.txt|*.txt|*.*|*.*"
    .ShowOpen

End With
    txtTestplan.Text = Me.CommonDialog1.FileName
    bRunTestplan = True
    If Dir(txtTestplan.Text) = "" Then
      txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtTestplan.SetFocus
        strTestplanPath = ""
        bRunTestplan = False
      Exit Sub
    End If
 
Exit Sub

errh:
MsgBox Err.Description, vbCritical
    txtTestplan.Text = " Please open testplan file!(DblClick me open file!)"
    txtTestplan.SetFocus
        strTestplanPath = ""
        bRunTestplan = False
End Sub
