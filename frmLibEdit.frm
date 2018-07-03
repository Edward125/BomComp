VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLibEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   6120
      TabIndex        =   4
      Top             =   2160
      Width           =   4575
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
         TabIndex        =   8
         Top             =   360
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
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   6
         Top             =   1200
         Width           =   4335
      End
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
         TabIndex        =   5
         Top             =   1200
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdBom 
      Caption         =   "Bom To 3070 Board Format File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "OutLibFile"
      Height          =   375
      Left            =   9480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Out Pin Lib File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtTestplan 
      BackColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Please Open File!(DblClick me open file!"")"
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmLibEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtPinLib_Change()

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   cmdGo.Enabled = True
   cmdBom.Enabled = False
   Else
   cmdGo.Enabled = False
   cmdBom.Enabled = True
End If
End Sub

Private Sub cmdBom_Click()

End Sub
Private Sub BomAndBomComp_Bom1()
 Dim strBom1_DeviceName As String
 Dim intDevice_Ge As Integer
 Dim strDevice_LiaoHao As String
 Dim strDevice_Text As String
 Dim intFile_Line As Integer
 Dim MyStr As String
 Dim strMyStr As String
 Dim strTmp() As String
 Dim strText() As String
 On Error Resume Next
  
intFile_Line = 0
intDevice_Ge = 0
strBom1Path = Trim(txtTestplan.Text)
If Dir(strBom1Path) = "" Then
   txtTestplan.Text = " Please open file!(DblClick me open file!)"
   txtTestplan = False
   txtTestplan.SetFocus
    strBom1Path = ""
   MsgBox "Bom1 not find!,please check path!", vbCritical
   Exit Sub
End If
      Open PrmPath & "BomCompare\3070_Connect.txt" For Output As #7
      Open PrmPath & "BomCompare\3070_Jumper.txt" For Output As #6
      Open PrmPath & "BomCompare\3070_Catacitor.txt" For Output As #3
      Open PrmPath & "BomCompare\3070_Resistor.txt" For Output As #4
      Open PrmPath & "BomCompare\3070_Diode.txt" For Output As #8
      Open PrmPath & "BomCompare\3070_Pin_Linrary.txt" For Output As #9
      Open PrmPath & "BomCompare\3070_Unknow.txt" For Output As #5




'open bom1 file
    Open PrmPath & "BomCompare\Bom_TO_Board.txt" For Output As #1
   Open strBom1Path For Input As #2
   
           Do Until EOF(2)
             Line Input #2, strBom1_DeviceName
               Msg1.Caption = "Reading bom file..."
               MyStr = LCase(Trim(strBom1_DeviceName))
               If MyStr <> "" Then
                  If Left(MyStr, 1) <> "-" Or Left(MyStr, 1) <> "!" Then
                  
 '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                     
                               MyStr1 = Replace(MyStr, " ", ",")
                    strTmp = Split(Replace(MyStr1, Chr(9), ""), ",")
  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                       strTmp(0) = Trim(strTmp(0))
                    
                      If Trim(strTmp(UBound(strTmp))) <> "" Then
                         If Dir(PrmPath & "Pin_Lib\" & strTmp(0)) <> "" Then
                                Open PrmPath & "Pin_Lib\" & strTmp(0) For Input As #10
                                     Input #10, strMyStr
                                     strMyStr = Trim(UCase(strMyStr))
                                     strText = Split(strMyStr, Chr(9))
                                     If strText(3) = "CAPACITOR" Then
                                        Print #3,
                                     End If
                                Close #10
                            Else
                              Print #5, MyStr
                         End If

                             intDevice_Ge = intDevice_Ge + 1

                      End If
                  End If '<>-
                  
               End If '<>""
                intFile_Line = intFile_Line + 1

               DoEvents
               
              
           Loop
 Close #2
 Close #1
 Close #3
 
  Close #4
 Close #5
 Close #6
  Close #7
 Close #8
 Close #9
        If intDevice_Ge = 0 Then
          MsgBox "Shit ,the bom1 file is null!", vbCritical
          Exit Sub
        End If
   
 
End Sub
Private Sub txtTestplan_DblClick()
On Error GoTo errh
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    .FileName = "Pin Lib Msg File"
    .Filter = "*.txt|*.txt|*.*|*.*"
 
    .ShowOpen
    txtTestplan.Text = Me.CommonDialog1.FileName
      bRunTestplan = True
    If Dir(txtTestplan.Text) = "" Then
      txtTestplan.Text = " Please open file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        strTestplanPath = ""
        bRunTestplan = False
      Exit Sub
      
   End If


      
      
      frmLibEdit.Show
      
      
     
    'strBom1Path
End With
Exit Sub
errh:
      txtTestplan.Text = " Please open file!(DblClick me open file!)"
 
        strTestplanPath = ""
        bRunTestplan = False
MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdGo_Click()
Dim tmpStr() As String
Dim MyStr As String
Dim strMyStr As String
Dim strPrintTmp As String
Dim I As Integer
I = 0
  On Error Resume Next
 MkDir PrmPath & "Pin_Lib"
Open txtTestplan.Text For Input As #1
         Do Until EOF(1)
             Line Input #1, MyStr
             strMyStr = Trim(UCase(MyStr))
         If strMyStr <> "" Then
           'tmpStr = Split(strMyStr, " ")
             tmpStr = Split(strMyStr, Chr(9))
             Open PrmPath & "Pin_Lib\" & tmpStr(0) For Output As #2
             
                  Print #2, strMyStr
             Close #2
'             For g = 0 To UBound(tmpStr)
'
'                If tmpStr(g) <> "" Then
'
'                    strPrintTmp = strPrintTmp & " ," & Trim(tmpStr(g))
'
'                End If
'             Next
'             strPrintTmp = Trim(strPrintTmp)
'             If Left(strPrintTmp, 1) = "," Then
'                 strPrintTmp = Right(strPrintTmp, Len(strPrintTmp) - 1)
'             End If
'
'             strPrintTmp = ""
            strMyStr = ""
          End If
          DoEvents
          Msg1.Caption = "Findd Pard Number:" & I
          I = I + 1
         Loop
Close #1
 MsgBox "Creat Pin Lib File OK!", vbInformation
End Sub

