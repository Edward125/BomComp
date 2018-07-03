VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBomValue 
   Caption         =   "Read BOM device value to 3070board format..3.0"
   ClientHeight    =   4455
   ClientLeft      =   6585
   ClientTop       =   945
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8145
   Begin VB.Frame Frame3 
      Caption         =   "Cad device part number"
      Enabled         =   0   'False
      Height          =   4455
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   7935
      Begin VB.TextBox txtCadPartFile 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Please open *.asc file!(DblClick me open file!)"
         Top             =   240
         Width           =   7695
      End
      Begin VB.TextBox txtBomPath_2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Please open bom file!(DblClick me open file!)"
         Top             =   1560
         Visible         =   0   'False
         Width           =   6015
      End
      Begin VB.CommandButton cmdCreatDevicePart 
         Caption         =   "CreatDevicePart"
         Height          =   855
         Left            =   4680
         TabIndex        =   33
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   3615
         Left            =   360
         Picture         =   "frmBomValue.frx":0000
         Top             =   720
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7935
      Begin VB.CommandButton mcGo 
         Caption         =   "&GO"
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
         Left            =   6600
         TabIndex        =   1
         Top             =   280
         Width           =   1095
      End
      Begin VB.CheckBox Checklb 
         Caption         =   "Pin Linrary"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox CheckR 
         Caption         =   "Resistor"
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox CheckD 
         Caption         =   "Diode"
         ForeColor       =   &H000040C0&
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox CheckCn 
         Caption         =   "Connector"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5160
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox CheckC 
         Caption         =   "Catacitor"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.TextBox txtBomPath 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Please open bom file!(DblClick me open file!)"
      Top             =   240
      Width           =   6735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   7935
      Begin VB.CheckBox cBoardHead 
         Caption         =   "Output Board Head"
         Height          =   255
         Left            =   4920
         TabIndex        =   36
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox cUndefined 
         Caption         =   "Undefined Device Check"
         Height          =   255
         Left            =   2400
         TabIndex        =   32
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CheckBox cDisplayName 
         Caption         =   "Display Device Name"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   7320
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Output Agilent 3070 Version Board Format File."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   29
         Top             =   1680
         Width           =   6375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtDH 
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   25
         Text            =   "0.8"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtDL 
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
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "0.2"
         Top             =   840
         Width           =   495
      End
      Begin VB.CheckBox CheckIND 
         Caption         =   "Inductance and Fuse to Jumper test"
         Height          =   375
         Left            =   4560
         TabIndex        =   23
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtJumper 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   20
         Text            =   "10"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox CheckJumper 
         Caption         =   "Resistor value low to Jumper"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtRL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "10"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtRH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   7200
         MaxLength       =   3
         TabIndex        =   16
         Text            =   "10"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "20"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtCH 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "20"
         Top             =   360
         Width           =   495
      End
      Begin VB.Line Line4 
         X1              =   5400
         X2              =   5520
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Resistor"
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Line Line3 
         X1              =   5520
         X2              =   5640
         Y1              =   600
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   5520
         X2              =   5640
         Y1              =   960
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   5520
         X2              =   5520
         Y1              =   960
         Y2              =   600
      End
      Begin VB.Label Label10 
         Caption         =   "Diode High Limit:"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Diode Low Limit:"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "o"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Value low >"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   1400
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Resistor Low Limit:"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Resistor High Limit:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Catacitor Low Limit:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Catacitor High Limit:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.TextBox txtBoardHead 
      Height          =   735
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "frmBomValue.frx":54FA6
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "BomPath:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmBomValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BomPath As String
Dim PrmPath As String
Dim bListCatacitor As Boolean
Dim bListResistor As Boolean
Dim bListDiode As Boolean
Dim bListConnect As Boolean
Dim strCH As String
Dim strCL As String
Dim strRH As String
Dim strRL As String
Dim strDH As String
Dim strDL As String
Dim strBom_All_Device() As String
Dim intBom_All_Number As Integer
Dim bConnectFindOk As Boolean
Dim PinLibOk As Boolean
Dim tmpPartNumber As String
 

Private Sub CheckJumper_Click()
If CheckJumper.Value = 1 Then
  txtJumper.Enabled = True
  Else
   
  txtJumper.Enabled = False
End If
End Sub

Private Sub CheckR_Click()
If CheckR.Value = 1 Then
  CheckJumper.Enabled = True
  'txtJumper.Enabled = True
  Else
 CheckJumper.Enabled = False
  txtJumper.Enabled = False
   CheckJumper.Value = 0
  
End If
End Sub

Private Sub cmdCreatDevicePart_Click()
  Call CreatDevicePart_File
   MsgBox "Creat Cad Device Part Number List OK!", vbInformation
End Sub

Private Sub Command1_Click()
Unload Me
 
End Sub

Private Sub cUndefined_Click()
 If cUndefined.Value = 1 Then
     Frame3.Enabled = True
     Me.Height = 9480
     Else
     Frame3.Enabled = False
     Me.Height = 4965
 End If
End Sub

Private Sub Form_Load()
On Error Resume Next
 
 
PrmPath = App.Path
If Right(PrmPath, 1) <> "\" Then PrmPath = PrmPath & "\"
MkDir PrmPath & "ReadBomValue"

End Sub

Private Sub Read_BomFile()
On Error Resume Next
Dim Mystr As String
Dim intI As String
Dim strDeviceName As String
Dim TmpStr() As String
Dim strXXX() As String
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
Dim strPN_Str As String
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
If CheckCn.Value = 1 Then
   bListConnect = True
   Else
   bListConnect = False
End If



  ' Open PrmPath & "ReadBomValue\WaitCheck.txt" For Output As #7
  '   Print #7, Now
   '   Print #7,
   Open PrmPath & "ReadBomValue\Jumper.txt" For Output As #6
      Print #6, "JUMPER"
      Print #6,
   If bListCatacitor = True Then
      Open PrmPath & "ReadBomValue\Capacitor.txt" For Output As #2
      Print #2, "CAPACITOR"
      Print #2,
   End If
   If bListResistor = True Then
      Open PrmPath & "ReadBomValue\Resistor.txt" For Output As #4
      Print #4, "RESISTOR"
      Print #4,
   End If
   If bListDiode = True Then
      Open PrmPath & "ReadBomValue\Diode.txt" For Output As #8
      Print #8, "DIODE"
      Print #8,
   
   End If
   
   If bListPinLib = True Then
      Open PrmPath & "ReadBomValue\Pin Library.txt" For Output As #9
      Print #9, "PIN LIBRARY"
      Print #9,
   
   End If
   
 If bListConnect = True Then
      Open PrmPath & "ReadBomValue\Connector.txt" For Output As #24
       Print #24, "CONNECTOR"
      Print #24,
 
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
                   Mystr = Replace(Mystr, ",", " ")
                  Mystr = Replace(Mystr, """", "")
                  Mystr = Replace(Mystr, ",", " ")
                  
                   Mystr = Replace(Mystr, "_", " ")
                   Mystr = Replace(Mystr, "BAT CON1", "BAT_CON1")
                  TmpStr = Split(Mystr, " ")
                  strDeviceNomber = Trim(TmpStr(0))
                  strDeviceName = TmpStr(UBound(TmpStr))
                  ReDim Preserve strBom_All_Device(intBom_All_Number + 1)
                      strBom_All_Device(intBom_All_Number) = strDeviceName
                      intBom_All_Number = intBom_All_Number + 1

                  'PART NUMBER
                  '==================================
                   If strDeviceNomber = "78.10234.1F1" Or strDeviceNomber = "78.22523.5B1" _
                        Or strDeviceNomber = "78.47523.5CL" _
                        Or strDeviceNomber = "77.C1571.09L" _
                     Then
                       Mystr = Replace(Mystr, "CHIP C", "CHIP C ") 'CAP
                   End If
                   '79.47719.9BL pt cat 3pin
                   If strDeviceNomber = "79.47719.9BL" Then
                      Mystr = Replace(Mystr, "CHIP CAP", "CHIP LIB")
                   End If
                  '==================================
                  
                  tmpSTR1 = Trim(tmpSTR1)
                  tmpSTR1 = Trim(Replace(Mystr, TmpStr(0), ""))
                     strPN_Str = tmpSTR1
                   If cDisplayName.Value = 0 Then
                      strXXX = Split(tmpSTR1, "   ")
                      strPN_Str = strXXX(0)
                      Erase strXXX
                   End If
                     
                  TmpStr = Split(tmpSTR1, " ")
                  DeviceType_ = Trim(TmpStr(0))
                  '=============================
                     If Left(DeviceType_, 4) = "POLY" Then
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         DeviceType_ = "POLY"
                         TmpStr(0) = "POLY"
                     End If
                     
                     'THERMISTOR
                     If Trim(TmpStr(0)) = "THERMISTOR" Then
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         DeviceType_ = "THERMIS"
                         TmpStr(0) = "THERMIS"
                     End If
                     
                     If InStr(UCase(Mystr), "ANTENNA SPRING") <> 0 Then
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         DeviceType_ = "ANTENNA"
                         TmpStr(0) = "ANTENNA"
                     End If
                     If Trim(TmpStr(1)) = "EMI" Then
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         DeviceType_ = "EMI"
                         TmpStr(0) = "EMI"
                     End If
                    
                     If DeviceType_ = "CHP" Then DeviceType_ = "CHIP"
                     
                     If DeviceType_ = "MLCC" Then
                          DeviceType_ = "CHIP"
                          TmpStr(0) = "CHIP"
                          tmpSTR1 = Trim(Replace(tmpSTR1, "MLCC", "CHIP CAP"))
                     End If
                     'SSIP CAP
                     If DeviceType_ = "SSIP" Then
                       If InStr(UCase(Mystr), "SSIP CAP") <> 0 Then
                       
                          DeviceType_ = "CHIP"
                          TmpStr(0) = "CHIP"
                          tmpSTR1 = Trim(Replace(tmpSTR1, "SSIP", "CHIP"))
                       End If
                     End If
                     
                     If DeviceType_ = "RES" Then
                          DeviceType_ = "CHIP"
                          TmpStr(0) = "CHIP"
                          tmpSTR1 = Trim(Replace(tmpSTR1, "RES", "CHIP RES"))
                     End If
                      If DeviceType_ = "CAP" Then
                          DeviceType_ = "CHIP"
                          TmpStr(0) = "CHIP"
                          tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", "CHIP CAP"))
                     End If
                     If DeviceType_ = "FERRITE" Then
                          DeviceType_ = "CHIP"
                          TmpStr(0) = "CHIP"
                          tmpSTR1 = Trim(Replace(tmpSTR1, "FERRITE", "CHIP BEAD"))
                     End If
                     If DeviceType_ = "FEREITE" Then
                          DeviceType_ = "CHIP"
                          TmpStr(0) = "CHIP"
                          tmpSTR1 = Trim(Replace(tmpSTR1, "FEREITE", "CHIP BEAD"))
                     End If
                   'diode
                   If strDeviceNomber = "83.01921.P70" Then
                                        If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                         End If
                                          DeviceType_ = ""

                   End If
                     
                  '===================================================
                  Select Case DeviceType_
                  
                  
                    Case "HDD"
                    
                          'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                    
                     Case "H17HV-1A"
                       
                     
                     'conn NT
                         If InStr(Mystr, "SPRING") <> 0 Then
                         
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                         End If
                     Case "OSC"
                        'OSC TO CONN
                       'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "FLASH"
                          'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "TEMP.SENSOR"
                          'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "TRANSLATOR"
                          'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "LNA"
                     
                         'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "ACCELEROMETER"
                         'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "SW"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "802.11B/G/N"
                         'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "POLYSWITCH"
                             'JUMPER
                                If CheckIND.Value = 1 Then
                                    Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str ' ;Tab(35); "PN""" & strDeviceName & """  ;"
                                    strReadText = "OK"
                                 End If
                     
                     Case "ARIES"
                      'conn NT
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "CHASSIS"
                     'conn NT
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "SAW"
                        'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "C.S"
                     
                        'PIN LIB
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                     Case "ANTENNA"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "WTOB"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "SCHOTTKY"
                     
                        'PIN LIB
                                 If TmpStr(1) = "BAT54CW-LSOT323" Then
                                     If bListPinLib = True Then
                                          Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                     End If
                                 End If
                        'PIN LIB
                                 If TmpStr(1) = "BAT54CW-L" Then
                                     If bListPinLib = True Then
                                          Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                     End If
                                 End If
                                 
                                 
                     'diode
                        If TmpStr(1) = "SS0520-L" Then
                            If bListDiode = True Then
                                Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                strReadText = "OK"
                            End If
                        End If
                'PIN LIB
                        If TmpStr(1) = "SBR10U45SP5-13" Then
                            If bListPinLib = True Then
                                 Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                        End If
                        
                        
                        
                        
                     Case "INDUCTOR"
                      'jumper
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                     Case "FPC"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                                       
                   
                     Case "STAND-OFF"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "SPRING"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "ID"
                        'PIN LIB
                                
                                 If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
           
                     Case "EMISPRINGS8-100G"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "EP101TA"
                      'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "BOSS-MB-STAND-OFF-A-LS20"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "POLYSWITH"
                            'JUMPER
                                If CheckIND.Value = 1 Then
                                    Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str ' ;Tab(35); "PN""" & strDeviceName & """  ;"
                                    strReadText = "OK"
                                 End If
                     
                     Case "THERMIS"
                     'RES
                          If Left(UCase(Trim(strDeviceName)), 2) = "RN" Then
                                If bListPinLib = True Then
                                      Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                    strReadText = "OK"
                                 End If
                              
                              
                          
                              Else
                              
                                If bListResistor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "THERMIS", ""))
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
                                             Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                        End If
                                      End If
                                      Else
                                         Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                          strReadText = "OK"
                                   End If
                                   
'                                   Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
'                                    strReadText = "OK"
                                 End If
                            End If
                     
                     Case "POLY"
                     'jumper
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                     
                     Case "POLYSW"
                     'jumper
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                     
                     Case "HOLDER"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     
                     Case "SOCKET"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "COMN"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "CONN"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "SKT"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "HEAD"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "EMI"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If

                     
                     Case "BOSS"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     
                     
                     Case "2HIP"
                     'jump
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                     
                     
                     Case "SKT"
                     'conn
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     
                     Case "CHIP"
   'great
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         tmpSTR1 = Trim(tmpSTR1)
                         TmpStr = Split(tmpSTR1, " ")
                         DeviceType_A = Trim(TmpStr(0))
                         
                         'ANTENNA
                             If UCase(DeviceType_A) = "ANTENNA" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "ANTENNA", "LIB"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "ANTENNA", "LIB"))
                            End If
                         
                         'DUPLEXER
                            If UCase(DeviceType_A) = "DUPLEXER" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "DUPLEXER", "LIB"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "DUPLEXER", "LIB"))
                            End If
                         
                         'POSCAP
                         
                            If UCase(DeviceType_A) = "POSCAP" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "POS", ""))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                            End If
                        'FPCAP
                            
                            If UCase(DeviceType_A) = "FPCAP" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "FP", ""))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "FP", ""))
                            End If
                        'ARRESTER
                            If UCase(DeviceType_A) = "ARRESTER" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "ARRESTER", "OPN"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "ARRESTER", "OPN"))
                            End If
                         'BAED
                            If UCase(DeviceType_A) = "BAED" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "BAED", "BAE"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "BAED", "BAE"))
                            End If
                          'COIL
                            If UCase(DeviceType_A) = "COIL" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "COIL", "BAE"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "COIL", "BAE"))
                            End If
                          'CMM
                           'TRXXX
                            If UCase(DeviceType_A) = "CMM" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "CMM", "LIB"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "CMM", "LIB"))
                            End If
                          'COMM
                            If UCase(DeviceType_A) = "COMM" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "COMM", "LIB"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "COMM", "LIB"))
                            End If
                           'EMIFIL
                             If UCase(DeviceType_A) = "EMIFIL" Then
                               DeviceType_A = Trim(Replace(DeviceType_A, "EMIFIL", "LIB"))
                               tmpSTR1 = Trim(Replace(tmpSTR1, "EMIFIL", "LIB"))
                            End If

                          If Len(DeviceType_A) > 1 Then
                             Select Case Left(DeviceType_A, 3)
                               Case "LIB"
                               'PIN LIB
                                         If bListPinLib = True Then
                                              Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                         End If
                               Case "ESD"
                               '
                                     If Trim(UCase(TmpStr(1))) = "GT1206150ASMD" Then
                                       
                                        'conn
                                               If bListConnect = True Then
                                                  Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                                                  strReadText = "OK"
                                               End If
                                       Else
                                        'diode a c
                                        'MS04A03T2V
                                      If Trim(UCase(TmpStr(3))) = "MS04A03T2V" Then
                                          If bListDiode = True Then
                                             Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                             strReadText = "OK"
                                          End If
                                        
                                        Else
                                            ' PIN LIB
                                              If bListPinLib = True Then
                                                   Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                 strReadText = "OK"
                                           End If
                                       End If
                                     End If
                                 
                               Case "BAE"
                                'BAED  IND TO JUMP
                                 If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                               Case "OPN"
                               'JUMPER OPEN ESD
                                 If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "OPEN;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                               Case "FIL"
                               
                               'PIN LIB
                                 If UCase(Trim(TmpStr(0))) = "FILTER" Then
                                        
                                         If bListPinLib = True Then
                                              Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                         End If
                                 End If
                               Case "CHK"
                               'jump
                                    If CheckIND.Value = 1 Then
                                       Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                       strReadText = "OK"
                                    End If
                               
                               Case "CAP"
                               
                                 If bListCatacitor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "TAN", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "MLCC", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "CAP", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "T", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "F", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NEO", ""))
                                   
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POS", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "POL", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NPO", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C ", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "C", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "EL", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "ARR", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "FP", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "NI ", ""))
                                   
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strCAP = Split(tmpSTR1, " ")

                                     If InStr(strCAP(0), "U") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                         CValue = Left(strCAP(0), InStr(strCAP(0), "U"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                               strReadText = "OK"
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                               Else
                                                 CValue = strCAP(0)
                                               
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                     
                                 End If 'bListCatacitor=true
                                 
                               Case "RES"
                               
                               
                               'RES
                        
                              If Left(UCase(Trim(strDeviceName)), 2) = "RN" Then
                                        If bListPinLib = True Then
                                              Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                         End If
                                Else
                                 If bListResistor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "RES", ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   strRES = Split(tmpSTR1, " ")
                                   RValue = strRES(0)
                                    RValue = Trim(Replace(strRES(0), "F", ""))
                                    RValue = Trim(Replace(strRES(0), "KF", "K"))
                                    
                                       RValue1 = Val(RValue)
                                    If Right(RValue, 1) <> "K" And Right(RValue, 1) <> "M" And InStr(RValue, "K") = 0 And InStr(RValue, "M") = 0 Then
                                       If CheckJumper.Value = 1 Then
                                           
                                           LowToJumper = Val(txtJumper.Text)
                                        If RValue1 < LowToJumper Then
                                           Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); ' "PN""" & strDeviceName & """  ;      !BOM Value: " & RValue
                                           strReadText = "OK"
                                           Else
                                             Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                        End If
                                      End If
                                      Else
                                         Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                          strReadText = "OK"
                                   End If
                                   
'                                   Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
'                                    strReadText = "OK"
                                 End If
                              End If
                                 
                                 
                               Case "LED"
'
'                                   If UCase(Trim(TmpStr(2))) <> "LTW-C191UC5" Then
'
''                                      DIODE
'                                         If bListDiode = True Then
'                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
'                                           strReadText = "OK"
'                                         End If
'
'
'                                      Else
                                      'PIN LIB
                                        If bListPinLib = True Then
                                            Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                        End If
                                  
                                  
                                   
'                                  End If
                                  
                               Case "FUS"
                               'JUMPER
                                If CheckIND.Value = 1 Then
                                   Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                   strReadText = "OK"
                                End If
                               Case "NTW"
                                  'PIN LIB
                                   If bListPinLib = True Then
                                        Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
                               Case "BEA"
                               'JUMPER
                                   If CheckIND.Value = 1 Then
                                       Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str ' ;Tab(35); "PN""" & strDeviceName & """  ;"
                                       strReadText = "OK"
                                    End If
                               
                               Case "CHO"
                               'jumper
                                   
                                       If Left(UCase(strDeviceName), 2) <> "TR" And Trim(UCase(TmpStr(2))) <> "DLW21SN900SQ2L" Then
                                          'JUMP
                                           If CheckIND.Value = 1 Then
                                                Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                                strReadText = "OK"
                                           End If
                                         Else
                                          
                                            'PIN LIB
                                            If bListPinLib = True Then
                                                 Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                               strReadText = "OK"
                                            End If
                                       End If
                               
                               Case "IND"
                               'jumper
                                If InStr(strPN_Str, "CTX17-18765-R") = 0 Then
                                       If CheckIND.Value = 1 Then
                                          Print #6, strDeviceName; Tab(25); "CLOSED;"; Tab(100); "!" & strPN_Str '; Tab(35); "PN""" & strDeviceName & """  ;"
                                          strReadText = "OK"
                                       End If
                                    Else
                                      'PIN LIB
                                            If bListPinLib = True Then
                                                 Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                               strReadText = "OK"
                                            End If
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
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                               strReadText = "OK"
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                               Else
                                                 CValue = strCAP(0)
                                               
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
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
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                        Else
                                         If InStr(strCAP(0), "N") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                             CValue = Left(strCAP(0), InStr(strCAP(0), "N"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                               strReadText = "OK"
                                            
                                            Else
                                             If InStr(strCAP(0), "P") <> 0 And InStr(strCAP(0), "V") <> 0 Then
                                               CValue = Left(strCAP(0), InStr(strCAP(0), "P"))
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                               strReadText = "OK"
                                               
                                               Else
                                                 CValue = strCAP(0)
                                               Print #2, strDeviceName; Tab(25); CValue; Tab(35); strCH; Tab(40); strCL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                             End If 'P,V
                                         End If 'N,V

                                     End If 'U,V
                                 End If 'bListCatacitor=true
 
                    
                    
                    
                           End If 'Left(DeviceType_A, 1) = "C"
                             
                             
                          End If 'Len(DeviceType_A) > 1
                     Case "IC"
                        'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                     Case "XFORM"
                        'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                     
                     Case "THERM"
                     'RES
                         
  'great 091123
                                If Left(UCase(Trim(strDeviceName)), 2) = "RN" Then
                                        If bListPinLib = True Then
                                              Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                         End If
                                Else
                                If bListResistor = True Then
                                   tmpSTR1 = Trim(Replace(tmpSTR1, strDeviceName, ""))
                                   tmpSTR1 = Trim(tmpSTR1)
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "THERM", ""))
                                   tmpSTR1 = Trim(Replace(tmpSTR1, "SMD", ""))
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
                                             Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                        End If
                                      End If
                                      Else
                                         Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                          strReadText = "OK"
                                   End If
                                   
'                                   Print #4, strDeviceName; Tab(25); RValue; Tab(35); strRH; Tab(40); strRL; Tab(45); Tab(50); "f    PN""" & strDeviceName & """    ;"
'                                    strReadText = "OK"
                                 End If
                           End If
                                  
                     Case "AUDIO"
                         'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
                     Case "N-MOSFET"
                         'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
                     Case "P-MOSFET"
                        'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
                     Case "MOS"
                        'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
  
                     Case "IR"
                       'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
                     Case "RESO"
                         'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                     Case "DUAL"
                         'PIN LIB
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                      strReadText = "OK"
                                   End If
                     
                     Case "XTAL"
                         'conn  OSC TO CONN
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
'                         'PIN LIB
'                                   If bListPinLib = True Then
'                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
'                                      strReadText = "OK"
'                                   End If
                     
                     Case "DIODE"
                     
                         tmpSTR1 = Trim(Replace(tmpSTR1, TmpStr(0), ""))
                         tmpSTR1 = Trim(tmpSTR1)
                         TmpStr = Split(tmpSTR1, " ")
                         DeviceType_A = Trim(TmpStr(0))
                          If Len(DeviceType_A) > 1 Then
                             Select Case UCase(DeviceType_A)
                               Case "RB521S-30"
                               'DIODE
                                         If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                         End If

                               Case "V.R"
                               'DIODE
                                  If UCase(TmpStr(1)) = "BZX384-C15" Then
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                                  End If
                               Case "BAT54A"
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                               Case "ARRAY"
                               'pin lib
                                             If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                               Case "ARR"
                                 'PIN LIB
                                      'IP4220CZ6
                                     
                                     If UCase(TmpStr(1)) = "IP4220CZ6" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                               
                               Case "ESD"
                                     'PRTR5V0U2X
                                     'PIN LIB
                                     If UCase(TmpStr(1)) = "PRTR5V0U2X" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                     'PESD5V0S2BT
                                     'PIN LIB
                                     If UCase(TmpStr(1)) = "PESD5V0S2BT" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                     'IP4223CZ6
                                     'PIN LIB
                                     If UCase(TmpStr(1)) = "IP4223CZ6" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                               
                               
                               
                               
                               Case "SCHOTTKY"
                               'PIN LIB
                                    If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                               Case "SDMG0340LC-7-F"
                               'PIN LIB
                                    If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                               Case "CH715FPT"
                               'PIN LIB
                                    If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                               Case "SW"
                               'PIN LIB
                                    If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                               Case "FAST"
                               'PIN LIB
                                    If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                               
                               Case "BAT54SPT"
                               'PIN LIB
                                    If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                               Case "CH751H-40PT"
                               'DIODE
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                               Case "SDMK0340L-7-F"
                               'DIODE
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                               Case "CH521S-30"
                                'DIODE
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                               Case "SSM5818SLPT"
                                'DIODE
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                               Case "ZEN"
                               
                               'DIODE
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                               Case "SB"
                                  'diode
                                    'B340LA-13-F
                                      If UCase(TmpStr(1)) = "B340LA-13-F" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                               Case "TVS"
                               'DIODE
                                    If bListDiode = True Then
                                        Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & "DIODE" & """;"; Tab(100); "!" & strPN_Str
                                        strReadText = "OK"
                                    End If
                               
                               
                               Case "S.B."
                                   'BAT54CM
                                     If UCase(TmpStr(1)) = "BAT54CM" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                    'SS3020-HE
                                     If UCase(TmpStr(1)) = "SS3020-HE" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                               
                               
                               
                               Case "S.B"
                                'diode to lib
                                   'BAT54A-7-F
                                       If UCase(TmpStr(1)) = "BAT54A-7-F" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   'BAS40-05W
                                      If UCase(TmpStr(1)) = "BAS40-05W" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   

                                   
                                   
                                   'RB715F
                                     If UCase(TmpStr(1)) = "RB715F" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                               
                               
                                   
                                   'BAT54SW
                                     If UCase(TmpStr(1)) = "BAT54SW" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   'BAT54CW
                                     If UCase(TmpStr(1)) = "BAT54CW" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   'BAT54-7-F
                                     If UCase(TmpStr(1)) = "BAT54-7-F" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   
                                   'CH731UPT
                                     If UCase(TmpStr(1)) = "CH731UPT" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   'BAT54
                                     If UCase(TmpStr(1)) = "BAT54" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   'BAT54CW-7-F
                                     If UCase(TmpStr(1)) = "BAT54CW-7-F" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                               
                               
                                   'SBR10U45SP5-13
                                     If UCase(TmpStr(1)) = "SBR10U45SP5-13" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   'BAT54C-7-F
                                     If UCase(TmpStr(1)) = "BAT54C-7-F" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                     End If
                                   
                                   
                                   

                                   
                                   'BAT54CPT
                                   '
                                   
                                   If UCase(TmpStr(1)) = "BAT54CPT" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                   End If
                                   
                                   'RB551V-30
                                   If UCase(TmpStr(1)) = "RB551V-30" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                   End If
                                   
                                   
                                   'BAS40CW
                                   If UCase(TmpStr(1)) = "BAS40CW" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                   End If
                                    'BAT54C
                                   If UCase(TmpStr(1)) = "BAT54C" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                   End If
                                   'BAT54S
                                   If UCase(TmpStr(1)) = "BAT54S" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                   End If
                                   'BAT54S-7-F
                                   If UCase(TmpStr(1)) = "BAT54S-7-F" Then
                                            If bListPinLib = True Then
                                                Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                                strReadText = "OK"
                                            End If
                                   End If

                                   
                                   
                                   

                                 'diode
                                       'RB551V30
                                      If UCase(TmpStr(1)) = "RB551V30" Then
                                         If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                         End If
                                       End If
                                 
                                      'CH551H-30PT
                                        If UCase(TmpStr(1)) = "CH551H-30PT" Then
                                         If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                         End If
                                       End If
                                     
                                     'SMD34JGP
                                        If UCase(TmpStr(1)) = "SMD34JGP" Then
                                         If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                         End If
                                       End If
                                    'SBR3U40P1-7
                                       If UCase(TmpStr(1)) = "SBR3U40P1-7" Then
                                         If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                         End If
                                       End If
                                    
                                    'B140-13-F
                                    
                                      If UCase(TmpStr(1)) = "B140-13-F" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                    'RB751S40T1G
                                       If UCase(TmpStr(1)) = "RB751S40T1G" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                    
                                    'MBR130T1G
                                       If UCase(TmpStr(1)) = "MBR130T1G" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                    
                                    
                                    
                                    'B240A-13
                                      If UCase(TmpStr(1)) = "B240A-13" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                
                                    'RB500V-40TE-17
                                      If UCase(TmpStr(1)) = "RB500V-40TE-17" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                    
                                    
                                   'CH520S-30PT
                                     If UCase(TmpStr(1)) = "CH520S-30PT" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                     'SMD24JGP
                                     If UCase(TmpStr(1)) = "SMD24JGP" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                     
                                    'SSM0160SGP
                                    If UCase(TmpStr(1)) = "SSM0160SGP" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                    'SS0520
                                    If UCase(TmpStr(1)) = "SS0520" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                    End If
                                    'B0530WS-7-F
                                    If UCase(TmpStr(1)) = "B0530WS-7-F" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                                    
                                    '1PS76SB21
                                   If UCase(TmpStr(1)) = "1PS76SB21" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                                    
                                    
                                    
                                   'CH035H-40PT
                                   If UCase(TmpStr(1)) = "CH035H-40PT" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                                
                                   'CH751H-40PT
                                   If UCase(TmpStr(1)) = "CH751H-40PT" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                                   'RB520CS-30
                                    If UCase(TmpStr(1)) = "RB520CS-30" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                                   
                                   
                                   'SD103AWS
                                   If UCase(TmpStr(1)) = "SD103AWS" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If
                                  '
                                   'RB751V-40
                                   If UCase(TmpStr(1)) = "RB751V-40" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                   End If

                                   
                             End Select
                             
                             Else
                             
                          End If
                     
                     
                         
'                       If bListDiode = True Then
'                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
'                           strReadText = "OK"
'                       End If
                     
                     Case "DIODES"
'                       If bListDiode = True Then
'                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
'                           strReadText = "OK"
'                       End If
                     Case "ESD"
                     'PIN LIB
                            If Trim(TmpStr(1)) = "PROTECTION" Then
                                   If bListPinLib = True Then
                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                       strReadText = "OK"
                                   End If
                            End If
                     Case "LED"
'                     'PIN LIB
'                                  If UCase(Trim(TmpStr(1))) <> "WHI" And UCase(Trim(TmpStr(1))) <> "WHITE" Then
'                                        If bListPinLib = True Then
'                                            Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
'                                            strReadText = "OK"
'                                        End If
'                                   Else
'                                       If bListDiode = True Then
'                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
'                                           strReadText = "OK"
'                                       End If
'                                  End If
                                     If strDeviceNomber = "83.00326.G70" Or strDeviceNomber = "83.01221.R70" Then
                                        If bListPinLib = True Then
                                            Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                                            strReadText = "OK"
                                        End If
                                     End If
                     
                     
                     Case "XTOR"
                     
                                    If UCase(TmpStr(2)) = "B0520LW-7-F" Then
                                       If bListDiode = True Then
                                           Print #8, strDeviceName; Tab(25); strDH; Tab(35); strDL; Tab(45); "PN""" & strDeviceName & """;"; Tab(100); "!" & strPN_Str
                                           strReadText = "OK"
                                       End If
                                       
                                       
                                       Else
                                       
                                        If bListPinLib = True Then
                                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
    
                                          strReadText = "OK"
    
                                       End If
                                   End If
'                      'PIN LIB
'                                   If bListPinLib = True Then
'                                       Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
'
'                                      strReadText = "OK"
'
'                                   End If
                     Case "MICROPHONE"
                      'PIN LIB
                      'SPM0423HD4H-WB
                      If bListPinLib = True Then
                        If UCase(TmpStr(1)) = "SPM0423HD4H-WB" Then
                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                           strReadText = "OK"
                        End If
                      End If
                     Case "LVDS"
                     'PIN LIB
                      If bListPinLib = True Then
                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                           strReadText = "OK"
                      
                      End If
                     Case "NFC"
                     'PIN LIB
                      If bListPinLib = True Then
                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                           strReadText = "OK"
                      
                      End If
                     Case "WLAN"
                     'PIN LIB
                      If bListPinLib = True Then
                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                           strReadText = "OK"
                      
                      End If
                     Case "FET"
                     'PIN LIB
                      If bListPinLib = True Then
                           Print #9, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """;"; Tab(100); "!" & strPN_Str
                           strReadText = "OK"
                      
                      End If
                     Case "STANDOFF"

                      'conn NT  H1,H2
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                      Case "NUT-SMT"
                      'conn NT  H1,H2
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                      Case "BOSSNUT"
                      'conn NT  H1,H2
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "NUT"
                     
                      'conn NT  H1,H2
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "NT;"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     
                     Case "BRKT"
                     'conn  RCT
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
                     Case "HLD"

                     'conn  RCT
                            If bListConnect = True Then
                               Print #24, strDeviceName; Tab(25); "PN""" & strDeviceNomber & """" & ";"; Tab(100); "!" & strPN_Str
                               strReadText = "OK"
                            End If
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
  If bListConnect = True Then
      Close #24
  End If
   'MsgBox "OK" & Chr(13) & Chr(10) & "File save path:" & PrmPath & "ReadBomValue\", vbInformation
  
Exit Sub
EX:
MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me

End Sub

Private Sub mcGo_Click()
On Error Resume Next
  Erase strBom_All_Device
  intBom_All_Number = 0
  
  
   If Trim(txtBomPath.Text) = "" Then txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
    If Dir(txtBomPath.Text) = "" Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtBomPath.SetFocus
      Exit Sub
    End If
    If FileLen(txtBomPath.Text) = 0 Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "The file text is null ,please check!", vbCritical
        txtBomPath.SetFocus
        Exit Sub
    End If
     If cUndefined.Value = 1 Then
     
       If Dir(txtCadPartFile.Text) = "" Then
           txtCadPartFile.Text = " Please open *.asc file!(DblClick me open file!)"
           MsgBox "File not find!", vbCritical
           txtCadPartFile.SetFocus
        Exit Sub
       End If
            If FileLen(txtCadPartFile.Text) = 0 Then
                txtCadPartFile.Text = " Please open *.asc file!(DblClick me open file!)"
                MsgBox "The file text is null ,please check!", vbCritical
                txtCadPartFile.SetFocus
                Exit Sub
            End If
     
     
     End If
     Frame1.Enabled = False
 
    Frame2.Enabled = False
    Frame3.Enabled = False
    mcGo.Enabled = False
    txtBomPath.Enabled = False
    'start
       If CheckJumper.Value = 1 Then
          If Trim(txtJumper.Text) = "" Then
             txtJumper.Text = 11
          End If
          a = Val(txtJumper.Text)
          txtJumper.Text = a
       End If
       
      Call Read_BomFile
       If cUndefined.Value = 1 Then
           Call CreatDevicePart_File
       End If
      Call File_HE_Bing
      MsgBox "OK" & Chr(13) & Chr(10) & "File save path:" & PrmPath & "ReadBomValue\", vbInformation
    'end
    mcGo.Enabled = True
    txtBomPath.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    mcGo.SetFocus
    Exit Sub
EX:
mcGo.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
txtBomPath.Enabled = True
 mcGo.SetFocus
 MsgBox Err.Description, vbCritical
End Sub

Private Sub txtBomPath_2_DblClick()
On Error Resume Next
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    
    .Filter = "bom file *.txt|*.txt|*.*|*.*"
    .ShowOpen

End With
    txtBomPath_2.Text = Me.CommonDialog1.FileName
    If Dir(txtBomPath_2.Text) = "" Then
        txtBomPath_2.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtBomPath_2.SetFocus
      Exit Sub
    End If
    If FileLen(txtBomPath_2.Text) = 0 Then
        txtBomPath_2.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "The file text is null ,please check!", vbCritical
        txtBomPath_2.SetFocus
        Exit Sub
    End If
 
Exit Sub

errh:
MsgBox Err.Description, vbCritical
    txtBomPath_2.Text = "Please open bom file!(DblClick me open file!)"
    txtBomPath_2.SetFocus
End Sub

Private Sub txtBomPath_DblClick()
On Error Resume Next
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    
    .Filter = "bom file *.txt|*.txt|*.*|*.*"
    .ShowOpen

End With
    txtBomPath.Text = Me.CommonDialog1.FileName
    If Dir(txtBomPath.Text) = "" Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtBomPath.SetFocus
      Exit Sub
    End If
    If FileLen(txtBomPath.Text) = 0 Then
        txtBomPath.Text = " Please open bom file!(DblClick me open file!)"
        MsgBox "The file text is null ,please check!", vbCritical
        txtBomPath.SetFocus
        Exit Sub
    End If
 
Exit Sub

errh:
MsgBox Err.Description, vbCritical
    txtBomPath.Text = "Please open bom file!(DblClick me open file!)"
    txtBomPath.SetFocus


End Sub
Private Sub File_HE_Bing()
    On Error Resume Next
   Dim strFenPei() As String
   Dim strFenPei_1() As String
   Dim Mystr As String
   Dim TmpStr_1 As String
   Dim strMyStr As String
  ' Dim PinLibOk As Boolean
       Open PrmPath & "ReadBomValue\unFind_PinLib.txt" For Output As #29
     Open PrmPath & "ReadBomValue\Bom_To_Board.txt" For Output As #26
     
   'board head
   If cBoardHead.Value = 1 Then
      Print #26, txtBoardHead.Text
      Print #26,
   End If
     
     'Capacitor
   If CheckC.Value = 1 Then
      Open PrmPath & "ReadBomValue\Capacitor.txt" For Input As #27
         Do Until EOF(27)
           Line Input #27, Mystr
           Print #26, Mystr
         Loop
      Close #27
        Print #26,
       
   End If
   'connect
   If CheckCn.Value = 1 Then
      Open PrmPath & "ReadBomValue\Connector.txt" For Input As #27
         Do Until EOF(27)
           Line Input #27, Mystr
           'ADD testjet
           'If InStr(Mystr, "  NT;") = 0 Then
             strFenPei_1 = Split(Mystr, "!")
          ' End If
          If Trim(Mystr) <> "" And Trim(UCase(Mystr)) <> "CONNECTOR" Then
             'strFenPei_1 = Split(Mystr, """")
             TmpStr_1 = Mystr
               
                     Do
                       strlB = Replace(TmpStr_1, "   ", "  ")
                         If TmpStr_1 = strlB Then Exit Do
                          TmpStr_1 = strlB
                     Loop
             strFenPei = Split(TmpStr_1, "  ")
                 strMyStr = strFenPei(1)
                strFenPei(1) = Replace(strFenPei(1), "PN" & """", "")
                strFenPei(1) = Replace(strFenPei(1), """" & ";", "")
                strFenPei(1) = Replace(strFenPei(1), """", "")
          

'
            
                
            Select Case Trim(UCase(strFenPei(0)))
                 Case "BATT1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                 Case "CRTBD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "TPLED1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "MEDIA1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "LEDBD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "DM1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "BAT_CON1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "FP1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "DM2"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "PWRBT2"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "PWRBT1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CPU1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "HDD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "WLAN1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "DB1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "FAN1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "LCD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "IOBD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "KB1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "ODD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "IOBD2"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "TPAD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "DTB1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "WWAN1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CN4"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "TNAN1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                 Case "IOSBD1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "SPK1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CN1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CON1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CON2"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CON3"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CON4"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "CON5"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "ITP1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case "NEW1"
                    strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                    bConnectFindOk = True
                Case Else
                    
                        
                            'part number
                        Select Case Trim(UCase(strFenPei(1)))
                             Case "12016-00070600"
                                strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                                Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                                bConnectFindOk = True
                              Case "12018-00020300"
                                strFenPei(1) = "PN" & """" & "CONN" & """" & " TJ;"
                                Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1)
                                bConnectFindOk = True
                                
                        End Select
                        
                          
                                Call Connect_Tj(Trim(UCase(strFenPei(1))), strFenPei(0), strFenPei_1(1))
                         
                    
                    
                 If bConnectFindOk = False Then
                    Print #26, Mystr
                    
                 End If
            End Select
             Else
                    Print #26, Mystr
          End If
          bConnectFindOk = False
          ' Print #26, Mystr
               Erase strFenPei
               Erase strFenPei_1
               Mystr = ""
               TmpStr_1 = ""
               
         Loop
      Close #27
        Print #26,
        
   End If
      'diode
   If CheckD.Value = 1 Then
      Open PrmPath & "ReadBomValue\Diode.txt" For Input As #27
         Do Until EOF(27)
           Line Input #27, Mystr
           Print #26, Mystr
         Loop
      Close #27
      Print #26,
   End If
   If CheckJumper.Value = 1 Or CheckIND.Value = 1 Then
   
      'Jumper
       Open PrmPath & "ReadBomValue\Jumper.txt" For Input As #27
         Do Until EOF(27)
           Line Input #27, Mystr
           Print #26, Mystr
         Loop
      Close #27
      Print #26,
   End If
      'Pin Library.txt
   If Checklb.Value = 1 Then
       Open PrmPath & "ReadBomValue\Pin Library.txt" For Input As #27
         Do Until EOF(27)
           Line Input #27, Mystr
             strFenPei_1 = Split(Mystr, "!")
          ' End If
          If Trim(Mystr) <> "" And Trim(UCase(Mystr)) <> "PIN LIBRARY" Then
             'strFenPei_1 = Split(Mystr, """")
             TmpStr_1 = Mystr
               
                     Do
                       strlB = Replace(TmpStr_1, "   ", "  ")
                         If TmpStr_1 = strlB Then Exit Do
                          TmpStr_1 = strlB
                     Loop
             strFenPei = Split(TmpStr_1, "  ")
                 strMyStr = strFenPei(1)
                strFenPei(1) = Replace(strFenPei(1), ";", "")
                strFenPei(1) = Replace(strFenPei(1), "PN" & """", "")
                strFenPei(1) = Replace(strFenPei(1), """" & ";", "")
                strFenPei(1) = Replace(strFenPei(1), """", "")
                tmpPartNumber = Trim(UCase(strFenPei(1)))
            Select Case Trim(UCase(strFenPei(1))) '1
                Case "79.47719.9BL"
                    strFenPei(1) = "PN" & """" & "3pin_cap_470u" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10136.04L"
                    strFenPei(1) = "PN" & """" & "100_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10236.04L"
                    strFenPei(1) = "PN" & """" & "1K_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "09M21"
                    strFenPei(1) = "PN" & """" & "Madison" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "71.0HM75.00U"
                    strFenPei(1) = "PN" & """" & "PCH" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "47DM1"
                    strFenPei(1) = "PN" & """" & "PCH" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "KCJVN"
                    strFenPei(1) = "PN" & """" & "VGA" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "1RW8N"
                    strFenPei(1) = "PN" & """" & "PCH" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "HDV1R"
                    strFenPei(1) = "PN" & """" & "CPU" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "16005149"
                    strFenPei(1) = "PN" & """" & "16005149" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                Case "WHVFN"
                    strFenPei(1) = "PN" & """" & "CPU" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                 Case "WJ4Y2"
                 
                    strFenPei(1) = "PN" & """" & "CPU" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                Case "1JDF3"
                    strFenPei(1) = "PN" & """" & "PCH" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "418F3"
                    strFenPei(1) = "PN" & """" & "CPU" & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10036.08L"
                    strFenPei(1) = "PN" & """" & "10_8P4R" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10136.A8L"
                    strFenPei(1) = "PN" & """" & "100_8p4r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10156.04L"
                    strFenPei(1) = "PN" & """" & "100_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "82.20037.111"
                    strFenPei(1) = "PN" & """" & "100Mhz_4p" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "72.05224.00U"
                    strFenPei(1) = "PN" & """" & strFenPei(1) & """" & " TJ;"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "83.BAT54.Q81"
                    strFenPei(1) = "PN" & """" & "BAT54C" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "83.00054.Y81"
                    strFenPei(1) = "PN" & """" & "BAT54C" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "83.BAT54.K81"
                    strFenPei(1) = "PN" & """" & "BAT54C" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "83.BAT54.N81"
                    strFenPei(1) = "PN" & """" & "BAT54C" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "83.BAT54.P81"
                    strFenPei(1) = "PN" & """" & "BAT54C" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.R003A.04L"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "0_4p2r_open_13" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "0_4p2r" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "74.08236.013"
                    strFenPei(1) = "PN" & """" & "UPC8236T6N" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                Case "82.30005.891"
                    strFenPei(1) = "PN" & """" & "14mhz_2p" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "82.30004.791"
                    strFenPei(1) = "PN" & """" & "24mhz" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "82.30023.821"
                    strFenPei(1) = "PN" & """" & "25mhz" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "84.02305.G31"
                    strFenPei(1) = "PN" & """" & "2n7002" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "84.03541.F31"
                    strFenPei(1) = "PN" & """" & "2n7002" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                    
                 Case "84.3K329.031"
                    strFenPei(1) = "PN" & """" & "2n7002" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                
                
                Case "07G005C69010"
                    strFenPei(1) = "PN" & """" & "2n7002" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                Case "84.04435.U37"
                    strFenPei(1) = "PN" & """" & "2n7002_5p" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "74.62122.093"
                    strFenPei(1) = "PN" & """" & "TPS62122DRVR" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "84.27002.W31"
                    strFenPei(1) = "PN" & """" & "2n7002" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "84.S0610.B31"
                    strFenPei(1) = "PN" & """" & "2n7002" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "84.27002.D3F"
                    strFenPei(1) = "PN" & """" & "2n7002dw" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "84.04935.C37"
                    strFenPei(1) = "PN" & """" & "2n7002dw" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                    
                Case "66.4733A.04L"
                    strFenPei(1) = "PN" & """" & "47k_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.15036.08L"
                    strFenPei(1) = "PN" & """" & "8P_4r_15" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10336.04L"
                    strFenPei(1) = "PN" & """" & "10K_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                Case "66.10336.04L"
                    strFenPei(1) = "PN" & """" & "10K_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                    
                Case "66.1033A.04L"
                    strFenPei(1) = "PN" & """" & "10K_4p2r" & """" & ";"
                    strFenPei(1) = UCase(strFenPei(1))
                    Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                    PinLibOk = True
                   
                   
                   
'                Case Else
'
'                    Print #26, Mystr
            End Select
            '=================not testjet
            Select Case Trim(UCase(strFenPei(1))) '2
                Case "84.03904.I11"
                     strFenPei(1) = "PN" & """" & "3904" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.19223.A70"
                     strFenPei(1) = "PN" & """" & "DIODE_4P_2D" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.R2002.B8E"
                     strFenPei(1) = "PN" & """" & "DIODE_6P_3D" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.09091.F3F"
                     strFenPei(1) = "PN" & """" & "G9091" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.09091.G3F"
                     strFenPei(1) = "PN" & """" & "G9091" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.00195.G70"
                     strFenPei(1) = "PN" & """" & "LED_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.01426.037"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00035.037"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                     
                Case "84.00510.03D"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "84.00309.037"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "84.01712.037"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.04134.037"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.04168.037"
                     strFenPei(1) = "PN" & """" & "MOS8P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03904.T11"
                     strFenPei(1) = "PN" & """" & "TRAN_NPN_BEC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.02043.011"
                     strFenPei(1) = "PN" & """" & "TRAN_NPN_BEC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                     
                     
                Case "84.03904.L06"
                     strFenPei(1) = "PN" & """" & "TRAN_NPN_BEC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.M3904.B11"
                     strFenPei(1) = "PN" & """" & "TRAN_NPN_BEC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.02222.V11"
                     strFenPei(1) = "PN" & """" & "TRAN_NPN_BEC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.T3906.A11"
                     strFenPei(1) = "PN" & """" & "TRAN_PNP_EBC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00069.B1B"
                     strFenPei(1) = "PN" & """" & "TRAN_PNP_EBC" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00144.P11"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "84.05124.011"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.05124.A11"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.DT144.A11"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                
                Case "84.02143.011"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.05144.011"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                
                Case "84.00143.N11"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                
                Case "84.00044.B1K"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00143.M11"
                     strFenPei(1) = "PN" & """" & "tran_r_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00124.K1K"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00124.H1K"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00144.I11"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00124.T1K"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00143.D1K"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00143.F1K"
                     strFenPei(1) = "PN" & """" & "TRAN_R_12" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "71.SEYMR.M01"
                     strFenPei(1) = "PN" & """" & "VGA" & """" & " TJ" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                 
                Case "68.68160.30B"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform12p_open_8l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform12p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                   
                Case "68.IH189.301"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                   
                Case "68.IH106.30C"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                   
                Case "68.HD131.301"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.IH601.301"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.NS14Y.301"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.89246.301"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                   
                   
                 Case "68.68167.30B"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform16p_open_8l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform16p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                  Case "68.0NS14.30B"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform16p_open_8l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform16p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                   
                   
                   
                  Case "68.HD081.30B"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform16p_open_8l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform16p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.01284.30B"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.69241.30C"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.HH085.301"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.89241.30A"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "68.IH115.30A"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "xform24p_open_16l" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "xform24p" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "66.R0036.04L"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "0_4p2r_open_13" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "0_4p2r" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
            
            
            End Select
            
            Select Case Trim(UCase(strFenPei(1))) '3
                Case "66.R0036.A4L"
                   If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                             strFenPei(1) = "PN" & """" & "0_4p2r_open_13" & """" & ";"
                            strFenPei(1) = UCase(strFenPei(1))
                            Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                            PinLibOk = True
                   
                      Else
                        strFenPei(1) = "PN" & """" & "0_4p2r" & """" & ";"
                        strFenPei(1) = UCase(strFenPei(1))
                        Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                        PinLibOk = True

                   End If
                Case "66.R0036.08L"
                  If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                     strFenPei(1) = "PN" & """" & "0_8p4r_open_18" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                  
                    Else
                     strFenPei(1) = "PN" & """" & "0_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                 End If
                Case "66.R0036.A8L"
                  If UCase(Trim(strFenPei_1(1))) = UCase("Bom in Cad not find device.") Then
                     strFenPei(1) = "PN" & """" & "0_8p4r_open_18" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                  
                    Else
                     strFenPei(1) = "PN" & """" & "0_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                 End If
                Case "66.12236.080"
                     strFenPei(1) = "PN" & """" & "1.2K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.15236.04L"
                     strFenPei(1) = "PN" & """" & "1.5K_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.15236.08L"
                     strFenPei(1) = "PN" & """" & "1.5K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.18236.04L"
                     strFenPei(1) = "PN" & """" & "1.8K_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.10436.04L"
                     strFenPei(1) = "PN" & """" & "100K_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.10436.A8L"
                     strFenPei(1) = "PN" & """" & "100K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.10436.L0L"
                     strFenPei(1) = "PN" & """" & "100K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "66.1043A.08L"
                     strFenPei(1) = "PN" & """" & "100K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.10336.10L"
                     strFenPei(1) = "PN" & """" & "10K_10P8R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.10336.L0L"
                     strFenPei(1) = "PN" & """" & "10K_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.10336.08L"
                     strFenPei(1) = "PN" & """" & "10K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.1033A.08L"
                     strFenPei(1) = "PN" & """" & "10K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "66.10336.A8L"
                     strFenPei(1) = "PN" & """" & "10K_8P4R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.20040.131"
                     strFenPei(1) = "PN" & """" & "12.88MHZ_13" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.931"
                     strFenPei(1) = "PN" & """" & "12.88MHZ_13" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30006.161"
                     strFenPei(1) = "PN" & """" & "12M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30006.271"
                     strFenPei(1) = "PN" & """" & "12M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30006.381"
                     strFenPei(1) = "PN" & """" & "12M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30034.651"
                     strFenPei(1) = "PN" & """" & "12M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                            
            
            
            End Select
            
            Select Case Trim(UCase(strFenPei(1))) '4
                    Case "82.30005.941"
                         strFenPei(1) = "PN" & """" & "14.31818M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.901"
                         strFenPei(1) = "PN" & """" & "14.318M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.951"
                         strFenPei(1) = "PN" & """" & "14.318M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.A51"
                         strFenPei(1) = "PN" & """" & "14.318M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.A61"
                         strFenPei(1) = "PN" & """" & "14.318M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.B81"
                         strFenPei(1) = "PN" & """" & "14.318M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.C51"
                         strFenPei(1) = "PN" & """" & "14.318M" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.20001.181"
                         strFenPei(1) = "PN" & """" & "14.318M_34" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "82.30005.E11"
                         strFenPei(1) = "PN" & """" & "14.318M_34" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.15156.08L"
                         strFenPei(1) = "PN" & """" & "150_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "69.10110.341"
                         strFenPei(1) = "PN" & """" & "1580MHZ" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "09G061041350"
                         strFenPei(1) = "PN" & """" & "1580MHZ" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.15336.04L"
                         strFenPei(1) = "PN" & """" & "15K_4P2R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.15336.08L"
                         strFenPei(1) = "PN" & """" & "15K_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.10236.08L"
                         strFenPei(1) = "PN" & """" & "1K_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.10236.A8L"
                         strFenPei(1) = "PN" & """" & "1K_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.1023A.08L"
                         strFenPei(1) = "PN" & """" & "1K_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.22236.04L"
                         strFenPei(1) = "PN" & """" & "2.2K_4P2R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                     Case "66.2223A.04L"
                         strFenPei(1) = "PN" & """" & "2.2K_4P2R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                         
                         
                    Case "66.22236.08L"
                         strFenPei(1) = "PN" & """" & "2.2K_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
                    Case "66.22236.A8L"
                         strFenPei(1) = "PN" & """" & "2.2K_8P4R" & """" & ";"
                         strFenPei(1) = UCase(strFenPei(1))
                         Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                         PinLibOk = True
            
            End Select
            
            Select Case Trim(UCase(strFenPei(1))) '5
                Case "66.20336.04L"
                     strFenPei(1) = "PN" & """" & "20K_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.2033A.04A"
                     strFenPei(1) = "PN" & """" & "20K_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.22036.04L"
                     strFenPei(1) = "PN" & """" & "22_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.2203A.04L"
                     strFenPei(1) = "PN" & """" & "22_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.22136.04L"
                     strFenPei(1) = "PN" & """" & "220_4P2R" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30023.511"
                     strFenPei(1) = "PN" & """" & "24.576M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30023.601"
                     strFenPei(1) = "PN" & """" & "24.576M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30023.611"
                     strFenPei(1) = "PN" & """" & "24.576M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30023.651"
                     strFenPei(1) = "PN" & """" & "24.576M" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30004.131"
                     strFenPei(1) = "PN" & """" & "24MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30004.761"
                     strFenPei(1) = "PN" & """" & "24MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30023.391"
                     strFenPei(1) = "PN" & """" & "24MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.06208.073"
                     strFenPei(1) = "PN" & """" & "isl6208crz" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "82.30023.701"
                     strFenPei(1) = "PN" & """" & "24MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.561"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.571"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.601"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.761"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.791"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.851"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.971"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                            
            
            End Select
            Select Case Trim(UCase(strFenPei(1))) '6
                Case "82.30020.A31"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.A41"
                     strFenPei(1) = "PN" & """" & "25MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30004.641"
                     strFenPei(1) = "PN" & """" & "25MHZ_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30004.831"
                     strFenPei(1) = "PN" & """" & "25MHZ_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.D41"
                     strFenPei(1) = "PN" & """" & "25MHZ_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.G61"
                     strFenPei(1) = "PN" & """" & "25MHZ_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30020.G71"
                     strFenPei(1) = "PN" & """" & "25MHZ_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30067.271"
                     strFenPei(1) = "PN" & """" & "26MHZ_4P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.20056.441"
                     strFenPei(1) = "PN" & """" & "26MHZ_6P" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30034.641"
                     strFenPei(1) = "PN" & """" & "27MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30034.681"
                     strFenPei(1) = "PN" & """" & "27MHZ" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30034.701"
                     strFenPei(1) = "PN" & """" & "27MHZ_13" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00084.F31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                 Case "84.21607.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                Case "84.00102.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00138.F31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00138.G31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00138.H31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00601.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00610.C31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.02130.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
            
            End Select
            
            Select Case Trim(UCase(strFenPei(1))) '7
                Case "84.02301.G31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03203.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03400.B37"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03401.B31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03403.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03404.B31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03404.C31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03413.A31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03418.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03419.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03434.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03K15.A31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.05067.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.07002.I31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.07401.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.27002.E3F"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.27002.I31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.27002.L04"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.27002.N31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
            
            End Select
            
            Select Case Trim(UCase(strFenPei(1))) '8
                Case "84.2N702.A31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.D31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.E31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.I31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.J31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.W31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.P8503.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00304.B31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.04501.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.44P02.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.02305.A31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.02301.N31"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.C3K"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.04151.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03409.031"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "07005-00520000"
                     strFenPei(1) = "PN" & """" & "pmdpb" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "07G005886010"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "07G005D51010"
                     strFenPei(1) = "PN" & """" & "2N7002" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00301.A31"
                     strFenPei(1) = "PN" & """" & "2N7002_TWO_DIODE" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.04813.B37"
                     strFenPei(1) = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.08200.037"
                     strFenPei(1) = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei(1) = UCase(strFenPei(1))
                     Print #26, strFenPei(0); Tab(25); strFenPei(1); Tab(100); "!" & strFenPei_1(1) & "  " & tmpPartNumber
                     PinLibOk = True
            
            
            End Select
            

            
            
              If PinLibOk = False Then
                Call Pin_Lib_List_Not_Tj_1(Trim(UCase(strFenPei(1))), strFenPei(0), strFenPei_1(1))
            
              End If
              'TR
              If PinLibOk = False Then
                   Call Pin_Lib_List_Not_Tj_2(Trim(UCase(strFenPei(1))), strFenPei(0), strFenPei_1(1))
              End If
            '===========================

            ' testjet add====================
 
              If PinLibOk = False Then
                Call Pin_Lib_List_Tj_1(Trim(UCase(strFenPei(1))), strFenPei(0), strFenPei_1(1))
            
              End If
 
            '================================
            'TESTJET LIB'================================
 
            
              If PinLibOk = False Then
                   Call Pin_Lib_List_Tj_And_Lib_1(Trim(UCase(strFenPei(1))), strFenPei(0), strFenPei_1(1))
              End If
            
            
            
            '================================
            'wait lib  Pin_Lib_List_WaitLib_Tj_1
              If PinLibOk = False Then
                   Call Pin_Lib_List_WaitLib_Tj_1(Trim(UCase(strFenPei(1))), strFenPei(0), strFenPei_1(1))
              End If
            
            
            '==============================
                If PinLibOk = False Then
                    Print #26, Mystr
                    Print #29, Mystr
                End If
           
            Else
                    Print #26, Mystr
          End If
          ' Print #26, Mystr
               Erase strFenPei
               Erase strFenPei_1
               PinLibOk = False
               Mystr = ""
               TmpStr_1 = ""
             tmpPartNumber = ""
         Loop
      Close #27
      Print #26,
   End If
      'Resistor
   If CheckR.Value = 1 Then
       Open PrmPath & "ReadBomValue\Resistor.txt" For Input As #27
         Do Until EOF(27)
           Line Input #27, Mystr
           Print #26, Mystr
         Loop
      Close #27
   End If
      
     Close #29
     Close #26
End Sub
Private Sub Pin_Lib_List_Not_Tj_2(strPartNumber As String, strFenPei_0 As String, strFenPei_1_1 As String)
        '
              'TRXXX
       Select Case strPartNumber
                 Case "69.10103.061"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                 Case "69.10084.101"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.00201.141"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.00201.201"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.03216.20B"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.10110.10G"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10084.081"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10087.011"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.01012.201"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "10G302000004030"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "10G3025R1004030"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10087.041"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                   

                   
                   
                Case "68.02012.20C"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                 Case "68.02002.021"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                   
                   
                Case "68.02002.011"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                   
                   
                Case "68.01210.201"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.02012.201"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.11900.201"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "68.CTX17.101"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10080.021"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                 Case "69.10166.001"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                   
                   
                Case "69.10084.071"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10098.031"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10103.041"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_12" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "69.10118.001"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
                Case "ZZ.0R04P.ZZZ"
                   If UCase(Trim(strFenPei_1_1)) = UCase("Bom in Cad not find device.") Then
                        strFenPei_1 = "PN" & """" & "4P2l_OPEN14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                     Else
                        strFenPei_1 = "PN" & """" & "4P2L_14" & """" & ";"
                        strFenPei_1 = UCase(strFenPei_1)
                        Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                        PinLibOk = True
                   End If
        End Select
End Sub
Private Sub Pin_Lib_List_Not_Tj_1(strPartNumber As String, strFenPei_0 As String, strFenPei_1_1 As String)


Select Case strPartNumber '1 200 line
                Case "02038-00010000"
                      strFenPei_1 = "PN" & """" & "02038-00010000" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "54.03346.F21"
                      strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                 Case "71.32151.A0U"
                      strFenPei_1 = "PN" & """" & "71.32151.A0U" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                  Case "73.02102.023"
                      strFenPei_1 = "PN" & """" & "73.02102.023" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                   Case "74.53318.073"
                      strFenPei_1 = "PN" & """" & "74.53318.073" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                    Case "74.61181.A73"
                      strFenPei_1 = "PN" & """" & "74.61181.A73" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                     
                     
                 Case "71.97324.00W"
                      strFenPei_1 = "PN" & """" & "71.97324.00W" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "71.CLOVE.D0U"
                      strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                 Case "71.60802.A0U"
                      strFenPei_1 = "PN" & """" & "71.60802.A0U" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "71.00691.00G"
                      strFenPei_1 = "PN" & """" & "KBC" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "71.03261.003"
                      strFenPei_1 = "PN" & """" & "71.03261.003" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "71.01211.A0U"
                      strFenPei_1 = "PN" & """" & "71.01211.A0U" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                     
                     
                     
                     
                Case "02135-00080100"
                      strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "02135-00120000"
                      strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "02G145001600"
                      strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "03100-00021500"
                      strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "03100-00090000"
                      strFenPei_1 = "PN" & """" & "03100-00090000" & """" & " TJ;"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                
                
                
                
                Case "84.27002.B3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.27002.C3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.27002.F3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.A3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.F3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.780SN.A3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.DM601.03F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.DMN66.03F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00138.03F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.UM6K1.A3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.2N702.E3F"
                     strFenPei_1 = "PN" & """" & "2N7002DW" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.33236.08L"
                     strFenPei_1 = "PN" & """" & "3.3K_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.30136.04L"
                     strFenPei_1 = "PN" & """" & "300_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.10026.021"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.661"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.691"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.841"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.861"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.961"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.A41"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.A81"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.B21"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.B81"
                     strFenPei_1 = "PN" & """" & "32.768K" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.C01"
                     strFenPei_1 = "PN" & """" & "32.768K_12" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.C21"
                     strFenPei_1 = "PN" & """" & "32.768K_12" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.931"
                     strFenPei_1 = "PN" & """" & "32.768K_12" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30001.D51"
                     strFenPei_1 = "PN" & """" & "32.768K_12" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.33036.04L"
                     strFenPei_1 = "PN" & """" & "33_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.33036.08L"
                     strFenPei_1 = "PN" & """" & "33_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.33036.A8L"
                     strFenPei_1 = "PN" & """" & "33_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30091.001"
                     strFenPei_1 = "PN" & """" & "37.4MHZ" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03904.H11"
                     strFenPei_1 = "PN" & """" & "3904" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03904.P11"
                     strFenPei_1 = "PN" & """" & "3904" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03904.Q11"
                     strFenPei_1 = "PN" & """" & "3904" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03906.G11"
                     strFenPei_1 = "PN" & """" & "3906" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03906.H11"
                     strFenPei_1 = "PN" & """" & "3906" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03906.N11"
                     strFenPei_1 = "PN" & """" & "3906" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03906.P11"
                     strFenPei_1 = "PN" & """" & "3906" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03906.R11"
                     strFenPei_1 = "PN" & """" & "3906" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "79.33719.2EL"
                     strFenPei_1 = "PN" & """" & "3P_CAP" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "77.24771.15L"
                     strFenPei_1 = "PN" & """" & "3P_CAP_470U" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.47236.04L"
                     strFenPei_1 = "PN" & """" & "4.7K_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.4723A.04L"
                     strFenPei_1 = "PN" & """" & "4.7K_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "66.47236.08L"
                     strFenPei_1 = "PN" & """" & "4.7K_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.47236.A8L"
                     strFenPei_1 = "PN" & """" & "4.7K_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.47036.08L"
                     strFenPei_1 = "PN" & """" & "47_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.47336.08L"
                     strFenPei_1 = "PN" & """" & "47_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.4713A.04A"
                     strFenPei_1 = "PN" & """" & "470_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.47136.08L"
                     strFenPei_1 = "PN" & """" & "470_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.47136.A8L"
                     strFenPei_1 = "PN" & """" & "470_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30026.231"
                     strFenPei_1 = "PN" & """" & "48MHZ" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "82.30026.271"
                     strFenPei_1 = "PN" & """" & "48MHZ_4P" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.49R96.04L"
                     strFenPei_1 = "PN" & """" & "49.9_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.49R56.08L"
                     strFenPei_1 = "PN" & """" & "49.9_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "68.2415N.101"
                     strFenPei_1 = "PN" & """" & "4P_2L_5CLOSE" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.5R136.04L"
                     strFenPei_1 = "PN" & """" & "4P_2R_5.1O_12_34" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.51036.04L"
                     strFenPei_1 = "PN" & """" & "4P_2R_51" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True

                Case "66.51036.A8L"
                     strFenPei_1 = "PN" & """" & "51_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.56036.04L"
                     strFenPei_1 = "PN" & """" & "56_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.56035.08L"
                     strFenPei_1 = "PN" & """" & "56_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.56036.08L"
                     strFenPei_1 = "PN" & """" & "56_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.00545.073"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.07534.C79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.00547.079"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.00547.A79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.00547.C79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.02181.079"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.02182.071"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.02191.079"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.07534.079"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.07534.A79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.09715.079"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.09715.B79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.02311.079"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.00547.G79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.07534.D79"
                     strFenPei_1 = "PN" & """" & "5VUSB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.45N03.A30"
                     strFenPei_1 = "PN" & """" & "6_PIN_MOS" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.68036.04L"
                     strFenPei_1 = "PN" & """" & "68_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.68136.08L"
                     strFenPei_1 = "PN" & """" & "680_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.00005.BAE"
                     strFenPei_1 = "PN" & """" & "6P_10DIODE_2A_5C" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.17N03.030"
                     strFenPei_1 = "PN" & """" & "7_PIN_MOS" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.07414.BAB"
                     strFenPei_1 = "PN" & """" & "74AHC14" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.07414.L20"
                     strFenPei_1 = "PN" & """" & "74AHC14" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True


                Case "73.74125.L12"
                     strFenPei_1 = "PN" & """" & "74AHCT125" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.01G08.ABH"
                     strFenPei_1 = "PN" & """" & "74AHCT1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03131.031"
                     strFenPei_1 = "PN" & """" & "2n7002" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00280.03F"
                     strFenPei_1 = "PN" & """" & "2n7002dw" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.01034.C3F"
                     strFenPei_1 = "PN" & """" & "2n7002dw" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                Case "74.06288.07F"
                     strFenPei_1 = "PN" & """" & "check_5pin_o1_s4_a3" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.06288.A7F"
                     strFenPei_1 = "PN" & """" & "check_5pin_o1_s4_a3" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.09193.A33"
                     strFenPei_1 = "PN" & """" & "check_7pin_s1_o4_a6" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.00413.03F"
                     strFenPei_1 = "PN" & """" & "2n7002_7pin" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.03660.037"
                     strFenPei_1 = "PN" & """" & "fdms3600_10p" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "73.01G32.BAH"
                     strFenPei_1 = "PN" & """" & "74AHCT1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                 Case "84.08402.03F"
                     strFenPei_1 = "PN" & """" & "2n7002dw" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                 Case "73.01T45.0HH"
                     strFenPei_1 = "PN" & """" & "sn74lvc1t" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "73.01G08.CJH"
                     strFenPei_1 = "PN" & """" & "74AHCT1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.01G08.L04"
                     strFenPei_1 = "PN" & """" & "74AHCT1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.1G125.0JH"
                     strFenPei_1 = "PN" & """" & "74AHCT1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.03306.D0B"
                     strFenPei_1 = "PN" & """" & "74CBTD3306" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.03306.E0B"
                     strFenPei_1 = "PN" & """" & "74CBTD3306" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.07408.L15"
                     strFenPei_1 = "PN" & """" & "74LVC08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.07408.L16"
                     strFenPei_1 = "PN" & """" & "74LVC08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.01G08.DHG"
                     strFenPei_1 = "PN" & """" & "74LVC1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "73.01G02.AHG"
                     strFenPei_1 = "PN" & """" & "74LVC1G08" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                Case "73.01G32.AHH"
                     strFenPei_1 = "PN" & """" & "74LVC1G32" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.7SZ32.CAH"
                     strFenPei_1 = "PN" & """" & "74LVC1G32" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.7503A.04B"
                     strFenPei_1 = "PN" & """" & "75_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.03157.C0H"
                     strFenPei_1 = "PN" & """" & "7SB3157" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "73.74157.CHH"
                     strFenPei_1 = "PN" & """" & "7SB3157" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.82238.10L"
                     strFenPei_1 = "PN" & """" & "8.2K_10P8R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.82236.04L"
                     strFenPei_1 = "PN" & """" & "8.2K_4P2R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.82236.08L"
                     strFenPei_1 = "PN" & """" & "8.2K_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.82236.A8L"
                     strFenPei_1 = "PN" & """" & "8.2K_8P4R" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.01065.0AJ"
                     strFenPei_1 = "PN" & """" & "8P_13DIODE_78" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.33136.A8L"
                     strFenPei_1 = "PN" & """" & "8P4R_330" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.75036.08L"
                     strFenPei_1 = "PN" & """" & "8P4R_75" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "66.75036.A8L"
                     strFenPei_1 = "PN" & """" & "8P4R_75" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.04030.01B"
                     strFenPei_1 = "PN" & """" & "9435" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.05350.A1B"
                     strFenPei_1 = "PN" & """" & "9435" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "84.09435.03B"
                     strFenPei_1 = "PN" & """" & "9435" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.00900.079"
                     strFenPei_1 = "PN" & """" & "A00900AIDCNR" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.03010.0B3"
                     strFenPei_1 = "PN" & """" & "AL3010_SETUP" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "06G087026010"
                     strFenPei_1 = "PN" & """" & "AL3010_SETUP" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                     
                Case "74.02101.079"
                     strFenPei_1 = "PN" & """" & "AP2101MPG" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.02010.013"
                     strFenPei_1 = "PN" & """" & "APA2010_SETUP" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.51463.043"
                     strFenPei_1 = "PN" & """" & "tps51463" & """" & " TJ LIB" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.05930.03D"
                     strFenPei_1 = "PN" & """" & "APL5930" & """" & " TJ LIB" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.G9731.03D"
                     strFenPei_1 = "PN" & """" & "APL5930" & """" & " TJ LIB" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "74.07078.043"
                     strFenPei_1 = "PN" & """" & "APW7078QBI_SETUP" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "72.45011.A01"
                     strFenPei_1 = "PN" & """" & "AT45DB011B" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.00016.B11"
                     strFenPei_1 = "PN" & """" & "BAS16" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "07G004393020"
                     strFenPei_1 = "PN" & """" & "BAS16" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                     
                Case "83.00016.K11"
                     strFenPei_1 = "PN" & """" & "BAS16" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.00016.F11"
                     strFenPei_1 = "PN" & """" & "BAS40" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.00040.E81"
                     strFenPei_1 = "PN" & """" & "BAS40" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.00040.M81"
                     strFenPei_1 = "PN" & """" & "BAS40" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True
                Case "83.R2004.C81"
                     strFenPei_1 = "PN" & """" & "BAS40" & """" & ";"
                     strFenPei_1 = UCase(strFenPei_1)
                     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
                     PinLibOk = True


End Select

Select Case strPartNumber '2 500 line
Case "83.00054.S81"
     strFenPei_1 = "PN" & """" & "BAT54" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00054.Z81"
     strFenPei_1 = "PN" & """" & "BAT54" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.D81"
     strFenPei_1 = "PN" & """" & "BAT54" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00016.G11"
     strFenPei_1 = "PN" & """" & "BAT54A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.0SM24.A11"
     strFenPei_1 = "PN" & """" & "BAT54A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "83.00054.R81"
     strFenPei_1 = "PN" & """" & "BAT54A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.U81"
     strFenPei_1 = "PN" & """" & "BAT54A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.V01"
     strFenPei_1 = "PN" & """" & "BAT54A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.R2003.P81"
     strFenPei_1 = "PN" & """" & "BAT54A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00054.I81"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00054.Q81"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00054.X81"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00056.K11"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.081"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.A81"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.Y81"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.R2003.E81"
     strFenPei_1 = "PN" & """" & "BAT54C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00054.M81"
     strFenPei_1 = "PN" & """" & "BAT54CW" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07G004069020"
     strFenPei_1 = "PN" & """" & "BAT54CW" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "83.01222.X80"
     strFenPei_1 = "PN" & """" & "BAT54CW" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.B81"
     strFenPei_1 = "PN" & """" & "BAT54CW" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00054.B8K"
     strFenPei_1 = "PN" & """" & "BAT54S" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.0R203.081"
     strFenPei_1 = "PN" & """" & "BAT54S" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.C81"
     strFenPei_1 = "PN" & """" & "BAT54S" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00099.K11"
     strFenPei_1 = "PN" & """" & "BAV99" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.5R003.08F"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAV99.H11"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00099.M11"
     strFenPei_1 = "PN" & """" & "BAV99" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00099.P11"
     strFenPei_1 = "PN" & """" & "BAV99" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00099.T11"
     strFenPei_1 = "PN" & """" & "BAV99" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAV99.D11"
     strFenPei_1 = "PN" & """" & "BAV99" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00056.E11"
     strFenPei_1 = "PN" & """" & "BAW56" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00056.G11"
     strFenPei_1 = "PN" & """" & "BAW56" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00056.I11"
     strFenPei_1 = "PN" & """" & "BAW56" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00056.J11"
     strFenPei_1 = "PN" & """" & "BAW56" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03906.F11"
     strFenPei_1 = "PN" & """" & "CH3906" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.R0304.B81"
     strFenPei_1 = "PN" & """" & "CH715" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02700.093"
     strFenPei_1 = "PN" & """" & "CHECK_10_PIN_8PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09711.D7F"
     strFenPei_1 = "PN" & """" & "CHECK_1PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G051025010"
     strFenPei_1 = "PN" & """" & "CHECK_3PIN_O1" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09293.B43"
     strFenPei_1 = "PN" & """" & "CHECK_3PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G030159010"
     strFenPei_1 = "PN" & """" & "CHECK_4PIN_O1_S4" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.80933.BBF"
     strFenPei_1 = "PN" & """" & "CHECK_4PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.70033.A3F"
     strFenPei_1 = "PN" & """" & "CHECK_5PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09013.N7F"
     strFenPei_1 = "PN" & """" & "check_5pin_o5_s3" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09193.033"
     strFenPei_1 = "PN" & """" & "check_7pin_o4_s1" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.70018.03F"
     strFenPei_1 = "PN" & """" & "CHECK_5PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09013.D7F"
     strFenPei_1 = "PN" & """" & "CHECK_5PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08818.B3F"
     strFenPei_1 = "PN" & """" & "CHECK_5PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06007-00060000"
     strFenPei_1 = "PN" & """" & "CHECK_5PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06016-00110000"
     strFenPei_1 = "PN" & """" & "CHECK_6PIN_O6_S4" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05930.073"
     strFenPei_1 = "PN" & """" & "CHECK_6PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08010.D43"
     strFenPei_1 = "PN" & """" & "CHECK_6PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06007-00190000"
     strFenPei_1 = "PN" & """" & "CHECK_6PIN_S6_O3" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06016-00070000"
     strFenPei_1 = "PN" & """" & "CHECK_6PIN_S6_O3" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06018-00030000"
     strFenPei_1 = "PN" & """" & "CHECK_7PIN_O2_S4" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G070047010"
     strFenPei_1 = "PN" & """" & "CHECK_7PIN_O5" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05712.0BB"
     strFenPei_1 = "PN" & """" & "CHECK-2P_ALL3PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01803.07B"
     strFenPei_1 = "PN" & """" & "CHECK-2P_ALL3PIN_PWR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09091.I3F"
     strFenPei_1 = "PN" & """" & "CHECKPOWER" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.A1117.F3C"
     strFenPei_1 = "PN" & """" & "CHECKPOWER" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00268.07B"
     strFenPei_1 = "PN" & """" & "CI_TIE" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.03217.0BZ"
     strFenPei_1 = "PN" & """" & "CM3217A3OG_DIODE" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.58865.037"
     strFenPei_1 = "PN" & """" & "CSD58865" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00124.N1K"
     strFenPei_1 = "PN" & """" & "DDTA124EUA" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00115.G1K"
     strFenPei_1 = "PN" & """" & "DDTC115EUA" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00115.H11"
     strFenPei_1 = "PN" & """" & "DDTC115EUA" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.P6SBM.AAG"
     strFenPei_1 = "PN" & """" & "DIODE" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.01221.R70"
     strFenPei_1 = "PN" & """" & "DIODE_3P_1D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "83.01105.070"
     strFenPei_1 = "PN" & """" & "DIODE_3P_1D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "83.00056.Q11"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.10R04.F87"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.3X101.011"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.BAT54.Z81"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07004-00030000"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07004-00030300"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.12222.070"
     strFenPei_1 = "PN" & """" & "DIODE_3P_2D_12_13" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.10004.A8M"
     strFenPei_1 = "PN" & """" & "DIODE_3P_NC2" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.19223.B70"
     strFenPei_1 = "PN" & """" & "DIODE_4P_2D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00195.N70"
     strFenPei_1 = "PN" & """" & "DIODE_4P_2D_13_24" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.R0304.A8H"
     strFenPei_1 = "PN" & """" & "DIODE_6P_3D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.R3004.A8E"
     strFenPei_1 = "PN" & """" & "DIODE_6P_3D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.02304.0AE"
     strFenPei_1 = "PN" & """" & "DIODE_6P_9D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "83.42236.0AE"
     strFenPei_1 = "PN" & """" & "DIODE_6P_9D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "07G028075010"
     strFenPei_1 = "PN" & """" & "DIODE_6P_9D" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02618.07B"
     strFenPei_1 = "PN" & """" & "EC2618NLB1GR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05711.07B"
     strFenPei_1 = "PN" & """" & "EM6781" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06781.07B"
     strFenPei_1 = "PN" & """" & "EM6781" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09132.B7B"
     strFenPei_1 = "PN" & """" & "EM6781" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09132.C7B"
     strFenPei_1 = "PN" & """" & "EM6781" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06331.A3D"
     strFenPei_1 = "PN" & """" & "FDC6331L" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "84.07200.A37"
     strFenPei_1 = "PN" & """" & "FDMS3600_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "84.05524.037"
     strFenPei_1 = "PN" & """" & "FDMS3600_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

 Case "84.06920.037"
     strFenPei_1 = "PN" & """" & "FDMS3600_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00033.037"
     strFenPei_1 = "PN" & """" & "FDMS3600_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07200.037"
     strFenPei_1 = "PN" & """" & "FDMS3600_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03664.037"
     strFenPei_1 = "PN" & """" & "FDMS3600_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.09620.030"
     strFenPei_1 = "PN" & """" & "FDMS9620" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.02301.D31"
     strFenPei_1 = "PN" & """" & "FDN340P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.0340P.A31"
     strFenPei_1 = "PN" & """" & "FDN340P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00358.A31"
     strFenPei_1 = "PN" & """" & "FDN358P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01412.AE3"
     strFenPei_1 = "PN" & """" & "G14212RC1U_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05285.07F"
     strFenPei_1 = "PN" & """" & "G5285T" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.R9712.A71"
     strFenPei_1 = "PN" & """" & "G546A1P1UF" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00546.07D"
     strFenPei_1 = "PN" & """" & "G546B2P1U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05461.D71"
     strFenPei_1 = "PN" & """" & "G546B2P1U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05930.07P"
     strFenPei_1 = "PN" & """" & "G5930" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00690.I7B"
     strFenPei_1 = "PN" & """" & "G690" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00709.073"
     strFenPei_1 = "PN" & """" & "G709RCUF_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00709.A7F"
     strFenPei_1 = "PN" & """" & "G709T1UF" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09091.H3F"
     strFenPei_1 = "PN" & """" & "G9091" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09198.G7F"
     strFenPei_1 = "PN" & """" & "G9091" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00909.03F"
     strFenPei_1 = "PN" & """" & "G913C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True



End Select

Select Case strPartNumber '3 500 line
Case "74.00923.C3F"
     strFenPei_1 = "PN" & """" & "G913C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00991.031"
     strFenPei_1 = "PN" & """" & "G991P11U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05606.A71"
     strFenPei_1 = "PN" & """" & "G991P11U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.04220.0AE"
     strFenPei_1 = "PN" & """" & "IP4220" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.01222.K70"
     strFenPei_1 = "PN" & """" & "LED_12_13_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07G015701358"
     strFenPei_1 = "PN" & """" & "LED_12_13_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00110.J70"
     strFenPei_1 = "PN" & """" & "LED_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.21355.A70"
     strFenPei_1 = "PN" & """" & "LED_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "83.00110.R70"
     strFenPei_1 = "PN" & """" & "LED_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00326.G70"
     strFenPei_1 = "PN" & """" & "LED_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00327.D70"
     strFenPei_1 = "PN" & """" & "LED_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.01108.070"
     strFenPei_1 = "PN" & """" & "LED_3P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.00195.I70"
     strFenPei_1 = "PN" & """" & "LED_4P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00358.K11"
     strFenPei_1 = "PN" & """" & "LM358" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00358.Q11"
     strFenPei_1 = "PN" & """" & "LM358" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00393.I21"
     strFenPei_1 = "PN" & """" & "LM393ADR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00393.M21"
     strFenPei_1 = "PN" & """" & "LM393ADR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00331.A2F"
     strFenPei_1 = "PN" & """" & "LMV331" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.14566.093"
     strFenPei_1 = "PN" & """" & "MAX14566EETA" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.35200.01D"
     strFenPei_1 = "PN" & """" & "MBT35200" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "07G005607011"
     strFenPei_1 = "PN" & """" & "MOS_6P_2D_2C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "07005-00530000"
     strFenPei_1 = "PN" & """" & "MOS_6P_2D_2C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06725.030"
     strFenPei_1 = "PN" & """" & "MOS_7P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00166.037"
     strFenPei_1 = "PN" & """" & "MOS_8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07G005503130"
     strFenPei_1 = "PN" & """" & "MOS5P_1D_1C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07G005C78010"
     strFenPei_1 = "PN" & """" & "MOS5P_1D_1C" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00655.B3D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07G005107010"
     strFenPei_1 = "PN" & """" & "2mos6pin" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00658.B3D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03456.C3D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03457.A37"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06402.B3D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06721.030"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.P2703.03D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.02423.03D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04402.03D"
     strFenPei_1 = "PN" & """" & "MOS6P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00036.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06675.030"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07696.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08878.A30"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

Case "84.003K3.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00034.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00164.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00308.B30"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00351.036"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00365.036"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.003B9.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.003B9.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.003C9.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00402.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00406.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.02659.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03604.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03604.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04178.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04406.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04407.G37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04433.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04456.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04483.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04634.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04686.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04714.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04718.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04720.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04800.D37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04812.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04835.D37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04894.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06414.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06675.D37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06676.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06679.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06690.E37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06790.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07634.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07636.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07670.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07672.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07686.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07692.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07698.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08028.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08030.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08037.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08040.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08065.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08672.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08692.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08707.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08880.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08884.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08896.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08A03.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.12003.A37"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.17313.03D"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08342.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.09N03.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.30043.037"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03133.031"
     strFenPei_1 = "PN" & """" & "2n7002" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "5.4IPZZ.020"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "65.4IPZZ.020"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "65.4IPZZ.030"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "65.4IPZZ.028"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "65.4IPZZ.032"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "65.4IPZZ.026"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07G005D52010"
     strFenPei_1 = "PN" & """" & "MOS8P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25010.K01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "05006-00080100"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "06007-00370000"
     strFenPei_1 = "PN" & """" & "check_5pin_o1_s3" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "72.25Q64.F01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.93C66.A01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
End Select

Select Case strPartNumber '4 755 line
Case "72.25016.A01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25020.00D"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25020.C01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25020.E01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25021.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25032.D01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25106.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25160.B01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25165.A01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25200.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25205.B01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25320.C01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25321.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25325.A01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25512.F01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25512.J01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25P32.C01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25Q16.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25Q32.A01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25Q64.C01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25Q80.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25X16.A01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25X80.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.26161.001"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.45021.E01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.45081.B01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.45081.C01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25Q64.D01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25Q64.B01"
     strFenPei_1 = "PN" & """" & "MX25_READID" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.25X40.B03"
     strFenPei_1 = "PN" & """" & "MX25_READID_10P" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.7S125.09H"
     strFenPei_1 = "PN" & """" & "NC7SP125P5X" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01008.A73"
     strFenPei_1 = "PN" & """" & "NCT1008CMT3R2G_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00072.0B3"
     strFenPei_1 = "PN" & """" & "NCT1008CMT3R2G_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.03521.07F"
     strFenPei_1 = "PN" & """" & "NCT3521U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G030150010"
     strFenPei_1 = "PN" & """" & "NCT3521U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

Case "06016-00080000"
     strFenPei_1 = "PN" & """" & "NCT3521U" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G045160010"
     strFenPei_1 = "PN" & """" & tmpPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.03940.A71"
     strFenPei_1 = "PN" & """" & "NCT3940S" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "69.10120.041"
     strFenPei_1 = "PN" & """" & "NFA21SL307X1A45L" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "09G061062100"
     strFenPei_1 = "PN" & """" & "NFA21SL307X1A45L" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02793.A31"
     strFenPei_1 = "PN" & """" & "P2793A" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02800.A71"
     strFenPei_1 = "PN" & """" & "P2800EA1" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G107017010"
     strFenPei_1 = "PN" & """" & "PCA9306DCUR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09306.00A"
     strFenPei_1 = "PN" & """" & "PCA9306DCUR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G107017011"
     strFenPei_1 = "PN" & """" & "PCA9306DCUR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "83.10004.08M"
     strFenPei_1 = "PN" & """" & "PDS1040" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "68.5R610.10S"
     strFenPei_1 = "PN" & """" & "PL28" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.53257.B0C"
     strFenPei_1 = "PN" & """" & "PL5C3257QE" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "83.5V0U2.0A3"
     strFenPei_1 = "PN" & """" & "PRTR5V0U2X" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True



     


     
     
     
End Select
End Sub


Private Sub Pin_Lib_List_Tj_1(strPartNumber As String, strFenPei_0 As String, strFenPei_1_1 As String)
' testjet
Select Case strPartNumber
Case "74.09193.K3F"
     strFenPei_1 = "PN" & """" & "PT9193" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G029138024"
     strFenPei_1 = "PN" & """" & "PT9193" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.09013.J7F"
     strFenPei_1 = "PN" & """" & "RT9013" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09022.03P"
     strFenPei_1 = "PN" & """" & "RT9022GE" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06038-00020000"
     strFenPei_1 = "PN" & """" & "RT9022GE" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.09198.07F"
     strFenPei_1 = "PN" & """" & "RT9198" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09724.09F"
     strFenPei_1 = "PN" & """" & "RT9724" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.TCS20.03B"
     strFenPei_1 = "PN" & """" & "SENSOR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03456.D3D"
     strFenPei_1 = "PN" & """" & "SI3456" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.03993.03D"
     strFenPei_1 = "PN" & """" & "SI3993DV" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.06506.A3D"
     strFenPei_1 = "PN" & """" & "SI3993DV" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True




Case "84.00172.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00312.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00412.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

Case "84.00460.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.00462.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04407.F37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04468.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04496.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.04835.H37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07121.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07410.A37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07702.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07716.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08039.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08059.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08061.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.80925.0BF"
     strFenPei_1 = "PN" & """" & "check_5pin_o1" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
     
Case "84.08061.A37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True








Case "84.08062.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08064.A37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08065.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08065.B37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08882.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.08884.037"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.57N03.A37"
     strFenPei_1 = "PN" & """" & "SI4800" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.07686.A37"
     strFenPei_1 = "PN" & """" & "SI7686DP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.01G17.AHH"
     strFenPei_1 = "PN" & """" & "SN74AUC1G" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.03305.A0B"
     strFenPei_1 = "PN" & """" & "SN74CBTD" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True



Case "73.1G175.CHJ"
     strFenPei_1 = "PN" & """" & "SN74LVC1G175DCKR_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "82.40005.141"
     strFenPei_1 = "PN" & """" & "SPM0423HD4H_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "04G160008210"
     strFenPei_1 = "PN" & """" & "SPM0423HD4H_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     

Case "73.7SZ08.EAH"
     strFenPei_1 = "PN" & """" & "TC7SZ08" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.01G09.AAH"
     strFenPei_1 = "PN" & """" & "74vhc1g" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "84.02222.S11"
     strFenPei_1 = "PN" & """" & "tran_npn_ebc" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.62065.041"
     strFenPei_1 = "PN" & """" & "TLV62065DSGR_SETUP" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True



Case "74.70233.03F"
     strFenPei_1 = "PN" & """" & "TLV70233DBVR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06288.079"
     strFenPei_1 = "PN" & """" & "TPS2000CDGNR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06288.A79"
     strFenPei_1 = "PN" & """" & "TPS2000CDGNR" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02301.071"
     strFenPei_1 = "PN" & """" & "TPS2000CDGNR" & """" & ";"
      strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02001.079"
     strFenPei_1 = "PN" & """" & "TPS2000CDGNR" & """" & ";"
      strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02000.B71"
     strFenPei_1 = "PN" & """" & "TPS2000CDGNR" & """" & ";"
      strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.80923.0BF"
     strFenPei_1 = "PN" & """" & "check_5pin_o1" & """" & ";"
      strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00715.01Z"
     strFenPei_1 = "PN" & """" & "check_7pin_o6_s4" & """" & ";"
      strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.01G08.EHG"
     strFenPei_1 = "PN" & """" & "u74lvc1g08g" & """" & ";"
      strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     

Case "71.72020.D0U"
     strFenPei_1 = "PN" & """" & "UPD720200AF1" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

Case "YD3C5$AA"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "JP0F2$AA"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
          
Case "HT31P$AA"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

Case "74.08209.C73"
     strFenPei_1 = "PN" & """" & "+1_05V_PCH_P" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51123.073"
     strFenPei_1 = "PN" & """" & "+5V_PWR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51218.073"
     strFenPei_1 = "PN" & """" & "1.05VPOWER" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08209.B73"
     strFenPei_1 = "PN" & """" & "1.1VPWR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
 'pin ban
Case "74.01008.073"
     strFenPei_1 = "PN" & """" & "74.01008.073" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09297.043"
     strFenPei_1 = "PN" & """" & "74.09297.043" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
  
Case "74.65911.07Z"
     strFenPei_1 = "PN" & """" & "74.65911.07Z" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.03513.093"
     strFenPei_1 = "PN" & """" & "check_9pin_s3_o2" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08009.04P"
     strFenPei_1 = "PN" & """" & "check_6pin_o3_s1" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
     
Case "74.32411.093"
     strFenPei_1 = "PN" & """" & "check_5pin_o1_s4" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.52364.009"
     strFenPei_1 = "PN" & """" & "73.52364.009" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08975.0BZ"
     strFenPei_1 = "PN" & """" & "74.08975.0BZ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "1100302"
     strFenPei_1 = "PN" & """" & "1100302" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "20002494"
     strFenPei_1 = "PN" & """" & "20002494" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "73.02179.00J"
     strFenPei_1 = "NT PN" & """" & "73.02179.00J" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "69.10110.361"
     strFenPei_1 = "NT PN" & """" & "69.10110.361" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00S20.000"
     strFenPei_1 = "PN" & """" & "71.00S20.000" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.02412.B03"
     strFenPei_1 = "PN" & """" & "71.02412.B03" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09306.003"
     strFenPei_1 = "PN" & """" & "71.09306.003" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.42268.00U"
     strFenPei_1 = "PN" & """" & "71.42268.00U" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.32201.01Z"
     strFenPei_1 = "PN" & """" & "74.32201.01Z" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.63823.00U"
     strFenPei_1 = "PN" & """" & "71.63823.00U" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.02402.D01"
     strFenPei_1 = "PN" & """" & "mx25_readid" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09724.093"
     strFenPei_1 = "PN" & """" & "check_9pin_s3_o2" & """" & ";"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "84.86350.037"
     strFenPei_1 = "PN" & """" & "1_1V_VTT" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08207.B73"
     strFenPei_1 = "PN" & """" & "1_5V_SUS" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00105.013"
     strFenPei_1 = "PN" & """" & "ALC105-GR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05631.A03"
     strFenPei_1 = "PN" & """" & "ALC5631Q-VE-GRT_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00306.0B0"
     strFenPei_1 = "PN" & """" & "AMI306_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.92H81.G03"
     strFenPei_1 = "PN" & """" & "AUD" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.92H79.003"
     strFenPei_1 = "PN" & """" & "AUDIO" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.92H87.A03"
     strFenPei_1 = "PN" & """" & "AUDIO" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "TK259"
     strFenPei_1 = "PN" & """" & "BCM5756" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.51G63.B0U"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03100-00010100"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "02152-00010000"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03006-00030900"
     strFenPei_1 = "PN" & """" & "SDRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03006-00030800"
     strFenPei_1 = "PN" & """" & "SDRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03100-00120100"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03100-00120000"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "05G002602011"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "71.04751.A0U"
     strFenPei_1 = "PN" & """" & "BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.24170.073"
     strFenPei_1 = "PN" & """" & "BQ24170_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.24725.073"
     strFenPei_1 = "PN" & """" & "CHARGER_BQ24725" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51461.043"
     strFenPei_1 = "PN" & """" & "CHECK_0D85V_S0" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51427.073"
     strFenPei_1 = "PN" & """" & "CHECK_3D3V_S5" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08239.A73"
     strFenPei_1 = "PN" & """" & "CHECK_3V_5V_S5" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08628.003"
     strFenPei_1 = "PN" & """" & "CLKGEN" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09480.A03"
     strFenPei_1 = "PN" & """" & "CLKGEN" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.93197.003"
     strFenPei_1 = "PN" & """" & "CLKGEN" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SANDY.J1U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "102500205"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "71.00IVY.G0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00IVY.T0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00IVY.S0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0HM77.A0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "71.SANDY.S0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SANDY.T0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "122GM"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SANDY.DDU"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "71.00IVY.A0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.PANTH.00U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "72.41646.00U"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0HM76.A0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0N13P.M04"
     strFenPei_1 = "PN" & """" & "VGA-GPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "72.52G63.C0U"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "72.52G63.C0U"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "72.42164.G0U"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
  Case "72.41K26.00U"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
     
 Case "71.03221.A03"
     strFenPei_1 = "PN" & """" & "AUDIO" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
  Case "71.0N13P.00U"
     strFenPei_1 = "PN" & """" & "VGA-GPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
  Case "71.0N13M.K0U"
     strFenPei_1 = "PN" & """" & "VGA-GPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "KYY3T"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02004-00120000"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00IVY.C3U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.MCIMX.D0U"
     strFenPei_1 = "PN" & """" & "CPU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.04210.003"
     strFenPei_1 = "PN" & """" & "CS4210" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.04213.003"
     strFenPei_1 = "PN" & """" & "CS4213D" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.51G63.C0U"
     strFenPei_1 = "PN" & """" & "DDR3H5TQ1G" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05028.00G"
     strFenPei_1 = "PN" & """" & "ECE5028" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00034.003"
     strFenPei_1 = "PN" & """" & "FM34-NE-395_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00345.0BZ"
     strFenPei_1 = "PN" & """" & "FREE FALL SENSOR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06208.A73"
     strFenPei_1 = "PN" & """" & "ISL6208BCRZ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06260.A73"
     strFenPei_1 = "PN" & """" & "ISL6260" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.95838.033"
     strFenPei_1 = "PN" & """" & "ISL95838HRTZ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.95838.A33"
     strFenPei_1 = "PN" & """" & "ISL95838HRTZ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.EPF21.A0G"
     strFenPei_1 = "PN" & """" & "KBC" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00885.A0G"
     strFenPei_1 = "PN" & """" & "KBC" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00885.C0G"
     strFenPei_1 = "PN" & """" & "KBC" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.KXTF9.0BZ"
     strFenPei_1 = "PN" & """" & "KXTF9_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.13892.G7Z"
     strFenPei_1 = "PN" & """" & "MC13892AJVLR2_BGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.03050.073"
     strFenPei_1 = "PN" & """" & "MPU-3050_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05605.003"
     strFenPei_1 = "PN" & """" & "NCT58605Y_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.07718.0B9"
     strFenPei_1 = "PN" & """" & "NCT7718W_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.95838.B33"
     strFenPei_1 = "PN" & """" & "74.95838.B33" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
     
Case "74.01314.03Z"
     strFenPei_1 = "PN" & """" & "NEED_CHECK_VCCCORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01316.D33"
     strFenPei_1 = "PN" & """" & "NEED_CHECK_VCCCORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00358.03Z"
     strFenPei_1 = "PN" & """" & "NEED_NODE_1D5V_PWR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51216.073"
     strFenPei_1 = "PN" & """" & "NEED_NODE_1D5V_S3" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08208.A73"
     strFenPei_1 = "PN" & """" & "NEED_NODE_VCC_CORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.95831.A73"
     strFenPei_1 = "PN" & """" & "NEED_NODE_VCC_CORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00357.A3Z"
     strFenPei_1 = "PN" & """" & "NEED_NODE_VGA_PWR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00358.A3Z"
     strFenPei_1 = "PN" & """" & "NEED_NODE_VGA_PWR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02800.B71"
     strFenPei_1 = "PN" & """" & "P2800EA0" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.82S35.B0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "D96PV"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0HM67.00U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0HM67.A0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.COUGA.00U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.COUGA.E0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "CC13C"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0HM57.00U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0IBEX.A0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "YFX79"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.PANTH.D0U"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "ND27V"
     strFenPei_1 = "PN" & """" & "PCH" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.34951.003"
     strFenPei_1 = "PN" & """" & "PI3EQ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.03411.D03"
     strFenPei_1 = "PN" & """" & "PI3VDP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.5USB1.003"
     strFenPei_1 = "PN" & """" & "PL5USB14550AZEE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00121.003"
     strFenPei_1 = "PN" & """" & "PS121QFN48GTR_SETUP" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.5C847.00U"
     strFenPei_1 = "PN" & """" & "R5C847" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.5U241.00U"
     strFenPei_1 = "PN" & """" & "R5U241" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.5U242.00U"
     strFenPei_1 = "PN" & """" & "R5U242" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "C995R$AA"
     strFenPei_1 = "PN" & """" & "RAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "D021T$AA"
     strFenPei_1 = "PN" & """" & "RAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08237.073"
     strFenPei_1 = "PN" & """" & "RT8237" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09025.03D"
     strFenPei_1 = "PN" & """" & "RT9025" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08103.00G"
     strFenPei_1 = "PN" & """" & "RTL8103" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08105.B03"
     strFenPei_1 = "PN" & """" & "RTL8105E" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08111.N03"
     strFenPei_1 = "PN" & """" & "RTL8111F" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08103.B03"
     strFenPei_1 = "PN" & """" & "RTL8130" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SB820.M02"
     strFenPei_1 = "PN" & """" & "SB820M" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03006-00030700"
     strFenPei_1 = "PN" & """" & "SDRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08172.00W"
     strFenPei_1 = "PN" & """" & "TCM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02800.071"
     strFenPei_1 = "PN" & """" & "THERMAL_SENSOR_NO_LIB" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.32030.003"
     strFenPei_1 = "PN" & """" & "TLV320" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51219.073"
     strFenPei_1 = "PN" & """" & "TPS51219RTER" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51225.073"
     strFenPei_1 = "PN" & """" & "TPS51225RUKR" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51611.073"
     strFenPei_1 = "PN" & """" & "TPS51611" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.03410.A0G"
     strFenPei_1 = "PN" & """" & "U12_TUSB3410IVFG4" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.B2507.00G"
     strFenPei_1 = "PN" & """" & "U18_USB2507-ADT" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.01362.B0G"
     strFenPei_1 = "PN" & """" & "U27_SII1362ACLU" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.01352.00I"
     strFenPei_1 = "PN" & """" & "U8_UDA1352TS" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01316.033"
     strFenPei_1 = "PN" & """" & "VCC_CORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01317.03Z"
     strFenPei_1 = "PN" & """" & "VCC_CORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08152.B73"
     strFenPei_1 = "PN" & """" & "VCC_CORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06265.B73"
     strFenPei_1 = "PN" & """" & "VCCCORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.62882.A73"
     strFenPei_1 = "PN" & """" & "VCCPOWER" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.62883.A73"
     strFenPei_1 = "PN" & """" & "VCCPOWER" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.62883.073"
     strFenPei_1 = "PN" & """" & "VCORE" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.ROBSO.M01"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "65.4IEZZ.019"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "71.0N13M.D0U"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
 Case "P8F0H"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
 Case "R7M1G"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "72.41164.I0U"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.M96LP.M01"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0PARK.M02"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.MDSON.M02"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "F66V7"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.ROBSO.M02"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "6H87Y"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0N11P.A2U"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0N11P.D0U"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0N12P.B0U"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.THAME.M07"
     strFenPei_1 = "PN" & """" & "VGA" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0N13M.D0U         IC VGA N13M-GS-S-A1 GB2-64 BGA"
     strFenPei_1 = "PN" & """" & "VGATJ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.42164.D0U"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "5F43R$AA"
     strFenPei_1 = "PN" & """" & "VRAM" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01318.A73"
     strFenPei_1 = "PN" & """" & "VT1318MFQX" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01323.07Z"
     strFenPei_1 = "PN" & """" & "VT1323SFCX" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "54.03346.521"
     strFenPei_1 = "PN" & """" & "WLANTJ" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "04G030007170"
     strFenPei_1 = "PN" & """" & "WLAN+BT2.1" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08903.003"
     strFenPei_1 = "PN" & """" & tmpPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02G361001300"
     strFenPei_1 = "PN" & """" & tmpPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01318.B73"
     strFenPei_1 = "PN" & """" & tmpPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True


End Select
End Sub

Private Sub Pin_Lib_List_Tj_And_Lib_1(strPartNumber As String, strFenPei_0 As String, strFenPei_1_1 As String)
Select Case strPartNumber
Case "74.02541.A73"
     strFenPei_1 = "PN" & """" & "tps2541rter" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.16020.003"
     strFenPei_1 = "PN" & """" & "6v40088" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "73.74125.L13"
     strFenPei_1 = "PN" & """" & "74AHCT125" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "73.07474.FHB"
     strFenPei_1 = "PN" & """" & "74APWR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.74366.0A1"
     strFenPei_1 = "PN" & """" & "74HC366D" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.74366.AA1"
     strFenPei_1 = "PN" & """" & "74HC366D" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.03211.033"
     strFenPei_1 = "PN" & """" & "ADP3211" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00345.ABZ"
     strFenPei_1 = "PN" & """" & "ADXL345BCCZ_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09661.07D"
     strFenPei_1 = "PN" & """" & "APL5930KAI" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SANDY.J0U"
     strFenPei_1 = "PN" & """" & "ARD_BGA1288" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SANDY.P0U"
     strFenPei_1 = "PN" & """" & "ARD_BGA1288" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.24C16.Z01"
     strFenPei_1 = "PN" & """" & "AT24C16" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.24C32.Q01"
     strFenPei_1 = "PN" & """" & "AT24C16" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05625.003"
     strFenPei_1 = "PN" & """" & "ALC5625_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.83351.073"
     strFenPei_1 = "PN" & """" & "W83L351YG" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.41164.H0U"
     strFenPei_1 = "PN" & """" & "VRAMK4W1G" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "9TGTN$AA"
     strFenPei_1 = "PN" & """" & "VRAMK4W1G" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "C09DT$AA"
     strFenPei_1 = "PN" & """" & "VRAMK4W1G" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "ZZ.00PAD.R21"
     strFenPei_1 = "PN" & """" & "VRAMK4W1G" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.02514.F03"
     strFenPei_1 = "PN" & """" & "USB2514B_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.03520.003"
     strFenPei_1 = "PN" & """" & "TS3DV520E" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.03221.003"
     strFenPei_1 = "PN" & """" & "TS3USB221RSER" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06040.013"
     strFenPei_1 = "PN" & """" & "TPA6040" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09789.B13"
     strFenPei_1 = "PN" & """" & "TPA6040" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02062.079"
     strFenPei_1 = "PN" & """" & "TPS2062D" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02062.B71"
     strFenPei_1 = "PN" & """" & "TPS2062D" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02231.073"
     strFenPei_1 = "PN" & """" & "TPS2231" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02231.B73"
     strFenPei_1 = "PN" & """" & "TPS2231" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05538.073"
     strFenPei_1 = "PN" & """" & "TPS2231" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09716.073"
     strFenPei_1 = "PN" & """" & "TPS2231" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09026.079"
     strFenPei_1 = "PN" & """" & "TPS51100" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51100.079"
     strFenPei_1 = "PN" & """" & "TPS51100" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51110.B79"
     strFenPei_1 = "PN" & """" & "TPS51100" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08207.073"
     strFenPei_1 = "PN" & """" & "TPS51116" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51116.073"
     strFenPei_1 = "PN" & """" & "TPS51116" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51116.07B"
     strFenPei_1 = "PN" & """" & "TPS51116" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51117.073"
     strFenPei_1 = "PN" & """" & "TPS51117" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51120.073"
     strFenPei_1 = "PN" & """" & "TPS51120" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51125.073"
     strFenPei_1 = "PN" & """" & "TPS51125" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08554.003"
     strFenPei_1 = "PN" & """" & "SL28647BLCT" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09390.A03"
     strFenPei_1 = "PN" & """" & "SL28647BLCT" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.28647.003"
     strFenPei_1 = "PN" & """" & "SL28647BLCT" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.55221.0E3"
     strFenPei_1 = "PN" & """" & "SLG55221" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08595.003"
     strFenPei_1 = "PN" & """" & "SLG8LV595VTR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.93197.B03"
     strFenPei_1 = "PN" & """" & "SLG8LV595VTR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08513.003"
     strFenPei_1 = "PN" & """" & "SLG8SP513VTR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.60898.073"
     strFenPei_1 = "PN" & """" & "SN0608098" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.02G07.02J"
     strFenPei_1 = "PN" & """" & "SN74LVC2G" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.7WZ07.0AJ"
     strFenPei_1 = "PN" & """" & "SN74LVC2G" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.7SZ08.DAH"
     strFenPei_1 = "PN" & """" & "SNLVC1G08" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.02602.A0U"
     strFenPei_1 = "PN" & """" & "SSM2602" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.74125.F0B"
     strFenPei_1 = "PN" & """" & "TC74VHCT125A" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51311.073"
     strFenPei_1 = "PN" & """" & "TPS51311" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.53311.073"
     strFenPei_1 = "PN" & """" & "TPS51311" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.07402.EHB"
     strFenPei_1 = "PN" & """" & "TSLVC02" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "G602F"
     strFenPei_1 = "PN" & """" & "CANTIGA_PART" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "X965D"
     strFenPei_1 = "PN" & """" & "CANTIGA_PART" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.03257.A03"
     strFenPei_1 = "PN" & """" & "CBT3257ABQ" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.24707.073"
     strFenPei_1 = "PN" & """" & "CHARGER_BQ24707" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.24707.A73"
     strFenPei_1 = "PN" & """" & "CHARGER_BQ24707" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.24727.073"
     strFenPei_1 = "PN" & """" & "CHARGER_BQ24707" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "P802H"
     strFenPei_1 = "PN" & """" & "BCM5761E" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.24745.073"
     strFenPei_1 = "PN" & """" & "CHARGER_BQ24745" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.63020.073"
     strFenPei_1 = "PN" & """" & "CHECK_15_PIN_5PIN_PWR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02211.A39"
     strFenPei_1 = "PN" & """" & "CP2211" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "M2Y2R$BB"
     strFenPei_1 = "PN" & """" & "DDR2_60P" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00351.0B3"
     strFenPei_1 = "PN" & """" & "DE351DLTR8" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01423.0B9"
     strFenPei_1 = "PN" & """" & "EMC1423_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02102.A73"
     strFenPei_1 = "PN" & """" & "EMC2102" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.07922.0B3"
     strFenPei_1 = "PN" & """" & "EMC2102" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.E4002.BB3"
     strFenPei_1 = "PN" & """" & "EMC4002" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.04022.AB3"
     strFenPei_1 = "PN" & """" & "EMC4022" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05281.093"
     strFenPei_1 = "PN" & """" & "G5281" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09091.J3F"
     strFenPei_1 = "PN" & """" & "G9091-330T11U" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.18321.B0U"
     strFenPei_1 = "PN" & """" & "GDDR3_UP67" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.41032.C0U"
     strFenPei_1 = "PN" & """" & "GDDR3_UP67" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.41646.Q0U"
     strFenPei_1 = "PN" & """" & "GDDR3_UP67" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "MU628"
     strFenPei_1 = "PN" & """" & "ICH9M_PART" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.92H71.D03"
     strFenPei_1 = "PN" & """" & "X5NLGXB" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.92H81.C03"
     strFenPei_1 = "PN" & """" & "X5NLGXB" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.92H81.E03"
     strFenPei_1 = "PN" & """" & "X5NLGXB" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "54.03346.B21"
     strFenPei_1 = "PN" & """" & "WLAN_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "TMW4F"
     strFenPei_1 = "PN" & """" & "SB820M_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05138.003"
     strFenPei_1 = "PN" & """" & "RTS5138" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05531.A79"
     strFenPei_1 = "PN" & """" & "R5531" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05912.A71"
     strFenPei_1 = "PN" & """" & "PL5912" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05915.031"
     strFenPei_1 = "PN" & """" & "PL5912" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "73.3L500.003"
     strFenPei_1 = "PN" & """" & "PI3L500" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09539.00K"
     strFenPei_1 = "PN" & """" & "PCA9539" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00795.A0G"
     strFenPei_1 = "PN" & """" & "NPCE795P_XOR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00795.00G"
     strFenPei_1 = "PN" & """" & "NPCE795P_XOR" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.RS880.M05"
     strFenPei_1 = "PN" & """" & "NBRS880M_VOHVOL" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "V5VH6"
     strFenPei_1 = "PN" & """" & "NBRS880M_VOHVOL" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0NB9M.A0U"
     strFenPei_1 = "PN" & """" & "NB9MID" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05035.A0G"
     strFenPei_1 = "PN" & """" & "MEC5035" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05035.B0G"
     strFenPei_1 = "PN" & """" & "MEC5035" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06935.033"
     strFenPei_1 = "PN" & """" & "L6935" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.09035.033"
     strFenPei_1 = "PN" & """" & "L6935" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "J310N"
     strFenPei_1 = "PN" & """" & "M92" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.03243.AFG"
     strFenPei_1 = "PN" & """" & "MAX3243" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.04411.A13"
     strFenPei_1 = "PN" & """" & "MAX4411" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.04885.093"
     strFenPei_1 = "PN" & """" & "MAX4885" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08516.039"
     strFenPei_1 = "PN" & """" & "MAX8516" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08717.073"
     strFenPei_1 = "PN" & """" & "MAX8717" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08731.A73"
     strFenPei_1 = "PN" & """" & "MAX8731AETI" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08744.073"
     strFenPei_1 = "PN" & """" & "MAX8744" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06265.C73"
     strFenPei_1 = "PN" & """" & "ISL6265" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06266.073"
     strFenPei_1 = "PN" & """" & "ISL6266" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.06255.07S"
     strFenPei_1 = "PN" & """" & "ISL6255" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08068.A43"
     strFenPei_1 = "PN" & """" & "RT8068AZQWID_SETUP" & """" & " TJ LIB;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True



End Select
End Sub
Private Sub Connect_Tj(strPartNumber As String, strFenPei_0 As String, strFenPei_1_1 As String)
Select Case strPartNumber


    Case "12018-00020300"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12018-00210200"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12022-00012700"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12G091070038"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12G161E00200"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12G161H0020A"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12G171190408"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12G185404504"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "13GA5E10S010-1"
      strFenPei_1 = "NT PN" & """" & "CONN" & """" & ";"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "13GOK0310M090-10"
      strFenPei_1 = "NT PN" & """" & "CONN" & """" & ";"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
     Case "12G310080005"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
     Case "12018-00080900"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "12G171030040"
      strFenPei_1 = "PN" & """" & "CONN" & """" & " TJ;"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "13GOK0310M100-10"
      strFenPei_1 = "NT PN" & """" & "CONN" & """" & ";"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
    Case "06G008744011"
      strFenPei_1 = "NT PN" & """" & "CONN" & """" & ";"
      Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1
      bConnectFindOk = True
      
      
      
'DM1

End Select
End Sub
Private Sub Pin_Lib_List_WaitLib_Tj_1(strPartNumber As String, strFenPei_0 As String, strFenPei_1_1 As String)
Select Case strPartNumber
Case "71.08162.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "0C011-00080000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
     
Case "02001-00150000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02004-00180100"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.01323.A7Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01326.A7Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "71.03930.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.03930.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.06433.00I"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08114.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "07041-00010000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "71.09365.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.20561.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.57780.M04"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.AL272.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.51G63.H0U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00792.A79"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01454.013"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01454.A13"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05240.A7F"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08208.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08223.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51124.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.88731.B73"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "KG.PGE0V.001"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "KI.G5501.002"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "VR.2GB0G.001"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00268.00G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00268.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00271.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00711.B1G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00773.00G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00781.00G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.03360.A0K"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.04322.C0E"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05028.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05045.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05159.00G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05209.00G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05784.M03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.05787.M02"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.07320.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.07320.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.07412.B0U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.07500.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08105.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08111.I03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08122.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08151.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08585.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08763.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09365.00W"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.09635.A0W"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.0ICH8.A0U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.19N18.Q0W"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.20671.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.28541.A03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.72020.E0U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.87381.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.87383.A0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.CNTIG.I3U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.CREST.M03"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.ICH9M.C1U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.P8101.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.SEYMR.M07"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.WHIST.M02"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.24C64.F01"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.52G63.A0U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "72.55162.B0U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00546.A7D"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01316.E33"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01316.F33"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01317.B3Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01412.0E3"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01431.A1G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02103.A73"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05130.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.05916.031"
     strFenPei_1 = "PN" & """" & "check_6pin_o6_s8_all9" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
Case "74.00977.031"
     strFenPei_1 = "PN" & """" & "check_6pin_o6_s8_all9" & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
          
Case "74.07153.A73"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.07951.0B9"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08770.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.08958.00Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.17021.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51117.07G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.51125.A73"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.54318.043"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.88731.A73"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "KC.52001.UMB"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "KI.80101.031"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "KI.BZM01.LM1"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "KI.G4501.002"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.07583.00U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.00795.B0G"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.47511.00U"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03006-00030600"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.62361.04Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.65912.07Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02001-00130000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02004-00180000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02032-00020000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02043-00050000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02057-00020000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02143-00010000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02G105000300"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02G144000300"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02G561020900"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "02038-00010100"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
     
     
Case "02G561020901"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "03101-00020500"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "05G00100F151"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06016-00180000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06018-00020000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06018-00040000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06018-00130000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06023-00030100"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06046-00100000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06063-00030000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06063-00050000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06085-00040000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G008744011"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G012038020"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G023118020"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G041067010"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G042032010"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "06G108039012"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "09301-00170000"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.08162.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "71.64088.003"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00385.03Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.00386.03Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.01326.07Z"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True
Case "74.02541.073"
     strFenPei_1 = "PN" & """" & strPartNumber & """" & " TJ;"
     strFenPei_1 = UCase(strFenPei_1)
     Print #26, strFenPei_0; Tab(25); strFenPei_1; Tab(100); "!" & strFenPei_1_1 & "  " & tmpPartNumber
     PinLibOk = True

End Select
End Sub


Private Sub txtCadPartFile_DblClick()
On Error Resume Next
With Me.CommonDialog1
    .CancelError = True
    '.Filter = "*.txt|*.txt|*.log|*.log|*.*|*.*"
    
    .Filter = "asc file *.asc|*.asc|*.*|*.*"
    .ShowOpen

End With
    txtCadPartFile.Text = Me.CommonDialog1.FileName
    If Dir(txtCadPartFile.Text) = "" Then
        txtCadPartFile.Text = " Please open *.asc file!(DblClick me open file!)"
        MsgBox "File not find!", vbCritical
        txtCadPartFile.SetFocus
      Exit Sub
    End If
    If FileLen(txtCadPartFile.Text) = 0 Then
        txtCadPartFile.Text = " Please open *.asc file!(DblClick me open file!)"
        MsgBox "The file text is null ,please check!", vbCritical
        txtCadPartFile.SetFocus
        Exit Sub
    End If
 
Exit Sub

errh:
MsgBox Err.Description, vbCritical
    txtCadPartFile.Text = "Please open *.asc file!(DblClick me open file!)"
    txtCadPartFile.SetFocus

End Sub
Private Sub CreatDevicePart_File()
On Error Resume Next
Dim Mystr As String
Dim strFenPei() As String
Dim strFenPei_1() As String
Dim strFenPei_Gan() As String
Dim intI As Integer
Dim bFindOk As Boolean
Dim bCadOk As Boolean
Dim bBomFind_ok As Boolean
Dim tmpPart As String
Dim bPartTrue As Boolean
Dim TmpStr As String
Dim strVal As String
Dim strTol As String
Dim strType As String
Dim strFil As String
  Open PrmPath & "ReadBomValue\Wait_Check_Cad_Device.txt" For Output As #52
  Open PrmPath & "ReadBomValue\Cad_All_Device.txt" For Output As #50
  Open PrmPath & "ReadBomValue\Undefinded_Device.txt" For Output As #51
   Open Trim(txtCadPartFile.Text) For Input As #49
      Do Until EOF(49)
        Line Input #49, Mystr
        Mystr = Trim(UCase(Mystr))

        If Mystr <> "" And InStr(Mystr, "PART=TB_TP") = 0 Then
                'part
                If Left(Mystr, 5) = "PART=" Then
                  ' Mystr = Replace(Mystr, "-", ",")
                   TmpStr = Trim(Replace(Mystr, "PART=", ""))
                   tmpPart = tmpPart & TmpStr
                   strFenPei = Split(tmpPart, ",")
                   bFindOk = True
                   TmpStr = ""
                   intI = intI + 1
                End If
                'FIL
                    If bFindOk = True And Left(Mystr, 4) = "FIL=" Then
                      
                       TmpStr = Trim(Replace(Mystr, "FIL=", ""))
                       strFil = Trim(Replace(TmpStr, """", ""))
                       TmpStr = ""
                    End If
           'TOL
           If bFindOk = True And Left(Mystr, 4) = "TOL=" Then
              
              TmpStr = Trim(Replace(Mystr, "TOL=", ""))
              strTol = Trim(Replace(TmpStr, """", ""))
              TmpStr = ""
           End If
           'val
            If bFindOk = True And Left(Mystr, 4) = "VAL=" Then
              
              TmpStr = Trim(Replace(Mystr, "VAL=", ""))
              strVal = Trim(Replace(TmpStr, """", ""))
              TmpStr = ""
            End If
           'type
            If bFindOk = True And Left(Mystr, 4) = "TYP=" Then
              
              TmpStr = Trim(Replace(Mystr, "TYP=", ""))
              strType = Trim(Replace(TmpStr, """", ""))
              TmpStr = ""
           End If
           ':EOD
            
             If bFindOk = True And Left(Mystr, 4) = ":EOD" Then
                 For i = 0 To UBound(strFenPei)
                      If InStr(Trim(UCase(strFenPei(i))), "-") <> 0 Then
                               strFenPei_Gan = Split(Trim(UCase(strFenPei(i))), "-")
                              ' strFenPei_Gan (0)
                              ' strFenPei_Gan (1)
                                 Dim intFindNum
                                 Dim intFindNumOk_Form
                                 Dim intFindNumOk_To
                                 Dim strCurrDeviceGan As String
                                 Dim strDeviceConnectStr As String
                                 For t = 1 To Len(strFenPei_Gan(0))
                                     
                                    intFindNum = Mid(strFenPei_Gan(0), t, 1)
                                   ' intFindNum = Val(intFindNum)
                                    If Asc(intFindNum) >= 48 And Asc(intFindNum) <= 57 Then
                                          intFindNumOk_Form = intFindNumOk_Form & intFindNum
                                          intFindNumOk_Form = Val(intFindNumOk_Form)
                                       Else
                                       strDeviceConnectStr = Left(strFenPei_Gan(0), t)
                                    End If
                                 Next
                                 intFindNum = ""
                                 For t = 1 To Len(strFenPei_Gan(1))
                                    intFindNum = Mid(strFenPei_Gan(1), t, 1)
                                   ' intFindNum = Val(intFindNum)
                                    If Asc(intFindNum) >= 48 And Asc(intFindNum) <= 57 Then
                                          intFindNumOk_To = intFindNumOk_To & intFindNum
                                          intFindNumOk_To = Val(intFindNumOk_To)
                                     ' Else
                                     '  strDeviceConnectStr = Left(strFenPei_Gan(1), t)
                                    End If
                                 Next
                                 
                                 Do
                                        If Left(strFenPei_Gan(0), Len(strDeviceConnectStr) + 1) = strDeviceConnectStr & "0" Then
                                             strDeviceConnectStr = strDeviceConnectStr & "0"
                                             
                                            Else
                                             Exit Do
                                              
                                        End If
                                 Loop
                                 For h = intFindNumOk_Form To intFindNumOk_To
                                   strCurrDeviceGan = strDeviceConnectStr & h
                                        If Trim(strDeviceName_Gan_Add_All) = "" Then
                                           strDeviceName_Gan_Add_All = strDeviceName_Gan_Add_All & strCurrDeviceGan
                                           Else
                                             strDeviceName_Gan_Add_All = strDeviceName_Gan_Add_All & "," & strCurrDeviceGan
                                        End If

                                 Next
                                    intFindNumOk_To = ""
                                    intFindNum = ""
                                    intFindNumOk_Form = ""
                                     
                               Else
                                        If Trim(strDeviceName_Gan_Add_All) = "" Then
                                           strDeviceName_Gan_Add_All = strDeviceName_Gan_Add_All & Trim(UCase(strFenPei(i)))
                                           Else
                                             strDeviceName_Gan_Add_All = strDeviceName_Gan_Add_All & "," & Trim(UCase(strFenPei(i)))
                                        End If
                            End If
                        
                     Next
          strFenPei = Split(strDeviceName_Gan_Add_All, ",")
                 
                For i = 0 To UBound(strFenPei)
                    
                       If strType = "" Then strType = "UnDeFinde"
                       For j = 0 To UBound(strBom_All_Device) 'intBom_All_Number
                          If Trim(UCase(strFenPei(i))) = Trim(UCase(strBom_All_Device(j))) Then
                             bBomFind_ok = True
                             Exit For
                          End If
                          
                       Next
                       If bBomFind_ok = False Then
                            Print #51, strTol; Tab(30); strType; Tab(55); strFil; Tab(100); strFenPei(i)
                           
                        
                        'Undefinde
                         If UCase(strType) = UCase("UnDefinde") Then
                             'LIB
                              If Left(UCase(strFenPei(i)), 2) = "TR" Then
                               
                         
                                    Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                             End If
                             
                             'res
                              If strTol = "63.75034.L02" Then
                               
                         
                                    Open PrmPath & "ReadBomValue\Resistor.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "8.88M        10   10        f    PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                        bCadOk = True
                                    Close #53
                             End If
                             
                             'LIB
                              If strTol = "72.52G63.A0U" Then
                               
                         
                                    Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                             End If
                             'LIB
                              If strTol = "71.ROBSO.M02" Then
                               
                         
                                    Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                             End If
                             
                             
                             
                             'conn
                             If strTol = "83.01206.0A0" Then
                                   Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                    Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                   bCadOk = True
                                Close #53

                             End If
                             'diode
                             If strTol = "83.01921.P70" Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53

                             End If
                             'conn
                              If strTol = "20.K0320.004" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              'conn
                              If strTol = "34.4EM01.001" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              'conn
                              If strTol = "69.A0017.001" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If

                              
                              'conn
                              If strTol = "20.F1969.006" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                               If strTol = "ZZ.00PAD.Y41" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              
                              
                              
                              'jump
                                If strTol = "63.R0034.1DL" Then
                                   Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                       Print #53, strFenPei(i); Tab(25); "CLOSED PN" & """" & strFil & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                   Close #53
                                End If
                                If strTol = "ZZ.00PAD.M21" Then
                                   Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                       Print #53, strFenPei(i); Tab(25); "CLOSED PN" & """" & strFil & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                   Close #53
                                End If
                                If strTol = "ZZ.00PAD.M11" Then
                                   Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                       Print #53, strFenPei(i); Tab(25); "CLOSED PN" & """" & strFil & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                   Close #53
                                End If
                                
                                If strTol = "ZZ.CLOSE.001" Then
                                   Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                       Print #53, strFenPei(i); Tab(25); "CLOSED PN" & """" & strFil & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                   Close #53
                                End If
                                If strTol = "63.R003C.1IL" Then
                                   Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                       Print #53, strFenPei(i); Tab(25); "CLOSED PN" & """" & strFil & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                   Close #53
                                End If
                               'conn
                              If strTol = "34.4DM11.001" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                               'conn
                              If strTol = "20.K0589.004" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              'conn tj
                              If strTol = "20.F1621.004" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "PN" & """" & "CONN" & """" & " TJ;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              'conn tj
                              If strTol = "20.D0183.110" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "PN" & """" & "CONN" & """" & " TJ;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              
                               'conn fp1
                              If strTol = "20.K0320.006" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              
                         End If
                         'CHOKE
                            If UCase(strType) = "CHOKE" Then
                            ' If strVal = "IND-2D2UH-46-GP-U" Then
                                 If UCase(Left(strFenPei(i), 2)) = "TR" Then
                                        Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                            Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                            bCadOk = True
                                        Close #53
                                             
                                    
                                     Else

                                    Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "OPEN PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                                 End If
                           '  End If
'                                 If strVal = "IND-D47UH-22-GP" Then
'                                       Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
'                                           Print #53, strFenPei(i); Tab(25); "OPEN PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
'                                          bCadOk = True
'                                       Close #53
'                                End If
                            End If
                         'INTEGRATED CIRCUIT
                         If UCase(strType) = "INTEGRATED CIRCUIT" Then
                            Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                bCadOk = True
                            Close #53
                         End If
                         'TRANSISTOR  'pin lib
                         If UCase(strType) = "TRANSISTOR" Then
                            Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                bCadOk = True
                            Close #53
                         End If
                         'JUMP  jump
                         If UCase(strType) = "JUMP" Then
                            Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "CLOSED PN" & """" & strFil & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                bCadOk = True
                            Close #53
                         End If
                         'FUS 'jump
                         If UCase(strType) = "FUSE" Then
                            Open PrmPath & "ReadBomValue\JUMPER.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "OPEN PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                               bCadOk = True
                            Close #53
                         End If
                         'CONNECTOR
                         If UCase(strType) = UCase("CONNECTOR") Then
                            Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                               bCadOk = True
                            Close #53
                         End If
                         
                         
                         'TEST POINT conn
                           
                         If UCase(strType) = UCase("TEST POINT") Then
                            Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                               bCadOk = True
                            Close #53
                         End If
                         ' MECHANICAL conn
                         If UCase(strType) = UCase("MECHANICAL") Then
                            Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                               bCadOk = True
                            Close #53
                         End If
                         'RESISTOR NETWORK  pin lib
                         If InStr(UCase(strType), "RESISTOR NETWORK") <> 0 Then
                            Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                bCadOk = True
                            Close #53
                         End If
                         'CAPACITOR
                          If UCase(strType) = UCase("CAPACITOR") Then
                            Open PrmPath & "ReadBomValue\Capacitor.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "11.1p      20   20        f    PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                bCadOk = True
                            Close #53
                          End If
                         'RESISTOR
                          If UCase(strType) = UCase("RESISTOR") Then
                            Open PrmPath & "ReadBomValue\Resistor.txt" For Append As #53
                                Print #53, strFenPei(i); Tab(25); "8.88M        10   10        f    PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                bCadOk = True
                            Close #53
                          End If
                         
                            If UCase(strType) = "DIODE" Then
                            
                                                          'conn
                              If strTol = "ZZ.00PAD.0N1" Then
                                    Open PrmPath & "ReadBomValue\Connector.txt" For Append As #53
                                        Print #53, strFenPei(i); Tab(25); "NT;"; Tab(100); "! Bom in Cad not find device."
                                       bCadOk = True
                                    Close #53
                              End If
                              
                            
                            
                               'diode pin lib
                                 'CH715FPT-GP

                                If strVal = "CH715FPT-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                               '83.00054.Q81 BAT54C-U-GP
                                If strVal = "BAT54C-U-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                
                               'DA3X101F0L-GP
                                If strVal = "DA3X101F0L-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                               'BAT54CPT-GP

                                If strVal = "BAT54CPT-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                              'bas16
                                If strVal = "BAS16-6-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                             ' 'BAT54-7-F-GP
                                 If strVal = "BAT54-7-F-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                             
                             
                              'bav99
                                 If strVal = "BAW56-2-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                'bav99
                                 If strVal = "BAV99PT-GP-U" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                'BAT54A-7-F-1-GP
                                 If strVal = "BAT54A-7-F-1-GP" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                 'BAV99-5-GP-U
                                 If strVal = "BAV99-5-GP-U" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                             'SM24DTCT-GP-U
                                 If strVal = "SM24DTCT-GP-U" Then
                                  Open PrmPath & "ReadBomValue\Pin Library.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                            
                           'diode a c
                           
                               
                                'MS04A03T2V2-GP-U

                                If strVal = UCase("MS04A03T2V2-GP-U") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If

                           
                                'MMPZ5231BPT-GP

                                If strVal = UCase("MMPZ5231BPT-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If

                                  'B0530WS-7-F-GP

                                If strVal = UCase("B0530WS-7-F-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                 
                                


                                 '1SMA18AT3G-GP


                                If strVal = UCase("1SMA18AT3G-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                           
                                'CH551H-30PT-GP
                                If strVal = UCase("CH551H-30PT-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                'B240A-13-GP
                                If strVal = UCase("B240A-13-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                
                                'SDMK0340L-7-F-GP
                                If strVal = UCase("SDMK0340L-7-F-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If                               '
                               
                               
                               
                               'BZT52C15S-GP
                                If strVal = UCase("BZT52C15S-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                           
                            
                                'SMF18AT1G-GP
                                If strVal = UCase("SMF18AT1G-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                                'CH751H-40PT-GP
                                  If strVal = UCase("CH751H-40PT-GP") Then
                                  Open PrmPath & "ReadBomValue\Diode.txt" For Append As #53
                                      Print #53, strFenPei(i); Tab(25); "0.8       0.2       PN" & """" & strTol & """" & ";"; Tab(100); "! Bom in Cad not find device."
                                      bCadOk = True
                                  Close #53
                                End If
                            
                             End If 'diode
                         If bCadOk = False Then
                           Print #52, strTol; Tab(30); strType; Tab(55); strVal; Tab(100); strFenPei(i)
                         End If
                         bCadOk = False
                     End If
                          Print #50, strTol; Tab(30); strType; Tab(55); strVal; Tab(100); strFenPei(i)
                    bBomFind_ok = False
                Next
                bFindOk = False
                Erase strFenPei
                Erase strFenPei_Gan
                strDeviceConnectStr = ""
                intFindNumOk_To = ""
                intFindNumOk_Form = ""
                intFindNum = ""
                strTol = ""
                strType = ""
                TmpStr = ""
                strVal = ""
                tmpPart = ""
                strDeviceName_Gan_Add_All = ""
             End If
             Else
             bFindOk = False
        End If
        DoEvents
         fuck = fuck + 1
         Me.Caption = "Read file line: " & fuck
      Loop
   Close #49
   Close #51
   Close #50
   Close #52
End Sub
'Private Sub Read_Bom_Creat_dir()
'Dim Mystr As String
'
'On Error Resume Next
'   Open Trim(txtBomPath.Text) For Input As #51
'      Do Until EOF(51)
'        Line Input #51, Mystr
'      Loop
'   Close #51
'End Sub
