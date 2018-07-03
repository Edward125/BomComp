VERSION 5.00
Begin VB.Form frmFileName 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please input output  file name !"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   615
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtSN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Text1_Click()
   txtSN.SelStart = 0
   txtSN.SelLength = Len(txtSN.Text)
End Sub

Private Sub cmdOK_Click()
strOutFileName = Trim(txtSN.Text)
strOutFileName = Replace(strOutFileName, "*", "")
strOutFileName = Replace(strOutFileName, """", "")
strOutFileName = Replace(strOutFileName, "|", "")
strOutFileName = Replace(strOutFileName, "/", "")
strOutFileName = Replace(strOutFileName, "\", "")
strOutFileName = Replace(strOutFileName, "<", "")
strOutFileName = Replace(strOutFileName, ">", "")
strOutFileName = Trim(strOutFileName)
txtSN.Text = strOutFileName

If InStr(strOutFileName, ".") = 0 Then strOutFileName = strOutFileName & ".txt"
 If strOutFileName = "" Then
   Me.Show
   txtSN.SetFocus
 End If
 Unload Me
End Sub

Private Sub Form_Load()
txtSN.Text = strFileNameOpen
End Sub
