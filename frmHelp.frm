VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "help"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3360
      Picture         =   "frmHelp.frx":0000
      Top             =   8640
      Width           =   5685
   End
   Begin VB.Image Image1 
      Height          =   8550
      Left            =   0
      Picture         =   "frmHelp.frx":4782
      Top             =   0
      Width           =   9120
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
Unload Me
End Sub

