VERSION 5.00
Begin VB.Form frmBomVerOption 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Option"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3735
         Begin VB.CheckBox CheckCn 
            Caption         =   "Connector"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   1320
            TabIndex        =   4
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Checklb 
            Caption         =   "Pin Linrary"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CommandButton mcGo 
            Caption         =   "&OK"
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
            Left            =   2520
            TabIndex        =   2
            Top             =   280
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmBomVerOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
