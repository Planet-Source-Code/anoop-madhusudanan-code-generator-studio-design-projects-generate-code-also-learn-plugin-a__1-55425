VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import DAO Schema"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   300
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowser 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      Height          =   345
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Browse File"
      Top             =   1830
      Width           =   645
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4215
      TabIndex        =   3
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Import"
      Height          =   360
      Left            =   2835
      TabIndex        =   2
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox txtConnection 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Text            =   "c:\program files\microsoft visual studio\vb98\nwind.mdb"
      Top             =   1845
      Width           =   4545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "dao"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1410
      Left            =   2955
      TabIndex        =   4
      Top             =   75
      Width           =   2370
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Access 97 Format MDB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   0
      Top             =   1545
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   -15
      Picture         =   "frmImport.frx":000C
      Top             =   -750
      Width           =   2760
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public FileName As String

Private Sub cmdBrowser_Click()
cdMain.Filter = "Access 97 Database|*.mdb"
cdMain.ShowOpen
txtConnection.Text = cdMain.FileName

End Sub

Private Sub cmdCancel_Click()
FileName = ""
Unload Me
End Sub

Private Sub cmdStart_Click()
FileName = Me.txtConnection.Text
Unload Me
End Sub

Public Function GetFile()
Me.Show vbModal
GetFile = FileName
End Function
