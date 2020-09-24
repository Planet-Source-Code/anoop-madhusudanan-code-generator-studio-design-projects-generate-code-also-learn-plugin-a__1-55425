VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProject 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate ASP Code"
   ClientHeight    =   3630
   ClientLeft      =   2925
   ClientTop       =   2985
   ClientWidth     =   6615
   Icon            =   "frmEngine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picProgress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   2805
      ScaleHeight     =   720
      ScaleWidth      =   3300
      TabIndex        =   7
      Top             =   2025
      Visible         =   0   'False
      Width           =   3300
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   285
         Left            =   15
         TabIndex        =   8
         Top             =   270
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Creating Project. Please Wait.."
         Height          =   465
         Left            =   15
         TabIndex        =   9
         Top             =   -15
         Width           =   4485
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   300
      TabIndex        =   6
      Top             =   2850
      Width           =   6225
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Text            =   "C:\Inetpub\Wwwroot"
      Top             =   1365
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   2865
      TabIndex        =   0
      Top             =   615
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5235
      TabIndex        =   4
      Top             =   3090
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3915
      TabIndex        =   3
      Top             =   3090
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "asp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FEEADE&
      Height          =   1410
      Left            =   225
      TabIndex        =   10
      Top             =   1860
      Width           =   2370
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   0
      Picture         =   "frmEngine.frx":058A
      Top             =   0
      Width           =   2760
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Path :"
      Height          =   255
      Left            =   2850
      TabIndex        =   2
      Top             =   1095
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   255
      Left            =   2865
      TabIndex        =   5
      Top             =   315
      Width           =   585
   End
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'This module is developed or Re-Applied by Anoop M
'anoop@logicmatrixonline.com
'
'http://www.logicmatrixonline.com/anoop
'
'frmProject : Loads a directory
'
'====================================================================================

'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm

Option Explicit
Public Sch As SchemaModel.Schema
Public AppObj As Object

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()

    If Trim(txtName.Text) = "" Then
        MsgBox "Kindly provide a valid project name", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    
    If txtPath.Text = "" Then
        MsgBox "Invalid path. Kindly specify a valid directory", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    Me.cmdLoad.Enabled = False
    Me.cmdCancel.Enabled = False
    
    
    Dim ASPEngine As clsASPEngine
    
    picProgress.Visible = True
    Set ASPEngine = New clsASPEngine
        ASPEngine.Create txtPath.Text, txtName.Text, Sch, AppObj, Me
    
    Unload Me
    
    Set ASPEngine = Nothing


End Sub


'Smart Editor By Anoop : (1/13/2002 11:09:25 AM) 9 + 35 = 44 Lines

Private Sub Picture1_Click()

End Sub
