VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCreate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   990
      Left            =   1095
      TabIndex        =   2
      Top             =   1785
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1746
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   405
      Left            =   990
      TabIndex        =   1
      Top             =   585
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   714
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      Caption         =   "Creating Project. Please Wait.."
      Height          =   465
      Left            =   990
      TabIndex        =   0
      Top             =   240
      Width           =   4485
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmCreate.frx":0000
      Top             =   405
      Width           =   480
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================
'Used by clsASPEngine
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================

'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm


Private Sub Form_Deactivate()
'On Error Resume Next
Me.Refresh

End Sub

