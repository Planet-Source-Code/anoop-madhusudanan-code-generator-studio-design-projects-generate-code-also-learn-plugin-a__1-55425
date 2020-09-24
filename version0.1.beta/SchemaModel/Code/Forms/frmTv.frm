VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmXml 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Loading Schema"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tvMain 
      Height          =   5925
      Left            =   105
      TabIndex        =   0
      Top             =   615
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   10451
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblMain 
      Caption         =   "Please wait while loading the schema.."
      Height          =   465
      Left            =   105
      TabIndex        =   1
      Top             =   240
      Width           =   3870
   End
End
Attribute VB_Name = "frmXml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


