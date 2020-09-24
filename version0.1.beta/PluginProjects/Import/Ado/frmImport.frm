VERSION 5.00
Begin VB.Form frmImport 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import ADO Schema"
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
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   4215
      TabIndex        =   3
      Top             =   2310
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Import"
      Height          =   360
      Left            =   2835
      TabIndex        =   2
      Top             =   2310
      Width           =   1215
   End
   Begin VB.TextBox txtConnection 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   1800
      Width           =   5220
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1410
      Left            =   2955
      TabIndex        =   4
      Top             =   75
      Width           =   2370
   End
   Begin VB.Label lblMain 
      BackStyle       =   0  'Transparent
      Caption         =   "ADO Connection String:"
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
      Left            =   165
      TabIndex        =   0
      Top             =   1515
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
