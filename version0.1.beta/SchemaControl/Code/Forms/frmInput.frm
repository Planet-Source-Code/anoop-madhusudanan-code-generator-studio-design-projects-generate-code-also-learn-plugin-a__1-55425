VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   0  'None
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIn 
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   3090
      ScaleHeight     =   450
      ScaleWidth      =   735
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   15
      Width           =   735
      Begin VB.CommandButton cmdOk 
         Default         =   -1  'True
         Height          =   315
         Left            =   0
         Picture         =   "frmInput.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   330
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   360
         Picture         =   "frmInput.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1395
      TabIndex        =   1
      Text            =   "cmbItem"
      Top             =   15
      Width           =   1635
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1365
      TabIndex        =   0
      Top             =   15
      Width           =   1665
   End
   Begin VB.Label lblField 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Field"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   45
      Width           =   435
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================
'This form is for field input
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm


Public ReturnVal

Public Function ShowModal(X, Y, Wid)
On Error Resume Next
    Me.Move X, Y + 70
    Me.Width = Wid - 60
    Form_Resize
    Me.Show vbModal, frmProperty
    ShowModal = ReturnVal
End Function

Private Sub cmdCancel_Click()
ReturnVal = -1
Unload Me

End Sub

Private Sub cmdOk_Click()
    If cmbItem.Visible = True Then
        ReturnVal = Trim(cmbItem.Text)
    Else
        ReturnVal = Trim(txtItem.Text)
    End If
    Unload Me
End Sub

Private Sub Form_Deactivate()
cmdCancel_Click
End Sub

Private Sub Form_Resize()

On Error Resume Next
txtItem.Move lblField.Left + lblField.Width + 50
txtItem.Width = Me.ScaleWidth - picIn.Width - txtItem.Left - 100
cmbItem.Left = txtItem.Left
cmbItem.Width = txtItem.Width

picIn.Left = txtItem.Left + txtItem.Width + 50


End Sub

Private Sub txtItem_GotFocus()
txtItem.SelStart = 0
txtItem.SelLength = Len(txtItem.Text)
End Sub
