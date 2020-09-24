VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Property Editor"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmPropertyEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstMain 
      Height          =   4155
      Left            =   60
      TabIndex        =   3
      Top             =   645
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5385
      TabIndex        =   2
      Top             =   4935
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4020
      TabIndex        =   1
      Top             =   4935
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "frmPropertyEditor.frx":000C
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblProperty 
      Caption         =   "You Can Double Click A Property To Change It"
      Height          =   270
      Left            =   810
      TabIndex        =   0
      Top             =   225
      Width           =   3990
   End
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Action As Boolean

Private Sub cmdCancel_Click()
Action = False
Unload Me

End Sub

Private Sub cmdOk_Click()
Action = True
    Me.Hide
End Sub

'====================================================================
Function ShowProperties(ObjectType, Sch As SchemaModel.Schema, ParamArray AttribGroups() As Variant)
'====================================================================

Dim i
Dim thisGroup As Object
Dim thisAttrib As Attrib
Dim thisItem As ListItem

On Error GoTo Parseerror

Me.Caption = "Properties - " & ObjectType
lstMain.ListItems.Clear

For i = LBound(AttribGroups) To UBound(AttribGroups)

Set thisGroup = AttribGroups(i)

For Each thisAttrib In thisGroup
Dim Item As ListItem
Debug.Print LCase(thisAttrib.AttribName)
    Set Item = Me.lstMain.ListItems.Add(, LCase(thisAttrib.AttribName), FormatName(thisAttrib.AttribName))
    Item.SubItems(1) = thisAttrib.AttribValue
    LoadOptions ObjectType, Item, Sch
Next

Next

Me.Show vbModal

If Action Then

    For i = LBound(AttribGroups) To UBound(AttribGroups)
    
    Set thisGroup = AttribGroups(i)
    
    For Each thisAttrib In thisGroup
        thisAttrib.AttribValue = lstMain.ListItems(LCase(thisAttrib.AttribName)).SubItems(1)
    Next
    
    Next

End If

Unload Me
Exit Function

Parseerror:
    MsgBox "Unable to parse the XML Schema. Invalid xml directives found", vbCritical + vbOKOnly, "Parse Error"

'====================================================================
End Function
'====================================================================


Function FormatName(Str)
FormatName = UCase(VBA.Left(Str, 1)) & LCase(VBA.Right(Str, Len(Str) - 1))
End Function

Private Sub lstMain_DblClick()
    EditItem
End Sub

'====================================================================
Sub EditItem()
'Edit the selected item in the active list view
'If the item has tag, a combo is displayed
'====================================================================

Dim Item, RVal
Dim CurList As ListView

Set CurList = Me.lstMain

On Error Resume Next
frmInput.txtItem.Visible = False
frmInput.cmbItem.Visible = False


If CurList.Visible = False Then Exit Sub

Set Item = CurList.SelectedItem
If Err Then Exit Sub

frmInput.lblField = Item.Text


    If Item.Tag = "" Then
        frmInput.txtItem.Text = Item.SubItems(1)
        frmInput.txtItem.Visible = True
        frmInput.txtItem.SetFocus
        frmInput.txtItem.Visible = True
        GoTo NoCombo
    End If
    
    Dim Vals() As String, i
    
    Vals = Split(Item.Tag, Chr$(10))
    frmInput.cmbItem.Clear
    frmInput.cmbItem.TabIndex = lstMain.TabIndex - 1
    For i = LBound(Vals) To UBound(Vals)
        frmInput.cmbItem.AddItem Vals(i)
    Next i
    
    For i = 0 To frmInput.cmbItem.ListCount - 1
        With frmInput.cmbItem
            If LCase(Trim(.List(i))) = LCase(Item.SubItems(1)) Then
                .ListIndex = i
            End If
        End With
    Next i
    
    frmInput.cmbItem.Visible = True

NoCombo:

CurList.HideSelection = True

RVal = -1
RVal = frmInput.ShowModal(Me.Left + lstMain.Left + 80, Me.Top + lstMain.Top + Item.Top + 100 + Item.Height, lstMain.Width)

If RVal <> -1 Then CurList.SelectedItem.SubItems(1) = RVal
CurList.HideSelection = False

    
'====================================================================
End Sub
'====================================================================


'====================================================================
Function LoadOptions(ObjectType, Item As ListItem, Sch As SchemaModel.Schema)
'Load options from properties.dat file
'====================================================================

Dim sLine As String

Open App.Path & "\appdata\properties.dat" For Input As #1

    Do While Not EOF(1)
        Line Input #1, sLine
    
    Dim Vals() As String
    
    On Error GoTo ParseNext
    
    Vals = Split(sLine, "#")
    If Trim(CStr(sLine)) <> "" Then
    If LCase(ObjectType) = LCase(Vals(0)) Then
        If LCase(Vals(1)) = LCase(Item.Text) Then
            Item.Tag = Replace(Vals(2), "|", Chr$(10))
            Item.Tag = Trim(Item.Tag)
            LoadSpecialOptions Item, Sch
            
            On Error Resume Next
            Item.SubItems(2) = Vals(3)
            Close #1
            Exit Function
        End If
    End If
    End If
    
    
    
ParseNext:
    Loop

Close #1

'====================================================================
End Function
'====================================================================


'====================================================================
Public Function LoadSpecialOptions(Item As ListItem, Sch As SchemaModel.Schema)
'Special directives like :listentities
'====================================================================

Dim thisEnt As SchemaModel.Entity
Select Case (LCase(Item.Tag))
    Case ":listentities"
        Item.Tag = ""
        For Each thisEnt In Sch.SchemaEntities
            Item.Tag = Item.Tag & thisEnt.EntityAttributes("name").AttribValue & Chr$(10)
        Next
        
    Case ":listfields"
        Dim thisFld As SchemaModel.Field
        
        Item.Tag = ""
        For Each thisEnt In Sch.SchemaEntities
            Item.Tag = Item.Tag & ">>" & thisEnt.EntityAttributes("name").AttribValue & Chr$(10)
            For Each thisFld In thisEnt.Fields
                Item.Tag = Item.Tag & "          " & thisFld.FieldHeaderAttributes("name").AttribValue & Chr$(10)
            Next
        Next

End Select


'====================================================================
End Function
'====================================================================

