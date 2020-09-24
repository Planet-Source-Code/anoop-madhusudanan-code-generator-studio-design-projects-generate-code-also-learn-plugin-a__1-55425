VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl Entity 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2085
   DefaultCancel   =   -1  'True
   ScaleHeight     =   2025
   ScaleWidth      =   2085
   Begin MSComctlLib.ListView lstMain 
      Height          =   1155
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Double Click A Field"
      Top             =   270
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   2037
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Img"
      SmallIcons      =   "Img"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList Img 
      Left            =   900
      Top             =   2130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "entity.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "entity.ctx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "entity.ctx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   15
      ScaleHeight     =   255
      ScaleWidth      =   2010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1770
      Width           =   2010
      Begin VB.Image imgSize 
         Height          =   240
         Left            =   1785
         MousePointer    =   8  'Size NW SE
         Picture         =   "entity.ctx":10CE
         ToolTipText     =   "Click And Drag To Resize"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picTitle 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   2040
      Begin VB.PictureBox picBtn 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1515
         ScaleHeight     =   285
         ScaleWidth      =   525
         TabIndex        =   4
         Top             =   0
         Width           =   525
         Begin VB.CommandButton cmdEdit 
            Caption         =   "E"
            Default         =   -1  'True
            Height          =   270
            Left            =   270
            TabIndex        =   6
            ToolTipText     =   "Edit"
            Top             =   0
            Width           =   240
         End
         Begin VB.CommandButton cmdSize 
            Caption         =   "<"
            Height          =   270
            Left            =   15
            TabIndex        =   5
            ToolTipText     =   "Minimize Or Restore"
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Label lblMain 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   30
         Width           =   420
      End
   End
End
Attribute VB_Name = "Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'====================================================================
'An Entity In A Schema
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm


'Event Declarations:
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picTitle,picTitle,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picTitle,picTitle,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=picTitle,picTitle,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."

Event EditClick()
Event CloseClick()
Event Resized()

Event Error(Details As String)



'Some Variables
Dim mbSizing As Boolean
Public FieldList As ListView
Public TagStr As String

Dim myEntity As SchemaModel.Entity
Dim mySchema As SchemaModel.Schema

Dim prevX, prevY
Dim CurW, CurH





'====================================================================
Function LoadEntity(thisEnt As SchemaModel.Entity, thisSch As SchemaModel.Schema)
'Loads an entity to the module, from the schema object
'====================================================================

    Dim thisFld As SchemaModel.Field
    Dim thisAttrib As SchemaModel.Attrib
    
    
    Set myEntity = thisEnt
    Set mySchema = thisSch
    
    
    
    
    lstMain.ListItems.Clear
    
    For Each thisFld In thisEnt.Fields
            Set It = lstMain.ListItems.Add(, thisFld.FieldHeaderAttributes("name").AttribValue, thisFld.FieldHeaderAttributes("name").AttribValue)
             If LCase(thisFld.FieldHeaderAttributes("primary").AttribValue) = "true" Then It.SmallIcon = 1
        
        Props = ""
    '            For Each thisAttrib In thisFld.FieldAttributes
    '                    Props = Props + RetXML(thisAttrib.AttribName, thisAttrib.AttribValue) & vbCrLf
    '            Next
    '         It.Tag = Props
EndProps:
        
        If LCase(thisFld.FieldHeaderAttributes("display").AttribValue) = "true" Then It.Checked = True
    
    Next
    
    Dim mWidth, mHeight, mState
    
    On Error Resume Next
    
    mWidth = Val(thisEnt.EntityAttributes("width").AttribValue)
    mHeight = Val(thisEnt.EntityAttributes("height").AttribValue)
    mState = thisEnt.EntityAttributes("state").AttribValue
    
        If mWidth = 0 Then mWidth = UserControl.Width
        If mHeight = 0 Then mHeight = UserControl.Height
        
        If CStr(mState) = "0" Then
            mState = 0
        Else
            mState = 1
        End If
        
    
    
    UserControl.Width = mWidth
    UserControl.Height = mHeight
    
    If mState = 0 Then
        MinEntity
    Else
        MaxEntity
    End If
    
    UserControl_Resize
    
'====================================================================
End Function
'====================================================================

'====================================================================
Function AddRelation(FieldName As String, Relation As String) As Boolean
'For adding a relation to our entity
'====================================================================

AddRelation = True
On Error Resume Next
lstMain.ListItems(FieldName).Tag = Relation
If Err Then AddRelation = False
End Function

'====================================================================
Function GetRelation(FieldName As String)
'Returns a relation from an item
'====================================================================

GetRelation = -1
On Error Resume Next
GetRelation = lstMain.ListItems(FieldName).Tag
'====================================================================
End Function
'====================================================================

'====================================================================
Function GetItem(FieldName As String, Item As ListItem) As Boolean
'If there is an item, returns the item
'====================================================================

GetItem = True
On Error Resume Next
Set Item = lstMain.ListItems(FieldName)
If Err Then GetItem = False

'====================================================================
End Function
'====================================================================


Function InitArrays()
'Initialize an array


End Function


'====================================================================
'Event Handlers
'====================================================================

Private Sub cmdClose_Click()
RaiseEvent CloseClick
End Sub

Private Sub cmdEdit_Click()
    RaiseEvent EditClick
    On Error Resume Next
    
    frmProperty.ShowProperties "Entity", mySchema, myEntity.EntityAttributes
    
    If Err Then
        MsgBox "Parse Error. " & Err.Description, vbCritical + vbOKOnly, "Unable To Edit"
    Else
        Me.LoadEntity myEntity, mySchema
    End If
    

End Sub

Private Sub cmdSize_Click()

If cmdSize.Caption = "<" Then
CurW = UserControl.Width
CurH = UserControl.Height
UserControl.Height = picTitle.Height + 20
UserControl.Width = 1800
cmdSize.Caption = ">"
Else
UserControl.Height = CurH
UserControl.Width = CurW
cmdSize.Caption = "<"
End If
Updatestate
RaiseEvent Resized
End Sub

Private Sub imgSize_DblClick()
cmdSize_Click

End Sub

Private Sub imgSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbSizing = True
prevX = X
prevY = Y
End Sub

Private Sub imgSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If mbSizing Then
UserControl.Height = UserControl.Height + Y - prevY
UserControl.Width = UserControl.Width + X - prevX
RaiseEvent Resized
End If

End Sub

Private Sub imgSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbSizing = False
If UserControl.Height < picTitle.Height + picBottom.Height Then UserControl.Height = picTitle.Height + picBottom.Height + 100
If UserControl.Width < 1000 Then UserControl.Width = 1000
RaiseEvent Resized

End Sub

Private Sub lblMain_DblClick()
cmdSize_Click

End Sub

Private Sub lblMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMain.ToolTipText = "Entity : " & lblMain.Caption
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lstMain_DblClick()
    Err.Clear
        EditField
        
    If Err Then
        MsgBox "Parse Error. " & Err.Description, vbCritical + vbOKOnly, "Unable To Edit"
    Else
            Me.LoadEntity myEntity, mySchema
    End If
    
    
End Sub

Public Sub EditField()
On Error Resume Next
Dim Ret As String, thisField As SchemaModel.Field
Dim thisItem As ListItem

Set thisItem = lstMain.SelectedItem

If Err Then Exit Sub

For Each thisField In myEntity.Fields
    If LCase(thisField.FieldHeaderAttributes("name").AttribValue) = LCase(thisItem.Text) Then
        Ret = frmProperty.ShowProperties("Field", mySchema, thisField.FieldHeaderAttributes, thisField.FieldAttributes)
        Exit Sub
    End If
Next

End Sub

Private Sub lstMain_LostFocus()
RaiseEvent Resized
End Sub

Private Sub lstMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Resized
End Sub

Private Sub picTitle_DblClick()
cmdSize_Click

End Sub


Private Sub UserControl_Initialize()
Set FieldList = lstMain
InitArrays

On Error Resume Next



End Sub

Private Sub UserControl_Resize()
On Error Resume Next

picTitle.Move 0, 0, UserControl.ScaleWidth

picBtn.Move picTitle.ScaleWidth - picBtn.Width - 50
lblMain.Width = picBtn.Left - 100


lstMain.Move 0, lstMain.Top, UserControl.ScaleWidth - lstMain.Left, UserControl.ScaleHeight - (2 * picBottom.Height)

picBottom.Top = lstMain.Height + lstMain.Top
picBottom.Width = UserControl.ScaleWidth
picBottom.Left = 0


lstMain.ColumnHeaders(1).Width = lstMain.Width - 250
imgSize.Left = picBottom.ScaleWidth - imgSize.Width

If UserControl.Height = picTitle.Height + 20 Or UserControl.Width = 1800 Then
cmdSize.Caption = ">"
Else
cmdSize.Caption = "<"
End If

Updatestate

End Sub
Private Sub picTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picTitle,picTitle,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = picTitle.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    picTitle.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picTitle,picTitle,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = picTitle.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set picTitle.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picTitle.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    picTitle.BackColor = PropBag.ReadProperty("Titlecolor", &H8000000C)
    lblMain.Caption = PropBag.ReadProperty("Caption", "Name")
    m_ConnectionString = PropBag.ReadProperty("ConnectionString", m_def_ConnectionString)
    m_DataSource = PropBag.ReadProperty("DataSource", m_def_DataSource)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MousePointer", picTitle.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Titlecolor", picTitle.BackColor, &H8000000C)
    Call PropBag.WriteProperty("Caption", lblMain.Caption, "Name")
    Call PropBag.WriteProperty("ConnectionString", m_ConnectionString, m_def_ConnectionString)
    Call PropBag.WriteProperty("DataSource", m_DataSource, m_def_DataSource)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picTitle,picTitle,-1,BackColor
Public Property Get Titlecolor() As OLE_COLOR
Attribute Titlecolor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    Titlecolor = picTitle.BackColor
End Property

Public Property Let Titlecolor(ByVal New_Titlecolor As OLE_COLOR)
    picTitle.BackColor() = New_Titlecolor
    picBtn.BackColor() = New_Titlecolor
    PropertyChanged "Titlecolor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMain,lblMain,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblMain.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblMain.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get List() As Object
    Set List = lstMain
End Property


Public Function MaxEntity()
On Error Resume Next

If cmdSize.Caption = ">" Then
UserControl.Height = CurH
UserControl.Width = CurW
cmdSize.Caption = "<"
End If
Updatestate
RaiseEvent Resized

End Function

Public Function MinEntity()
On Error Resume Next

If cmdSize.Caption = "<" Then
    CurW = UserControl.Width
    CurH = UserControl.Height
    UserControl.Height = picTitle.Height + 20
    UserControl.Width = 1800
    cmdSize.Caption = ">"
End If
Updatestate
RaiseEvent Resized

End Function


Public Function Updatestate()
If cmdSize.Caption = "<" Then
    myState = 1
Else
    myState = 0
End If

SetAttrib "State", myState
End Function

Public Function UpdatePosition(myTop, myLeft, myWidth, myHeight)


SetAttrib "Left", myLeft
SetAttrib "Top", myTop
SetAttrib "Height", myHeight
SetAttrib "Width", myWidth
Updatestate
End Function

Public Function SetAttrib(Param, Value)
Err.Clear
On Error Resume Next
myEntity.EntityAttributes(LCase(Param)).AttribValue = Value

If Err Then
myEntity.EntityAttributes.Add Param, Value
End If

End Function


                                                                                                                       
cmdSize.Caption = ">"
Else
UserControl.Height = CurH
UserControl.Width = CurW
cmdSize.Caption = "<"
End If
Updatestate
RaiseEvent Resized
End Sub

Private Sub imgSize_DblClick()
cmdSize_Click

End Sub

Private Sub imgSize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbSizing = True
prevX = X
prevY = Y
End Sub

Private Sub imgSize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If mbSizing Then
UserControl.Height = UserControl.Height + Y - prevY
UserControl.Width = UserControl.Width + X - prevX
RaiseEvent Resized
End If

End Sub

Private Sub imgSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mbSizing = False
If UserControl.Height < picTitle.Height + picBottom.Height Then UserControl.Height = picTitle.Height + picBottom.Height + 100
If UserControl.Width < 1000 Then UserControl.Width = 1000
RaiseEvent Resized

End Sub

Private Sub lblMain_DblClick()
cmdSize_Click

End Sub

Private Sub lblMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMain.ToolTipText = "Entity : " & lblMain.Caption
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub