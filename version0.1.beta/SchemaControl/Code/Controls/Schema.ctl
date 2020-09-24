VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl Schema 
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10755
   ScaleHeight     =   6375
   ScaleWidth      =   10755
   Begin VB.PictureBox picIn 
      BackColor       =   &H00FEE6E2&
      Height          =   6105
      Left            =   3855
      ScaleHeight     =   6045
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   105
      Width           =   6735
      Begin SchemaControl.Entity entMain 
         CausesValidation=   0   'False
         Height          =   2160
         Index           =   0
         Left            =   5100
         TabIndex        =   8
         Top             =   12000
         Visible         =   0   'False
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   3810
      End
      Begin VB.HScrollBar hsMain 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   5430
         Width           =   2565
      End
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5640
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   3
         Top             =   5085
         Width           =   285
      End
      Begin VB.VScrollBar vsMain 
         Height          =   1665
         Left            =   6855
         TabIndex        =   1
         Top             =   0
         Width           =   300
      End
      Begin VB.PictureBox Picture1 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   4
         Top             =   0
         Width           =   0
      End
      Begin VB.Image imgLin 
         Height          =   240
         Left            =   2310
         Picture         =   "Schema.ctx":0000
         Top             =   12000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgArrow 
         Height          =   240
         Index           =   0
         Left            =   2235
         Picture         =   "Schema.ctx":058A
         ToolTipText     =   "Double Click To Edit Relation"
         Top             =   12000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgLine 
         Height          =   240
         Index           =   0
         Left            =   2235
         Picture         =   "Schema.ctx":06D4
         ToolTipText     =   "Double Click To Edit Relation"
         Top             =   12000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Line linMain 
         BorderColor     =   &H00000000&
         Index           =   0
         Visible         =   0   'False
         X1              =   6675
         X2              =   1635
         Y1              =   -30
         Y2              =   -30
      End
      Begin VB.Image imgRight 
         Height          =   240
         Left            =   1530
         Picture         =   "Schema.ctx":0C5E
         Top             =   12000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgLeft 
         Height          =   240
         Left            =   1875
         Picture         =   "Schema.ctx":11E8
         Top             =   12000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Line linSub 
         BorderColor     =   &H80000014&
         Index           =   0
         Visible         =   0   'False
         X1              =   2595
         X2              =   2160
         Y1              =   1305
         Y2              =   4260
      End
   End
   Begin VB.PictureBox picXML 
      BorderStyle     =   0  'None
      Height          =   3510
      Left            =   960
      ScaleHeight     =   3510
      ScaleWidth      =   4575
      TabIndex        =   7
      Top             =   270
      Visible         =   0   'False
      Width           =   4575
      Begin SHDocVwCtl.WebBrowser wbMain 
         Height          =   2670
         Left            =   15
         TabIndex        =   9
         Top             =   75
         Width           =   2250
         ExtentX         =   3969
         ExtentY         =   4710
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSComctlLib.ImageList imgSchema 
         Left            =   3180
         Top             =   2835
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Schema.ctx":1772
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Schema.ctx":1D0C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView tvMain 
      Height          =   795
      Left            =   255
      TabIndex        =   5
      Top             =   4725
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1402
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList img 
      Left            =   4800
      Top             =   5820
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Schema.ctx":22A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Schema.ctx":2840
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbMain 
      Height          =   4275
      Left            =   255
      TabIndex        =   6
      Top             =   240
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   7541
      Style           =   2
      HotTracking     =   -1  'True
      Placement       =   1
      Separators      =   -1  'True
      ImageList       =   "img"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Entity View"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "XML View"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Schema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'====================================================================
'Schema Control
'This control is for displaying and designing the application schema
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm



'Variables
Dim mbMoving() As Boolean

Dim mTotalRel As Integer
Dim prevX, prevY
Dim CurEntity As Long


Public Arranged As Boolean

'Schema Model Datastructure
Dim Sch As SchemaModel.Schema



Public Event Dirty()

Private Sub entMain_EditClick(Index As Integer)
Dim i

On Error Resume Next

For i = 0 To entMain.Count - 1
    entMain(i).Titlecolor = vbApplicationWorkspace
Next

entMain(Index).SetFocus

entMain(Index).Titlecolor = vbActiveTitleBar

End Sub

Private Sub entMain_GotFocus(Index As Integer)

Dim i

For i = 0 To entMain.Count - 1
    entMain(i).Titlecolor = vbApplicationWorkspace
    entMain(i).ZOrder 1
Next

entMain(Index).Titlecolor = vbActiveTitleBar

On Error Resume Next
CurEntity = Index

entMain(Index).ZOrder 0
vsMain.ZOrder 0
hsMain.ZOrder 0


End Sub

'Database Objects

Private Sub entMain_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Moving Procedure
On Error Resume Next
entMain(Index).SetFocus
mbMoving(Index) = True

prevX = X
prevY = Y

End Sub

Private Sub entMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Moving Procedure
On Error Resume Next


If mbMoving(Index) Then
    Arranged = True
    RaiseEvent Dirty
    entMain(Index).Move entMain(Index).Left + X - prevX, entMain(Index).Top + Y - prevY
    DrawRelations
End If


End Sub

Private Sub entMain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Moving Procedure
mbMoving(Index) = False



UserControl_Resize
picIn.SetFocus


RaiseEvent Dirty

End Sub

Private Sub entMain_Resized(Index As Integer)
DrawRelations
End Sub

Private Sub imgArrow_DblClick(Index As Integer)
'EditRelation Index
On Error Resume Next
Dim thisRel As SchemaModel.Relation

For Each thisRel In Sch.SchemaRelations
    If thisRel.InnerIndex = Index Then
        frmProperty.ShowProperties "Relation", Sch, thisRel.RelationAttributes
    End If
Next


End Sub


Private Sub imgLine_dblClick(Index As Integer)
'EditRelation Index
On Error Resume Next
Dim thisRel As SchemaModel.Relation

For Each thisRel In Sch.SchemaRelations
    If thisRel.InnerIndex = Index Then
        frmProperty.ShowProperties "Relation", Sch, thisRel.RelationAttributes
    End If
Next

End Sub


Private Sub tbMain_Click()
On Error Resume Next
Select Case tbMain.SelectedItem.Index
    Case 1
        picIn.Visible = True
        picXML.Visible = False
        
    Case 2
        
        Dim s As String
        On Error GoTo NoChange
        s = Me.DumpSchema()
        
        Open App.Path & "\temp.xml" For Output As #1
            Print #1, s
        Close #1
        
        If Trim(s) <> "" Then
            wbMain.Navigate2 App.Path & "\temp.xml"
        Else
            wbMain.Navigate2 "about:blank"
        End If
        
        
        picXML.Visible = True
        picIn.Visible = False
        Exit Sub
        
NoChange:
    tbMain.Tabs(1).Selected = True
    
        
End Select
End Sub

Private Sub UserControl_Initialize()

On Error Resume Next
Arranged = False
mbSliding = False
 

For i = 0 To entMain.Count - 1
    entMain(i).Visible = False
    Unload entMain(i)
Next i

wbMain.Navigate2 "about:blank"



End Sub

Public Sub EditField()
On Error Resume Next
entMain(CurEntity).EditField

End Sub


Private Sub UserControl_Resize()

'HANDLES THE RESIZING ALGORITHM

On Error Resume Next
tbMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight - 100
picIn.Move 75, 75, tbMain.Width - 150, tbMain.Height - 500
picXML.Move picIn.Left, picIn.Top, picIn.Width, picIn.Height
wbMain.Move 0, 0, picIn.ScaleWidth, picIn.ScaleHeight


vsMain.Width = hsMain.Height
vsMain.Move picIn.ScaleWidth - vsMain.Width, 0, vsMain.Width, picIn.ScaleHeight - hsMain.Height
hsMain.Move 0, picIn.ScaleHeight - hsMain.Height, picIn.ScaleWidth - hsMain.Height


Dim LeftMost, RightMost, TopMost, BottomMost, i

LeftMost = 0
RightMost = picIn.ScaleWidth
TopMost = 0
BottomMost = picIn.ScaleHeight

For i = 0 To entMain.Count - 1
With entMain(i)
    If entMain(i).Left < LeftMost Then LeftMost = entMain(i).Left
    If entMain(i).Left + .Width + 200 > RightMost Then RightMost = .Left + .Width + 200
    If .Top < TopMost Then TopMost = .Top
    If .Top + .Height + 200 > BottomMost Then BottomMost = .Top + .Height + 200
    
    entMain(i).UpdatePosition entMain(i).Top, entMain(i).Left, entMain(i).Width, entMain(i).Height

End With
Next




If TopMost >= 0 And BottomMost <= picIn.ScaleHeight Then
    vsMain.Enabled = False
Else
    
    vsMain.Min = TopMost
    vsMain.Max = BottomMost - picIn.ScaleHeight
    vsMain.SmallChange = 20
    vsMain.LargeChange = 100
    
    vsMain.Enabled = True
End If

If LeftMost >= 0 And RightMost <= picIn.ScaleWidth Then
    hsMain.Enabled = False
Else
    hsMain.Min = LeftMost
    hsMain.Max = RightMost - picIn.ScaleWidth
    hsMain.SmallChange = 20
    hsMain.LargeChange = 100
    
    hsMain.Enabled = True
End If

picBack.Move hsMain.Width, vsMain.Height

DrawRelations

End Sub




'====================================================================
Sub DrawRelations()
'====================================================================
'RESIZE ALL RELATIONS
'====================================================================

On Error Resume Next
Dim thisRel As SchemaModel.Relation

Dim EntFrom As Object, FromField As String, EntTo As Object, ToField As String, RelType
Dim FieldFrom As ListItem, FieldTo As ListItem

Dim i


On Error Resume Next
    If entMain(1).Visible = False Then Exit Sub
    If Err Then Exit Sub
Err.Clear

For Each thisRel In Sch.SchemaRelations
           
        Set EntFrom = Entity(thisRel.RelationAttributes("table").AttribValue)
        Set EntTo = Entity(thisRel.RelationAttributes("foreigntable").AttribValue)
        
        FromField = thisRel.RelationAttributes("field").AttribValue
        ToField = thisRel.RelationAttributes("foreignfield").AttribValue
        RelType = Val(thisRel.RelationAttributes("type").AttribValue)
        
        Set FieldFrom = EntFrom.FieldList.ListItems(FromField)
        Set FieldTo = EntTo.FieldList.ListItems(ToField)
        
        If FieldTo.SmallIcon <> 1 Then FieldTo.SmallIcon = 2
        
        i = thisRel.InnerIndex
        If EntTo.Left > EntFrom.Left Then
                AlignPicRight imgLine(i), EntFrom, FieldFrom
                AlignPicLeft imgArrow(i), EntTo, FieldTo
                imgArrow(i).Picture = imgRight.Picture
                linMain(i).Y1 = imgLine(i).Top + imgLine(i).Height / 2
                linMain(i).Y2 = imgArrow(i).Top + imgArrow(i).Height / 2
                linMain(i).X1 = imgLine(i).Left + imgLine(i).Width
                linMain(i).X2 = imgArrow(i).Left
            Else
                AlignPicLeft imgLine(i), EntFrom, FieldFrom
                AlignPicRight imgArrow(i), EntTo, FieldTo
                imgArrow(i).Picture = imgLeft.Picture
                linMain(i).Y2 = imgLine(i).Top + imgLine(i).Height / 2
                linMain(i).Y1 = imgArrow(i).Top + imgArrow(i).Height / 2
                linMain(i).X2 = imgLine(i).Left
                linMain(i).X1 = imgArrow(i).Left + imgArrow(i).Width
        End If
        
            
        If RelType = 0 Then
            imgArrow(i).Picture = imgLin.Picture
        End If
        
        With linMain(i)
            linSub(i).X1 = .X1
            linSub(i).X2 = .X2
            linSub(i).Y1 = .Y1
            linSub(i).Y2 = .Y2
        End With
        
        
Next
        
        
        



On Error Resume Next
For i = entMain.Count - 1 To 0 Step -1
    entMain(i).TabIndex = i
Next

vsMain.TabIndex = i
vsMain.TabIndex = i + 1



End Sub


'====================================================================
Sub DelRelation(thisRel As SchemaModel.Relation)
'Deletes A Relation
'====================================================================

'On Error resume next
Dim i

        i = thisRel.InnerIndex
        
        Unload imgArrow(i)
        Unload imgLine(i)
        Unload linMain(i)
        Unload linSub(i)
        
End Sub

'====================================================================
Sub AddRelation(thisRel As SchemaModel.Relation)
'====================================================================


'On Error resume next
    Dim i As Long
    'Will be unique all the time
        i = thisRel.InnerIndex
        Load imgArrow(i)
        Load imgLine(i)
        Load linMain(i)
        Load linSub(i)
        
        'If FieldTo.SmallIcon <> 1 Then FieldTo.SmallIcon = 2
        

End Sub

Private Sub AlignPicRight(Imagebox As Image, Ent As Object, Item As ListItem)
'Aligns the Imagebox to the right

Imagebox.Move Ent.Left + Ent.Width, Ent.Top + Ent.FieldList.Top + Item.Top

If Imagebox.Top > (Ent.Top + Ent.Height) Then
    Imagebox.Top = Ent.Top + Ent.Height - Imagebox.Height / 2
End If

If Imagebox.Top < Ent.Top Then
    Imagebox.Top = Ent.Top + Imagebox.Height / 2
End If

End Sub

Function AlignPicLeft(Imagebox, Ent, Item)
'Aligns the Imagebox to the left


Imagebox.Move Ent.Left - Imagebox.Width, Ent.Top + Ent.FieldList.Top + Item.Top
If Imagebox.Top > (Ent.Top + Ent.Height) Then
    Imagebox.Top = Ent.Top + Ent.Height - Imagebox.Height / 2
End If

If Imagebox.Top < Ent.Top Then
    Imagebox.Top = Ent.Top + Imagebox.Height / 2
End If

End Function

Private Sub hsMain_Change()
For i = 0 To entMain.Count - 1
With entMain(i)
    .Left = .Left - hsMain.Value
End With
Next
UserControl_Resize


End Sub


Private Sub UserControl_Terminate()
On Error Resume Next
Set Sch = Nothing


End Sub

Private Sub vsMain_Change()
For i = 0 To entMain.Count - 1
With entMain(i)
    .Top = .Top - vsMain.Value
End With
Next
UserControl_Resize

End Sub

'====================================================================
Function ShowSchema()
'Shows schema as diagram
''====================================================================

Dim thisEnt As SchemaModel.Entity
Dim thisRel As SchemaModel.Relation


Dim i

ReDim mbMoving(Sch.SchemaEntities.Count + 1)

For Each thisEnt In Sch.SchemaEntities

On Error Resume Next
i = thisEnt.InnerIndex

    Load entMain(i)
    entMain(i).LoadEntity thisEnt, Sch
    entMain(i).Caption = thisEnt.EntityAttributes("name").AttribValue
    mbMoving(i) = False
    entMain(i).Left = CLng(thisEnt.EntityAttributes("left").AttribValue)
    entMain(i).Top = CLng(thisEnt.EntityAttributes("top").AttribValue)
    Arranged = True
Next


For Each thisRel In Sch.SchemaRelations
On Error Resume Next
    AddRelation thisRel
Next


For Each thisEnt In Sch.SchemaEntities
On Error Resume Next
i = thisEnt.InnerIndex
    entMain(i).Visible = True
Next


For Each thisRel In Sch.SchemaRelations
i = thisRel.InnerIndex
        imgArrow(i).Visible = True
        imgLine(i).Visible = True
        linSub(i).Visible = True
        linMain(i).Visible = True
Next
        

DrawRelations



End Function


'====================================================================
Public Function LoadSchema(Schema As SchemaModel.Schema)
''====================================================================

Flush

On Error Resume Next
Set Sch = Nothing

Set Sch = Schema
ShowSchema


End Function


'====================================================================
Public Function LoadSchemaFromFile(SName As String)
''====================================================================

Flush

On Error Resume Next
Set Sch = Nothing

Set Sch = New SchemaModel.Schema
LoadSchemaFromFile = Sch.LoadSchema(SName)


ShowSchema


End Function



'====================================================================
Public Function LoadDatabase(DBName As String)
'Loads a database to the schema
'====================================================================

Flush


'If DBToSchema(DBName) = True Then
'    ShowSchema
'    LoadDatabase = True
'    Arranged = False
'    Arrange
'
'Else
'    LoadDatabase = False
'End If


'====================================================================
End Function
'====================================================================


Public Sub Arrange()
'For arranging all entities in the schema

If Arranged = True Then
UserControl_Resize
Exit Sub
End If


Dim i, row
row = 1

For i = 0 To entMain.Count - 1
    If i > 0 Then
        
        entMain(i).Left = entMain(i - 1).Left + entMain(i).Width + 200
        entMain(i).Top = entMain(i - 1).Top
        
        If entMain(i).Left + entMain(i).Height > picIn.ScaleWidth Then
            entMain(i).Left = 200
            entMain(i).Top = entMain(i - 1).Height + entMain(i - 1).Top + 200
            row = row + 1
        End If
    Else
        entMain(i).Left = 200
        entMain(i).Top = 200
    End If
    
        entMain(i).UpdatePosition entMain(i).Top, entMain(i).Left, entMain(i).Width, entMain(i).Height

Next

UserControl_Resize

End Sub

Function Entity(mName As String) As Object
'Returns the entity from table
Dim thisEnt As SchemaModel.Entity

For Each thisEnt In Sch.SchemaEntities
    If LCase(CStr(thisEnt.EntityAttributes("name").AttribValue)) = LCase(mName) Then
        Set Entity = entMain(thisEnt.InnerIndex)
        Exit Function
    End If
Next


End Function

'====================================================================
Function DumpSchema() As String
'Returns the schema tag file
'====================================================================
On Error Resume Next
DumpSchema = Sch.DumpSchema()
If Err Then DumpSchema = ""

'====================================================================
End Function
'====================================================================

Function GetSchema() As SchemaModel.Schema
    Set GetSchema = Sch
End Function



'====================================================================
Public Function Flush()
'Flushes the entity
'====================================================================

If Sch Is Nothing Then Exit Function

Dim thisEnt As SchemaModel.Entity
Dim thisRel As SchemaModel.Relation

For Each thisEnt In Sch.SchemaEntities
    Unload entMain(thisEnt.InnerIndex)
Next

For Each thisRel In Sch.SchemaRelations
    DelRelation thisRel
Next


Set Sch = Nothing


End Function


'====================================================================
Public Function Restore(Optional Minimize As Boolean = False)
'Restore all entities
'====================================================================

Dim Ent
    If Not Minimize Then
       For Ent = 0 To entMain.Count - 1
            entMain(Ent).MaxEntity
       Next
    Else
       For Ent = 0 To entMain.Count - 1
            entMain(Ent).MinEntity
       Next
    End If

'====================================================================
End Function
'====================================================================


'====================================================================
Function GetBool(Str) As Boolean
'Returns true or false
'====================================================================
    If LCase(Str) = "true" Then
        GetBool = True
    Else
        GetBool = False
    End If
    
'====================================================================
End Function
'====================================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picBack,picBack,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picBack.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picBack.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    picBack.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", picBack.BackColor, &H8000000F)
End Sub
