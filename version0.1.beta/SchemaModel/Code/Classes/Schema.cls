VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Entities"
Attribute VB_Ext_KEY = "Member1" ,"Relations"
Attribute VB_Ext_KEY = "Member2" ,"SchemaAttributes"
'====================================================================
'Schma Model. A DOM For the Schema
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================


Option Explicit

Private mvarSchemaAttributes As SchemaAttributes
Private mvarEntities As Entities
Private mvarRelations As Relations
Private Filled As Boolean

'====================================================================
Public Function LoadSchema(Filename)
'Loads the schema from a valid XML document
'====================================================================

If Filled = True Then Err.Raise 10010, , "Schema already filled"

Dim Xml As New SchemaXML
LoadSchema = Xml.FileToSchema(Filename, Me)
Set Xml = Nothing
Filled = True

'====================================================================
End Function
'====================================================================


'====================================================================
Public Function SaveSchema(Filename)
'Emits the schema to a file
'====================================================================

Dim Dumper As New SchemaDump

Dim RetStr
RetStr = Dumper.DumpSchema(Me)

On Error Resume Next
Open Filename For Output As #1
Debug.Print RetStr

If Err Then
    Err.Raise "File error. Unable to save changes. Try saving the file with a different name"
    Exit Function
End If

Print #1, RetStr
Close #1

Set Dumper = Nothing

'====================================================================
End Function
'====================================================================


'====================================================================
Public Function DumpSchema()
'Returns the schema string
'====================================================================
Dim Dumper As New SchemaDump

DumpSchema = Dumper.DumpSchema(Me)

Set Dumper = Nothing

'====================================================================
End Function
'====================================================================


'====================================================================
Public Property Get SchemaRelations() As Relations
'====================================================================
    
    If mvarRelations Is Nothing Then
        Set mvarRelations = New Relations
    End If

    Set SchemaRelations = mvarRelations
    
'====================================================================
End Property
'====================================================================


Public Property Set SchemaRelations(vData As Relations)
    Set mvarRelations = vData
End Property



Public Property Get SchemaEntities() As Entities
    If mvarEntities Is Nothing Then
        Set mvarEntities = New Entities
    End If

    Set SchemaEntities = mvarEntities
    
End Property


Public Property Set SchemaEntities(vData As Entities)
    Set mvarEntities = vData
End Property



Public Property Get SchemaAttributes() As SchemaAttributes
    If mvarSchemaAttributes Is Nothing Then
        Set mvarSchemaAttributes = New SchemaAttributes
    End If


    Set SchemaAttributes = mvarSchemaAttributes
End Property


Public Property Set SchemaAttributes(vData As SchemaAttributes)
    Set mvarSchemaAttributes = vData

End Property

Private Sub Class_Initialize()
    Filled = False
End Sub

Private Sub Class_Terminate()
  Set mvarRelations = Nothing
  Set mvarEntities = Nothing
  Set mvarSchemaAttributes = Nothing
End Sub



