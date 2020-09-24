VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Relation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"RelationAttributes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRelationAttributes As RelationAttributes
Public InnerIndex As Long


Public Property Get RelationAttributes() As RelationAttributes
    If mvarRelationAttributes Is Nothing Then
        Set mvarRelationAttributes = New RelationAttributes
    End If

    Set RelationAttributes = mvarRelationAttributes
End Property


Public Property Set RelationAttributes(vData As RelationAttributes)
    Set mvarRelationAttributes = vData
End Property
Private Sub Class_Terminate()
    Set mvarRelationAttributes = Nothing
End Sub


Public Function DumpRelation()
Dim RetS
Dim oXml As New SchemaXML
Dim thisAttrib As Attrib
RetS = ""
     RetS = RetS & "<RELATION>" & vbCrLf
        For Each thisAttrib In Me.RelationAttributes
            RetS = RetS & oXml.RetXML(thisAttrib.AttribName, thisAttrib.AttribValue) & vbCrLf
        Next
     RetS = RetS & "</RELATION>" & vbCrLf & vbCrLf

DumpRelation = RetS
Set oXml = Nothing

End Function



