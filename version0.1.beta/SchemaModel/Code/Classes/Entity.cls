VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EntityAttributes"
Attribute VB_Ext_KEY = "Member1" ,"Fields"
Option Explicit

Private mvarFields As Fields
Private mvarEntityAttributes As EntityAttributes
Public InnerIndex As Long

Public Property Get EntityAttributes() As EntityAttributes
    If mvarEntityAttributes Is Nothing Then
        Set mvarEntityAttributes = New EntityAttributes
    End If


    Set EntityAttributes = mvarEntityAttributes
End Property


Public Property Set EntityAttributes(vData As EntityAttributes)
    Set mvarEntityAttributes = vData
End Property



Public Property Get Fields() As Fields
    If mvarFields Is Nothing Then
        Set mvarFields = New Fields
    End If


    Set Fields = mvarFields
End Property


Public Property Set Fields(vData As Fields)
    Set mvarFields = vData
End Property

Private Sub Class_Terminate()
  Set mvarEntityAttributes = Nothing
    Set mvarFields = Nothing
End Sub


'====================================================================
Public Function DumpEntity() As String
'Returns the serialized entity
'====================================================================
Dim RetS As String
Dim oXml As New SchemaXML

RetS = "<ENTITY>" & vbCrLf & "<!--Definition Of Entity -->" & vbCrLf & vbCrLf

Dim thisField As SchemaModel.Field
Dim thisAttrib As SchemaModel.Attrib

For Each thisAttrib In Me.EntityAttributes
    RetS = RetS & oXml.RetXML(thisAttrib.AttribName, thisAttrib.AttribValue) & vbCrLf
Next


For Each thisField In Me.Fields

    RetS = RetS & "<FIELD>" & vbCrLf & "<!--Definition Of Field -->" & vbCrLf & vbCrLf
    
        For Each thisAttrib In thisField.FieldHeaderAttributes
            RetS = RetS & oXml.RetXML(thisAttrib.AttribName, thisAttrib.AttribValue) & vbCrLf
        Next
    
    RetS = RetS & "<ATTRIBUTES>" & vbCrLf
    
        For Each thisAttrib In thisField.FieldAttributes
            RetS = RetS & oXml.RetXML(thisAttrib.AttribName, thisAttrib.AttribValue) & vbCrLf
        Next
    
    RetS = RetS & "</ATTRIBUTES>" & vbCrLf
    
    RetS = RetS & "</FIELD>" & vbCrLf & vbCrLf

Next

RetS = RetS & "</ENTITY>" & vbCrLf & vbCrLf

DumpEntity = RetS

Set oXml = Nothing

'====================================================================
End Function
'====================================================================

