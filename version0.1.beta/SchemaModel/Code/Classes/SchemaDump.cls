VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchemaDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim oXml As SchemaXML


'====================================================================
Public Function DumpSchema(Sch As Schema)
'Returns the current Sch schema as xml
'====================================================================


Dim thisEnt As Entity
Dim thisRel As Relation
Dim thisFld As Field
Dim thisAttrib As Attrib

Set oXml = New SchemaXML

'Returns the schema tag file
Dim RetS, i

RetS = oXml.XMLDeclaration & vbCrLf
RetS = "<!-- " & vbCrLf & "======================================================" & vbCrLf
RetS = RetS & vbCrLf & "ERD DEFINITION FORMAT 1.0" & vbCrLf & vbCrLf
RetS = RetS & "Created At: " & Now & vbCrLf
RetS = RetS & vbCrLf & "Created By: Anoop Madhusudanan - http://www.logicmatrixonline.com" & vbCrLf & vbCrLf
RetS = RetS & "======================================================" & vbCrLf & " -->" & vbCrLf

RetS = RetS & "<SCHEMA>" & vbCrLf
        
        For Each thisAttrib In Sch.SchemaAttributes
            RetS = RetS & oXml.RetXML(thisAttrib.AttribName, thisAttrib.AttribValue) & vbCrLf
        Next
    
        
        For Each thisEnt In Sch.SchemaEntities
            RetS = RetS & thisEnt.DumpEntity() & vbCrLf
        Next
        
        For Each thisRel In Sch.SchemaRelations
            RetS = RetS & thisRel.DumpRelation() & vbCrLf
        Next
        

RetS = RetS & "</SCHEMA>" & vbCrLf

DumpSchema = RetS


Set oXml = Nothing

'====================================================================
End Function
'====================================================================





