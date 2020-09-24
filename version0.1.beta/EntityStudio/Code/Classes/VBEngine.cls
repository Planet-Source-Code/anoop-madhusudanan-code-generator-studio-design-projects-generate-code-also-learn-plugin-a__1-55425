VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================
Public Function Create(SchemaFile)
'Provide the schema file name as input
'====================================================================

Dim thisEnt As Entity
Dim thisRel As Relation
Dim thisFld As Field
Dim thisAttrib As Attrib


        For Each thisAttrib In Sch.SchemaAttributes
        Next
    
        
        For Each thisEnt In Sch.SchemaEntities
        Next
        
        For Each thisRel In Sch.SchemaRelations
        Next



'====================================================================
End Function
'====================================================================




