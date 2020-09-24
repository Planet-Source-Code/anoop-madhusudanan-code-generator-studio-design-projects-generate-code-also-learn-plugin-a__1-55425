VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAOImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim objDAO As New dao.DBEngine
Dim objDatabase As dao.Database
Dim objTable As dao.TableDef
Dim objRel As dao.Relation
Dim objField As dao.Field
Dim objProperty As dao.Property
Dim objIndex As dao.Index
Dim objIndexField As dao.Field

Const MAXATT = 15


Dim Attributes() As String


Function InitAttribs()

ReDim Attributes(MAXATT)


Attributes(0) = "Type"
Attributes(1) = "OrdinalPosition"
Attributes(2) = "Size"
Attributes(3) = "DefaultValue"
Attributes(4) = "ValidationRule"
Attributes(5) = "ValidationText"
Attributes(6) = "Required"
Attributes(7) = "AllowZeroLength"
Attributes(8) = "DecimalPlaces"
Attributes(9) = "Format"
Attributes(10) = "Caption"
Attributes(11) = "Min"
Attributes(12) = "Max"
Attributes(13) = "Lookup"
Attributes(14) = "LookupQuery"



End Function

'====================================================================
Function DBToSchema(DBPath As String, Sch As SchemaModel.Schema)
'Loads the database to schema
'====================================================================

Dim thisEntity As SchemaModel.Entity
Dim thisFld As SchemaModel.Field
Dim thisRel As SchemaModel.Relation

Dim daoProp As dao.Property


InitAttribs

Dim Primary As String, Keys() As String



If Trim(DBPath) = "" Then GoTo FoundError



'On Error GoTo FoundError

'Open The Database

Set objDatabase = objDAO.OpenDatabase(DBPath)

Sch.SchemaAttributes.Add "Name", objDatabase.Name


'Adding The Tables

'Loading Tables

For Each objTable In objDatabase.TableDefs
    
        If Mid(objTable.Name, 1, 4) <> "MSys" Then
            
            
            Set thisEntity = New SchemaModel.Entity
            
            
            thisEntity.EntityAttributes.Add "Name", objTable.Name
            
            Primary = ""
            'On Error GoTo Nokey
            For Each objIndex In objTable.Indexes
                For Each objIndexField In objIndex.Fields
                    If objIndex.Primary = True Then
                        'primary key located for table
                      Primary = Primary & objIndexField.Name & vbCrLf
                    End If
                Next
            Next
                
Nokey:
        
        
        If Primary <> "" Then
            Keys = Split(Primary, vbCrLf)
        Else
            Erase Keys
        End If
        
        
        
        For Each objField In objTable.Fields
        
        Set thisFld = New SchemaModel.Field
        thisEntity.Fields.Add thisFld

        On Error Resume Next
                
                thisFld.FieldHeaderAttributes.Add "Name", objField.Name
                thisFld.FieldHeaderAttributes.Add "Primary", "False"
                
                Dim L
                For L = LBound(Keys()) To UBound(Keys())
                    If LCase(objField.Name) = LCase(Keys(L)) Then
                        thisFld.FieldHeaderAttributes("Primary").AttribValue = "True"
                    End If
                Next L
                
                If thisFld.FieldHeaderAttributes("Primary").AttribValue = "True" Then
                    thisFld.FieldHeaderAttributes.Add "Display", "False"
                Else
                    thisFld.FieldHeaderAttributes.Add "Display", "True"
                End If
                
                
                For L = 0 To MAXATT
                        If LCase(objField.Properties(Attributes(L)).Name) = "type" Then
                             thisFld.FieldAttributes.Add "Type", TypeToString(objField.Properties("Type").Value, objField.Properties("Attributes").Value)
                        Else
                            On Error Resume Next
                             thisFld.FieldAttributes.Add objField.Properties(Attributes(L)).Name, CStr(objField.Properties(Attributes(L)).Value)
                        End If
                Next
                
        
        Next
        
        Sch.SchemaEntities.Add thisEntity
        
       End If
        
Next
                
       
        
'End Loading Tables

'Loading Relations
Dim Attrib


For Each objRel In objDatabase.Relations
    
   
        For Each objField In objRel.Fields
        
                Set thisRel = New SchemaModel.Relation
                
                If objRel.Attributes = 2 Then
                    thisRel.RelationAttributes.Add "Type", 0
                Else
                    thisRel.RelationAttributes.Add "Type", 1
                End If
            
                thisRel.RelationAttributes.Add "Table", objRel.Table
                thisRel.RelationAttributes.Add "Field", objField.Name
                thisRel.RelationAttributes.Add "ForeignTable", objRel.ForeignTable
                thisRel.RelationAttributes.Add "ForeignField", objField.ForeignName
                thisRel.RelationAttributes.Add "DisplayField", GetDisplayField(objRel.Table, Sch)
        Next
        
        

        Sch.SchemaRelations.Add thisRel

Next


objDatabase.Close
Set objDatabase = Nothing

DBToSchema = True

Exit Function

FoundError:
'MsgBox Err.Description

DBToSchema = False

End Function


Private Function TypeToString(Datatype, Attrib)
'Data Type To String

Dim Ret
On Error Resume Next


Ret = "Text"

    Select Case Datatype
        Case 10
            Ret = "Text"
        Case 100
            Ret = "Relation"
        Case 13
            Select Case Attrib
                Case 2
                    Ret = "Memo"
                Case 32779
                    Ret = "Hyperlink"
            End Select
            
        Case 12
            Select Case Attrib
                Case 2
                    Ret = "Memo"
                Case 32779
                    Ret = "Hyperlink"
            End Select
            
        Case 3, 4, 2, 6, 7, 15, 20
            Select Case Attrib
                Case 17
                    Ret = "AutoNumber"
                Case Else
                    Ret = "Number"
            End Select
        
        Case 8
             Ret = "DateTime"
        Case 5
             Ret = "Currency"
        Case 1
             Ret = "YesNo"
        Case 11
             Ret = "OLE"
    End Select
TypeToString = Ret

End Function


Public Function GetDisplayField(TableName, Sch As SchemaModel.Schema)
        
        Dim thisEnt As Entity
        Dim thisFld As SchemaModel.Field
        
        GetDisplayField = ""
        
        For Each thisEnt In Sch.SchemaEntities
            If LCase(thisEnt.EntityAttributes("name")) = LCase(TableName) Then
                For Each thisFld In thisEnt.Fields
                    If LCase(thisFld.FieldAttributes("type")) = "text" Then
                        GetDisplayField = thisFld.FieldHeaderAttributes("name")
                        Exit Function
                    End If
                Next
            End If
        Next
        
End Function
