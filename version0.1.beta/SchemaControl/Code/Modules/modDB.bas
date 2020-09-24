Attribute VB_Name = "modDAOImport"
Option Explicit

'====================================================================
'SCHEMADB
'====================================================================
'Functions For Loading DAO Database To Schema Control
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm


Dim objDAO As New DAO.DBEngine
Dim objDatabase As DAO.Database
Dim objTable As DAO.TableDef
Dim objRel As DAO.Relation
Dim objField As DAO.Field
Dim objProperty As DAO.Property
Dim objIndex As DAO.Index
Dim objIndexField As DAO.Field



'====================================================================
Function DBToSchema(DBPath As String)
'Loads the database to schema
'====================================================================

'Dim I, K, L, r, Primary As String, Keys() As String
'
'
'I = 0: K = 0: L = 0: r = 0
'
'
'If Trim(DBPath) = "" Then GoTo FoundError
'
'
'
'On Error GoTo FoundError
'
''Open The Database
'
'Set objDatabase = objDAO.OpenDatabase(DBPath)
'SH.Name = objDatabase.Name
'
'
''Adding The Tables
'
''Loading Tables
'
'For Each objTable In objDatabase.TableDefs
'
'        If Mid(objTable.Name, 1, 4) <> "MSys" Then
'
'            SH.Entities(I).Name = objTable.Name
'
'            Primary = ""
'            On Error GoTo Nokey
'            For Each objIndex In objTable.Indexes
'                For Each objIndexField In objIndex.Fields
'                    If objIndex.Primary = True Then
'                        'primary key located for table
'                      Primary = Primary & objIndexField.Name & vbCrLf
'                    End If
'                Next
'            Next
'
'Nokey:
'
'        K = 0
'
'        If Primary <> "" Then
'        Keys = Split(Primary, vbCrLf)
'        Else
'        Erase Keys
'        End If
'
'
'        For Each objField In objTable.Fields
'        On Error Resume Next
'                SH.Entities(I).Fields(K).Name = objField.Name
'
'                LoadProps SH.Entities(I).Fields(K), objField
'
'                SH.Entities(I).Fields(K).Primary = False
'
'                For L = LBound(Keys()) To UBound(Keys())
'
'                If LCase(objField.Name) = LCase(Keys(L)) Then
'                    SH.Entities(I).Fields(K).Primary = True
'                End If
'
'                SH.Entities(I).Fields(K).Display = Not SH.Entities(I).Fields(K).Primary
'
'                Next L
'
'
'                K = K + 1
'                SH.Entities(I).FieldCount = K
'
'       Next
'
'
'        I = I + 1
'        SH.EntCount = I
'        End If
'
'
'Next
''End Loading Tables
'
''Loading Relations
'Dim attrib
'
'For Each objRel In objDatabase.Relations
'
'
'    If objRel.Attributes = 2 Then
'    SH.Relations(r).Type = 0
'    Else
'    SH.Relations(r).Type = 1
'    End If
'
'
'
'For Each objField In objRel.Fields
'
'    SH.Relations(r).FromTable = objRel.Table
'    SH.Relations(r).FromField = objField.Name
'    SH.Relations(r).ToTable = objRel.ForeignTable
'    SH.Relations(r).ToField = objField.ForeignName
'    r = r + 1
'    SH.RelCount = r
'Next
'
'Next
'
'
'objDatabase.Close
'Set objDatabase = Nothing
'
'DBToSchema = True
'
'Exit Function
'
'FoundError:
''MsgBox Err.Description
'
'DBToSchema = False

End Function

Sub LoadPropb(MyFld As Field, PropBlock As String)
'Loads the property from a string block to the field
Dim Props() As String, I, K
Dim Param, Value
On Error Resume Next

Props = Split(PropBlock, vbCrLf)

For I = LBound(Props()) To UBound(Props())
    Param = ""
    Value = ""
    
    GetValue Props(I), Param, Value
    
    For K = 0 To FldPropCount
        If LCase(FieldProperty(K)) = LCase(Param) Then
            MyFld.Properties(K) = Value
        End If
    Next
    
    
Next
    
End Sub


Sub LoadProps(MyFld As Field, DbFld As DAO.Field)
'Dump a property to our schema

    Dim I
    Debug.Print DbFld.Properties("Name").Value & "-" & DbFld.Properties("Type").Value
    For I = 0 To FldPropCount
    On Error Resume Next
        If LCase(FieldProperty(I)) = "type" Then
            MyFld.Properties(I) = DTTS(DbFld.Properties("Type").Value, DbFld.Properties("Attributes").Value)
        Else
            MyFld.Properties(I) = DbFld.Properties(FieldProperty(I)).Value
        End If
    Next
    
End Sub


Sub InitArrays()
'Initialize the Arrays

'The difference from the same sub in SCE->modFun is
'there the field property starts with Attribute. Here,
'I'm merging Attribute and type to a single Type property
'
'See DTTS, and STDT functions


FieldProperty(0) = "Type"
FieldProperty(1) = "OrdinalPosition"
FieldProperty(2) = "Size"
FieldProperty(3) = "DefaultValue"
FieldProperty(4) = "ValidationRule"
FieldProperty(5) = "ValidationText"
FieldProperty(6) = "Required"
FieldProperty(7) = "AllowZeroLength"
FieldProperty(8) = "DecimalPlaces"
FieldProperty(9) = "Format"
FieldProperty(10) = "Caption"
FieldProperty(11) = "Min"
FieldProperty(12) = "Max"
FieldProperty(13) = "Lookup"
FieldProperty(14) = "LookupQuery"


FldPropCount = 16

Dim I, K, L

Erase SH.Entities
Erase SH.Relations

ReDim SH.Entities(MAXENT)
ReDim SH.Relations(MAXREL)



For I = 0 To MAXENT
    
    ReDim SH.Entities(I).Fields(MAXFLD)
    
    For K = 0 To MAXFLD
        SH.Entities(I).Interface = ""
        Erase SH.Entities(I).Fields(K).Properties
        ReDim SH.Entities(I).Fields(K).Properties(20)
    Next K
Next I

SH.EntCount = 0
SH.RelCount = 0


End Sub


Public Function GetValue(Str As String, Param, Value)
'Gets value from string. eg. font=arial. then returns arial

  Dim I, Pos

    On Error Resume Next
      Pos = InStr(1, Str, "=")
      If Pos = 0 Then
          Param = Str
          Value = ""
        Else
          Param = Trim$(Left$(Str, Pos - 1))
          Value = Trim$(Right$(Str, Len(Str) - Pos))
      End If
End Function

Public Function DTTS(Datatype, attrib)
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
            Select Case attrib
                Case 2
                    Ret = "Memo"
                Case 32779
                    Ret = "Hyperlink"
            End Select
            
        Case 12
            Select Case attrib
                Case 2
                    Ret = "Memo"
                Case 32779
                    Ret = "Hyperlink"
            End Select
            
        Case 3, 4, 2, 6, 7, 15, 20
            Select Case attrib
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
DTTS = Ret

End Function



