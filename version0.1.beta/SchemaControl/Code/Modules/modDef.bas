Attribute VB_Name = "modStructs"
'====================================================================
'Structures
'====================================================================
'General Structures
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm


'Model Diagram
'===============
'
'Schema
'   |
'   |____ Relations
'   |
'   |____ Entities
'           |
'           |_____ Fields
'           |       |
'                   |_____ Properties


'Structs for this model



Public Type Field
'Definition of a field
    Name As String
    Primary As Boolean
    Display As Boolean
    Properties() As String
End Type

Public Type Entity
'Definition of an entity
    Name As String
    FieldCount As Integer
    Interface As String
    Fields() As Field
    X As Variant
    Y As Variant
    Width As Variant
    Height As Variant
    State As Variant
End Type

Public Type Relat
'Definition of a relation
    FromTable As String
    FromField As String
    ToTable As String
    ToField As String
    DisplayField As String
    Type As Variant
End Type

Public Type Schema
'Definition of a schema
    Name As String
    EntCount As Integer
    RelCount As Integer
    Entities() As Entity
    Relations() As Relat
End Type


Type List
'Model of a list
    Name As String
    Remark As String
    FieldCount As Integer
    Fields() As Field
    SQLFrom As String
    SQLWhere As String
    SQLOther As String
    HTMLHeader As String
    HTMLRow As String
    IfRelation As Boolean
    IfLink As Boolean
End Type


Type View
    Name As String
    FieldCount As Integer
    Fields() As Field
    SQLFrom As String
    SQLWhere As String
    SQLOther As String
End Type



'Arrays for holding property names

Public FldPropCount
Public FieldProperty(17)


'Maximum possible relations
Public Const MAXREL = 500

'Maximum possible entities
Public Const MAXENT = 500

'Maximum fields in a table
Public Const MAXFLD = 100

'The Schema Object
Public SH As Schema



