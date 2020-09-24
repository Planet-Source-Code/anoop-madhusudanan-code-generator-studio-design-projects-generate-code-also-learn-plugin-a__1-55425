VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchemaXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'====================================================================
'XML
'====================================================================
'XML related functions
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm

Const XS = "<"
Const XE = ">"
Const XEND = "</"

'====================================================================
Private Function StripNull(mStr As String)
'====================================================================

StripNull = ""
StripNull = Trim(Replace(mStr, Chr$(9), ""))

End Function

'====================================================================
Public Function RetXML(Param, Value, Optional DoFormat As Boolean = True)
'====================================================================
'Return properly formatted XML string

Dim XMLValue As Variant, XMLName As String
On Error Resume Next

XMLName = Param
XMLValue = Value

If DoFormat = True Then
XMLValue = Replace(XMLValue, "&", "&amp;")
XMLValue = Replace(XMLValue, "<", "&lt;")
XMLValue = Replace(XMLValue, ">", "&gt;")
XMLValue = Replace(XMLValue, """", "&quot;")
XMLValue = Replace(XMLValue, "'", "&apos;")
End If

    Select Case VarType(XMLValue)
    Case vbByte, vbInteger, vbSingle, vbDouble, vbDecimal, vbBoolean, vbLong, vbCurrency
        RetXML = XS + XMLName + XE + Trim(Str(XMLValue)) + XEND + XMLName + XE
    Case vbString, vbVariant, vbDate
        RetXML = XS + XMLName + XE + XMLValue + XEND + XMLName + XE
    Case Else
        RetXML = ""
    End Select

End Function

'====================================================================
Public Function FileToTree(strFile, tvwXMl As Object) As Boolean
'Main Private Function for loading a schema to a file
'====================================================================
On Error GoTo ErrorHandler
    Dim oXml As MSXML.XMLDocument
    Dim oroot As Object
    tvwXMl.Nodes.Clear
    Set oXml = New MSXML.XMLDocument
    oXml.url = strFile
    Set oroot = oXml.root
    BuildTree oroot, "", tvwXMl
    
FileToTree = True
Exit Function
ErrorHandler:
FileToTree = False

End Function


'====================================================================
Public Function FileToSchema(strFile, Sch As Schema) As Boolean
'Main Private Function for loading a schema to a file
'====================================================================
On Error GoTo ErrorHandler

Dim fXml As New frmXml
fXml.Hide
Dim tvwXMl As TreeView
Set tvwXMl = fXml.tvMain


    Dim oXml As MSXML.XMLDocument
    Dim oroot As Object
    tvwXMl.Nodes.Clear
    Set oXml = New MSXML.XMLDocument
    oXml.url = strFile
    Set oroot = oXml.root
    BuildTree oroot, "", tvwXMl
    TreeToSchema tvwXMl, Sch
    
    
FileToSchema = True
On Error Resume Next
Unload fXml
Set fXml = Nothing
Exit Function

ErrorHandler:

FileToSchema = False
On Error Resume Next
Set fXml = Nothing

End Function


'=====================================================================================
'FUNCTIONS FOR XML PARSING
'=====================================================================================


'=====================================================================================
Private Function TreeToSchema(Tv As Object, Sch As Schema)
'=====================================================================================
'Loads a filled treeview to the schema

'It is important to note that the schema is not at all loading
'an entities interface. Instead, it just puts the Interface chunk
'of an entity in the Interface property of the Entity object.

'From the Engine, it is necessary to parse it separately.

Dim EFlag As Boolean, RFlag As Boolean, FFlag As Boolean
Dim Nd As Node


Dim thisEntity As Entity
Dim thisRelation As Relation
Dim thisField As Field

For Each Nd In Tv.Nodes

Select Case VBA.LCase(Nd.Text)
    Case "entity"
        EFlag = True
        FFlag = False
        RFlag = False
        Set thisEntity = New Entity
        Sch.SchemaEntities.Add thisEntity
    Case "relation"
        EFlag = False
        FFlag = False
        RFlag = True
        Set thisRelation = New Relation
        Sch.SchemaRelations.Add thisRelation
    Case "field"
        If VBA.LCase(NodeParent(Nd)) = "relation" Then GoTo ParseOther
        FFlag = True
        Set thisField = New Field
        thisEntity.Fields.Add thisField
    Case "attributes"
        Dim nIndex
       
       If Nd.children Then
            nIndex = Nd.Child.FirstSibling.Index
               ' Place FirstSibling's text & linefeed in string variable.
               thisField.FieldAttributes.Add Nd.Child.FirstSibling.Text, Nd.Child.FirstSibling.Tag
               While nIndex <> Nd.Child.LastSibling.Index
                  'Set n to the next node's index.
                  nIndex = Tv.Nodes(nIndex).Next.Index
                  thisField.FieldAttributes.Add Tv.Nodes(nIndex).Text, Tv.Nodes(nIndex).Tag
               Wend
        End If
               
    Case "interface"
   
    Case Else
            
ParseOther:
            Dim Param, Value
            
            Value = ""
            Param = ""
            
            Param = Nd.Text
            Value = Nd.Tag
            
            On Error Resume Next
            
            Select Case VBA.LCase(Param)
                Case "left", "top", "width", "height"
                If EFlag = True Then thisEntity.EntityAttributes.Add Param, Value
                
                Case "state"
                    If EFlag = True Then thisEntity.EntityAttributes.Add "State", Value
            
                Case "primary"
                    If FFlag = True Then
                        thisField.FieldHeaderAttributes.Add "primary", Value
                    End If
                        
                
                Case "display"
                    If FFlag = True Then
                        thisField.FieldHeaderAttributes.Add "display", Value
                    End If
                
                
                Case "table", "field", "foreigntable", "foreignfield", "displayfield", "type"
                     If RFlag = True Then thisRelation.RelationAttributes.Add Param, Value
                     
                Case "name"
                    
                    If FFlag = True Then
                        thisField.FieldHeaderAttributes.Add "Name", Value
                    ElseIf EFlag = True Then
                        thisEntity.EntityAttributes.Add "Name", Value
                    Else
                        Sch.SchemaAttributes.Add "Name", Value
                    End If
            End Select
    
End Select

Next




'=====================================================================================
End Function
'=====================================================================================


'=====================================================================================
Private Sub BuildTree(oElement As IXMLElement, sParentID As String, tvwXMl As TreeView)
'=====================================================================================


Dim sNodeLabel As String
Dim sNodeKey As String
Dim oChild As IXMLElement2

Static iParentId As Integer

'This is used to generate sequential unique IDs
'First check element type.

'If its an element, we add it as a node
'If its just text, we add it as an extension to the parent's text

Dim It, ItC, oAttrib

If oElement.Type = XMLELEMTYPE_ELEMENT Then
    
    sNodeKey = "Node-" & CStr(iParentId + 1)
   
    If (Not oElement.Parent Is Nothing) And (sParentID <> "") Then
        Set It = tvwXMl.Nodes.Add(sParentID, tvwChild, sNodeKey)
    Else
        Set It = tvwXMl.Nodes.Add(, , sNodeKey)
    End If
    
    It.Text = oElement.tagName
    iParentId = iParentId + 1
    
    If Not (oElement.Attributes Is Nothing) Then
        For Each oAttrib In oElement.Attributes
          Set ItC = tvwXMl.Nodes.Add(sNodeKey, tvwChild, , oAttrib.Name)
          ItC.Tag = oAttrib.Value
        Next oAttrib
    End If
    
    
    If Not (oElement.children Is Nothing) Then
        For Each oChild In oElement.children
            BuildTree oChild, sNodeKey, tvwXMl
        Next oChild
    End If
    
ElseIf oElement.Type = XMLELEMTYPE_TEXT And sParentID <> "" Then
    tvwXMl.Nodes(sParentID).Tag = oElement.Text
End If


End Sub

'=====================================================================================
Private Function BuildNodeLabel(oElement As IXMLElement, It As Node, iParentId) As String
'=====================================================================================

Dim oAttrib As IXMLAttribute
End Function

'=====================================================================================
Public Function XMLDeclaration() As String
'=====================================================================================

  XMLDeclaration = "<?xml version=""1.0""?>"
End Function


'=====================================================================================
Public Function XMLFormat(XMLValue)
'=====================================================================================

On Error Resume Next
'Formats XML string's special chars
XMLValue = Replace(XMLValue, "&", "&amp;")
XMLValue = Replace(XMLValue, "<", "&lt;")
XMLValue = Replace(XMLValue, ">", "&gt;")
XMLValue = Replace(XMLValue, """", "&quot;")
XMLValue = Replace(XMLValue, "'", "&apos;")
XMLFormat = XMLValue
End Function

'=====================================================================================
Public Function XMLReFormat(XMLValue)
'=====================================================================================

On Error Resume Next
'Reformat XML strings special chars
XMLValue = Replace(XMLValue, "&amp;", "&")
XMLValue = Replace(XMLValue, "&lt;", "<")
XMLValue = Replace(XMLValue, "&gt;", ">")
XMLValue = Replace(XMLValue, "&quot;", """")
XMLValue = Replace(XMLValue, "&apos;", "'")
XMLReFormat = XMLValue
End Function


Public Function XMLParse(ByVal XMLSource As String, _
    ByVal XMLName As String, _
    Optional Instance As Integer = 1, _
    Optional Default As Variant = "") As String

    Dim X As Integer, Y As Integer, XMLStart As Integer, XMLTag As String, c As String
    Dim XMLTagEnd As String
    Dim XMLMatch As Integer, XMLEnd As Integer, XMLLength As Integer
    
    XMLLength = Len(XMLSource)
    XMLTag = XS + XMLName + XE
    XMLTagEnd = XEND + XMLName + XE
    
    '*** Find the start of the requested intstance...
    XMLStart = 1
    For X = 1 To Instance
        Y = InStr(XMLStart, VBA.LCase(XMLSource), VBA.LCase(XMLTag))
        If Y >= XMLStart Then
            XMLStart = Y + Len(XMLTag)
        Else
            XMLParse = Default
            Exit Function
        End If
    Next
    
    '*** Find the end of the instance...
    XMLEnd = XMLStart
    XMLMatch = 1
    Do Until XMLMatch = 0
        c = Mid(XMLSource, XMLEnd, Len(XMLTagEnd))
        If c = XMLTagEnd Then
            XMLMatch = XMLMatch - 1
        ElseIf Left(c, 1) = XS Then
            XMLMatch = XMLMatch + 1
        End If
        XMLEnd = XMLEnd + 1
        If XMLEnd = XMLLength Then
            XMLParse = Default
            Exit Function
        End If
    Loop
    
    XMLParse = Trim(Mid(XMLSource, XMLStart, XMLEnd - XMLStart - 1))

End Function

'=====================================================================================
Private Function NodeParent(Nd As Node)
'=====================================================================================

NodeParent = ""
On Error Resume Next
NodeParent = Nd.Parent.Text
End Function

'=====================================================================================
Private Function GetNodeChildren(Nd As Node, Tv As Object, AChunk As String)
'=====================================================================================

Dim NData As String, nIndex
NData = ""


       
       If Nd.children Then
            nIndex = Nd.Child.FirstSibling.Index
               ' Place FirstSibling's text & linefeed in string variable.
                  
                  AChunk = AChunk & "<" & Nd.Child.FirstSibling.Text & ">"
                  GetNodeChildren Nd.Child.FirstSibling, Tv, AChunk
                  AChunk = AChunk & "</" & Nd.Child.FirstSibling.Text & ">" & vbCrLf
               
               While nIndex <> Nd.Child.LastSibling.Index
                  'Set n to the next node's index.
                  AChunk = AChunk & "<" & Tv.Nodes(nIndex).Next.Text & ">"
                  GetNodeChildren Tv.Nodes(nIndex).Next, Tv, AChunk
                  AChunk = AChunk & "</" & Tv.Nodes(nIndex).Next.Text & ">" & vbCrLf
                  nIndex = Tv.Nodes(nIndex).Next.Index
               Wend
       Else
        AChunk = AChunk & vbCrLf & Nd.Tag & vbCrLf
       End If
       
End Function




