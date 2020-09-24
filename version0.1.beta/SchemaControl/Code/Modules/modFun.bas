Attribute VB_Name = "modXml"
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
Function StripNull(mStr As String)
'====================================================================

StripNull = ""
StripNull = Trim(Replace(mStr, Chr$(9), ""))

End Function

'====================================================================
Function RetXML(Param, Value, Optional DoFormat As Boolean = True)
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
Function DataToTree(Base, Data, tvwXML As TreeView) As Boolean
'Main Function for loading a schema to a file
'====================================================================
Dim strFile

On Error GoTo ErrorHandler

Data = XS & Base & XE & Data & XEND & Base & XE


strFile = App.Path & "\temp.xml"

Open App.Path & "\temp.xml" For Output As #1
Print #1, Data
Close #1


    Dim oXML As MSXML.XMLDocument
    Dim oroot As Object
    tvwXML.Nodes.Clear
    Set oXML = New MSXML.XMLDocument
    oXML.url = strFile
    Set oroot = oXML.root
    BuildTree oroot, "", tvwXML
    
DataToTree = True
Exit Function
ErrorHandler:
DataToTree = False

End Function


'====================================================================
Function FileToTree(strFile, tvwXML As TreeView) As Boolean
'Main Function for loading a schema to a file
'====================================================================
On Error GoTo ErrorHandler
    Dim oXML As MSXML.XMLDocument
    Dim oroot As Object
    tvwXML.Nodes.Clear
    Set oXML = New MSXML.XMLDocument
    oXML.url = strFile
    Set oroot = oXML.root
    BuildTree oroot, "", tvwXML
    
FileToTree = True
Exit Function
ErrorHandler:
FileToTree = False

End Function


'====================================================================
Function FileToSchema(strFile, tvwXML As TreeView) As Boolean
'Main Function for loading a schema to a file
'====================================================================
On Error GoTo ErrorHandler
    Dim oXML As MSXML.XMLDocument
    Dim oroot As Object
    tvwXML.Nodes.Clear
    Set oXML = New MSXML.XMLDocument
    oXML.url = strFile
    Set oroot = oXML.root
    BuildTree oroot, "", tvwXML
    TreeToSchema tvwXML
    
    
FileToSchema = True
Exit Function
ErrorHandler:
FileToSchema = False

End Function


'=====================================================================================
'FUNCTIONS FOR XML PARSING
'=====================================================================================


'=====================================================================================
Function TreeToSchema(Tv As TreeView)
'=====================================================================================
'Loads a filled treeview to the schema

'It is important to note that the schema is not at all loading
'an entities interface. Instead, it just puts the Interface chunk
'of an entity in the Interface property of the Entity object.

'From the Engine, it is necessary to parse it separately.


Dim EFlag As Boolean, RFlag As Boolean, FFlag As Boolean

Dim Nd As Node
Dim i, K, L, r
i = -1: K = -1: L = -1: r = -1

SH.EntCount = 0
SH.RelCount = 0

Dim AChunk As String
Dim IChunk As String

AChunk = ""
IChunk = ""

For Each Nd In Tv.Nodes

Select Case VBA.LCase(Nd.Text)
    Case "entity"
        i = i + 1
        K = -1
        SH.EntCount = i + 1
        EFlag = True
        FFlag = False
        RFlag = False
    Case "relation"
        EFlag = False
        FFlag = False
        RFlag = True
        r = r + 1
        SH.RelCount = r + 1
    Case "field"
        If VBA.LCase(NodeParent(Nd)) = "relation" Then GoTo ParseOther
        K = K + 1
        FFlag = True
        SH.Entities(i).FieldCount = K + 1
    Case "attributes"
        AChunk = ""
        Dim nIndex
       
       If Nd.children Then
        
            nIndex = Nd.Child.FirstSibling.Index
               ' Place FirstSibling's text & linefeed in string variable.
               AChunk = Nd.Child.FirstSibling.Text & "=" & Nd.Child.FirstSibling.Tag & vbCrLf
               While nIndex <> Nd.Child.LastSibling.Index
                  AChunk = AChunk & Tv.Nodes(nIndex).Next.Text & "=" & Tv.Nodes(nIndex).Next.Tag & vbCrLf
                  'Set n to the next node's index.
                  nIndex = Tv.Nodes(nIndex).Next.Index
               Wend
               
            LoadPropb SH.Entities(i).Fields(K), AChunk
        
        End If
               
    Case "interface"
        Dim Ach As String
        Ach = ""
        GetNodeChildren Nd, Tv, Ach
        SH.Entities(i).Interface = Ach
    
    Case Else
            
ParseOther:
            Dim Param, Value
            
            Value = ""
            Param = ""
            
            GetValue Nd.Text & "=" & Nd.Tag, Param, Value
            
            On Error Resume Next
            
            Select Case VBA.LCase(Param)
                Case "left"
                If EFlag = True Then SH.Entities(i).X = Val(Value)
                
                Case "top"
                If EFlag = True Then SH.Entities(i).Y = Val(Value)
                
                Case "width"
                If EFlag = True Then SH.Entities(i).Width = Val(Value)
                Case "height"
                If EFlag = True Then SH.Entities(i).Height = Val(Value)
                Case "state"
                If EFlag = True Then SH.Entities(i).State = Val(Value)
            
                Case "primary"
                If FFlag = True Then
                    If VBA.LCase(Value) = "true" Then
                        SH.Entities(i).Fields(K).Primary = True
                    Else
                        SH.Entities(i).Fields(K).Primary = False
                    End If
                End If
                    
                
                Case "display"
                If FFlag = True Then
                    If VBA.LCase(Value) = "true" Then
                        SH.Entities(i).Fields(K).Display = True
                    Else
                        SH.Entities(i).Fields(K).Display = False
                    End If
                End If
                
                Case "table"
                     If RFlag = True Then SH.Relations(r).FromTable = Value
                Case "field"
                     If RFlag = True Then SH.Relations(r).FromField = Value
                Case "foreigntable"
                     If RFlag = True Then SH.Relations(r).ToTable = Value
                Case "foreignfield"
                     If RFlag = True Then SH.Relations(r).ToField = Value
                Case "displayfield"
                     If RFlag = True Then SH.Relations(r).DisplayField = Value
                Case "type"
                     If RFlag = True Then SH.Relations(r).Type = Val(Value)
                     
                Case "name"
                    
                    If FFlag = True Then
                        SH.Entities(i).Fields(K).Name = Value
                    ElseIf EFlag = True Then
                        SH.Entities(i).Name = Value
                    Else
                        SH.Name = Value
                    End If
                    
            End Select
    
End Select

Next



'=====================================================================================
End Function
'=====================================================================================


'=====================================================================================
Private Sub BuildTree(oElement As IXMLElement, sParentID As String, tvwXML As TreeView)
'=====================================================================================


Dim sNodeLabel As String
Dim sNodeKey As String
Dim oChild As IXMLElement2

Static iParentId As Integer

'This is used to generate sequential unique IDs
'First check element type.

'If its an element, we add it as a node
'If its just text, we add it as an extension to the parent's text

Dim It, ItC

If oElement.Type = XMLELEMTYPE_ELEMENT Then
    
    sNodeKey = "Node-" & CStr(iParentId + 1)
   
    If (Not oElement.Parent Is Nothing) And (sParentID <> "") Then
        Set It = tvwXML.Nodes.Add(sParentID, tvwChild, sNodeKey)
    Else
        Set It = tvwXML.Nodes.Add(, , sNodeKey)
    End If
    
    It.Text = oElement.tagName
    iParentId = iParentId + 1
    
    If Not (oElement.Attributes Is Nothing) Then
        For Each oAttrib In oElement.Attributes
          Set ItC = tvwXML.Nodes.Add(sNodeKey, tvwChild, , oAttrib.Name)
          ItC.Tag = oAttrib.Value
        Next oAttrib
    End If
    
    
    If Not (oElement.children Is Nothing) Then
        For Each oChild In oElement.children
            BuildTree oChild, sNodeKey, tvwXML
        Next oChild
    End If
    
ElseIf oElement.Type = XMLELEMTYPE_TEXT And sParentID <> "" Then
    tvwXML.Nodes(sParentID).Tag = oElement.Text
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
Function NodeParent(Nd As Node)
'=====================================================================================

NodeParent = ""
On Error Resume Next
NodeParent = Nd.Parent.Text
End Function

'=====================================================================================
Function GetNodeChildren(Nd As Node, Tv As TreeView, AChunk As String)
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
