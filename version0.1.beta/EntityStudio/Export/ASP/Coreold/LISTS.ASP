<%

'----------------------------------------------------------------------
'SQL Parts
Dim sqWhat,sqFrom,sqWhere,sqOrder

'SQL Command=Sum of SQL Parts
Dim sCommand

'Heading To Display
Dim sHeading
'----------------------------------------------------------------------

'Parameters for building the Where clause later      
pClear=Request.QueryString("Clear")          'Clear Previous Variables
pParam=Request.QueryString("Action")          'Parameter
pParamId=Request.QueryString("ActionId")      'ParemeterId
pFor=Request.QueryString("Text")             'Text To Search


if pSearch<>"" then
pClear="true"
pParam="search"
pFor=replace(pSearch,"'","''")
end if


'Field Map
Dim FieldMap(),FieldCount

'FieldMap(X,0)=FieldName
'FieldMap(X,1)=ForeignField
'FieldMap(X,2)=DisplayField
'FieldMap(X,3)=LinkValue
'FieldMap(X,4)=Caption
'FieldMap(X,5)=Visible

'Setting SQL Parts
%>

<!--$Definition -->

<%
'----------------------------------------------------------------------
'Now Checking Parameters
if pClear="true" then
sqWhere=""
	select case lcase(pParam)
		case "list"		
				
		case "search"

		case "latest"	
				
	end select

end if


'----------------------------------------------------------------------

%>

<%
'============================================================================
Sub DisPlayItems()
'
'This function counts
'============================================================================

Dim iPageSize       'How big our pages are
Dim iPageCount      'The number of pages we get back
Dim iPageCurrent    'The page we want to show
Dim pAction         'The Next, Prev etc
Dim mConMain        'The ADODB connection object
Dim mRst            'The ADODB recordset object
Dim iRecordsShown   'Loop controller for displaying just iPageSize records
Dim I               'Standard looping var

'We will use sCommand as the command text, later
'Now setting the sCommand parameter

if trim(session("$ENTITY_AQ_LIST"))="" or pClear="true" then
'We have to create new query
	if trim(sqWhere)<>"" then
	sCommand=sqWhat+ " " + sqFrom+ " "+sqWhere+" "+sqOrder
	else
	sCommand=sqWhat+ " "+sqFrom+ " "+sqOrder
	end if
else
	sCommand=session("$ENTITY_AQ_LIST")
end if

'Setting Article Query List
session("$ENTITY_AQ_LIST")=sCommand
'----------------------------------------------------------------------

'Current Page - iPageCurrent Parameter
If Request.QueryString("Page") = "" Then
	'For Evaluation Session's Content
	iPageCurrent = 0
Else
	iPageCurrent = CInt(Request.QueryString("Page"))
End If


'----------------------------------------------------------------------
'Now get the action if any
pAction=Request.QueryString("Action")

'----------------------------------------------------------------------
%>


<%

On error resume next

'Deals with the recordset

set mConMain=Server.CreateObject("ADODB.Connection")
set mRst=server.CreateObject("ADODB.Recordset")

sConnection=var_constring
mConMain.Open sConnection


iPageSize=10

'Setting CacheSize
mRst.CursorLocation=3 'Use Client
mRst.CacheSize=iPageSize
mRst.PageSize=iPageSize
mRst.Open sCommand,mConMain,adOpenKeyset

iPageCount=mRst.PageCount


If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1


'----------------------------------------------------------------------
'Checking actions. << < > >>
'Current Page - Checking if already in session
if trim(pAction) <> "" then
	if trim(Session("$ENTITY_CUR_PAGE"))="" then 
		Session("$ENTITY_CUR_PAGE")="1"
	end if
	iPageCurrent=cint(Session("$ENTITY_CUR_PAGE"))
End if

Select Case lcase(pAction)
	Case "first"
			iPageCurrent=1
	Case "last"		
			iPageCurrent=iPageCount
	case "next"
			iPageCurrent=iPageCurrent+1
	case "prev"		
			iPageCurrent=iPageCurrent-1
End Select

If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
If iPageCurrent < 1 Then iPageCurrent = 1

'----------------------------------------------------------------------


'No records found :-(
if iPageCount=0 then

%>
	
      	<tr>
		<td width="100%" bgcolor="black" height="20">
			<font color="#ffffff" face="verdana,arial" size="2">
			&nbsp;<b> NO <%=ucase("$ENTITY")%> FOUND</b>
			</font>
		</td>	
		</tr>

      	<tr>
			<td width="100%" height="20">
			<font face="verdana,arial" size="1"><b>
			SORRY, NO <%=ucase("$ENTITY")%> FOUND MATCHING THIS CATEGORY/CRITERIA
			</font></b>
			</td>
		</tr>
		
<%
else
'Heading
%>

      	<tr>
			<td width="100%" height="20">
			<font face="verdana,arial" size="1"><b>
			<%if ipagecount>1 then %>
			ABOUT <%=(ipageCount-1)*iPageSize%>+ <%=ucase(" $ENTITY")%> FOUND IN <%=IPageCount%> PAGE(S).
			<%elseif iPagecount=1 then %>
			<%=mRst.RecordCount  %> $ENTITY FOUND
			<%end if%>
			
			</b></font>
			</td>
		</tr>

<%


'----------------------------------------------------------------------

	'Dont forget to write the navigation bar
	if iPageCount>1 then
	Select case iPageCurrent
			case iPageCount
				mWriteNavBar true,false,iPageCount,iPageCurrent
			case 1
				mWriteNavBar false,true,iPageCount,iPageCurrent
			case else
				mWriteNavBar true,true,iPageCount,iPageCurrent
    end select
    else
    %>

<tr>
<td width="100%">

<table border="0" width="100%" bgcolor cellspacing="0" cellpadding="0" height="20">
  <tr>
          
          <td width="95%" valign="middle" bgcolor="#1997EE">
            <b><font color="#FFFFFF" face="VERDANA,ARIAL" size="1">
            
            &nbsp;ONE PAGE FOUND 
            </b>
           </td>
     </tr>
</table>
</td>
</tr>
           
    <%
       
    end if
'----------------------------------------------------------------------

	'Calling function for writing pages 1 2 3 etc
	WritePages iPageCount,iPageCurrent 
   'WriteHeader

	mRst.AbsolutePage=iPageCurrent
    Session("$ENTITY_CUR_PAGE")=cstr(iPageCurrent)
    
    %>
			<tr>
			<td width="100%">
						<table border="0">
						<tr>
						<%
						Dim TField
						For i=0 to fieldCount
						'Writing Headers
						if cbool(fieldmap(i,5))=true then
						%>
						
						<td bgcolor="#000080"><font face="verdana,arial" size="1" color="#FFFFFF"><b><%=fieldmap(i,4)%></b></font>&nbsp;&nbsp;</td>
						
						<%
						end if
						next%>
						
						</tr>

						<tr>
						<%
						For i=0 to fieldCount
						if cbool(fieldmap(i,5))=true then
						%>
						<td>&nbsp;</td>
						<%
						End If
						next%>
						</tr>
						
						<%bcolor = "#COCOCO"%>
						<%For intRecord = 1 To mrst.PageSize %>
						<tr>
						<% qString = "" %>

						<% 
						Dim PrimaryStr,RealField
						PrimaryStr=""
						

						For i=0 to fieldCount
						if cbool(fieldmap(i,6))=true then

						RealField=fieldmap(i,0)
						RealField=Replace(RealField,".","")
						RealField=Replace(RealField,"[","")
						RealField=Replace(RealField,"]","")
						RealField=Replace(RealField," ","")
						
						Dim fVal
						
						fVal=mRst(RealField)
						
						err.Clear 
						on error resume next
						fVal=cDbl(fVal)
						
						if err then
							fVal="'" & replace(fVal,"'","''") & "'"
						else
														
						end if
						
							if PrimaryStr="" then
							PrimaryStr=fieldmap(i,0) & "=" & fVal
							else
							PrimaryStr=PrimaryStr & " AND " & fieldmap(i,0) & "=" & fVal
							end if
						end if
						
						if cbool(fieldmap(i,5))=true then
						%>
						<td bgcolor="<%=bcolor%>"><font face="verdana,arial" size="1">&nbsp;<%=mRst(fieldmap(i,2))%></font></td>
						<%
						End If

						next
						
						Dim FromStr
						
						if Trim(PrimaryStr)<>"" then
						PrimaryStr="Item=" & Server.urlencode(PrimaryStr)
						FromStr="From=list.asp?" & "Page=" & iPageCurrent
						fromstr=server.URLEncode(fromStr)
						
						%>
						
						<td bgcolor="<%=bcolor%>">&nbsp;<a href="delete.asp?<%=PrimaryStr%>&<%=FromStr%>"><font face="verdana,arial" size="1" color="red"><b>Delete</b></font></a></td>
						<td bgcolor="<%=bcolor%>">&nbsp;<a href="update.asp?<%=PrimaryStr%>&<%=FromStr%>"><font face="verdana,arial" size="1" color="red"><b>Update</b></font></a></td>
						<%
						End If
						mRst.MoveNext
						If bcolor = "#COCOCO" Then
						bcolor = "#F3F3DC"
						Else
						bcolor = "#COCOCO"
						End If
						If mRst.EOF Then Exit For %>
						</tr>
						<%Next%>
						
						</table>

			</td>
			</tr>


<%
End If

'----------------------------------------------------------------------

	'Dont forget to write the navigation bar on bottom
	if iPageCount>1 then
	Select case iPageCurrent
			case iPageCount
				mWriteNavBar true,false,iPageCount,iPageCurrent
			case 1
				mWriteNavBar false,true,iPageCount,iPageCurrent
			case else
				mWriteNavBar true,true,iPageCount,iPageCurrent
    end select
    end if
'----------------------------------------------------------------------

on error resume next    
'Clean up    			
mRst.Close
mConMain.Close

Set mCommand=nothing
Set mConMain=nothing
%>

<%
'============================================================================
End Sub
'============================================================================

%>




<%  
'============================================================================
Sub WritePages(iPageCount,iPageCurrent) 
'============================================================================
%>

<tr>
<td width="100%">
<table border="0" width="100%">

<% if iPageCount>1 then %>

	<tr>
		<td width="100%" height="30" align="center"><b><font face="Verdana,Arial" size="2">Pages: 
		<%
		
		ShowStart=iPageCurrent-10
		ShowStop=iPageCurrent+10
		
		if iPageCurrent<10 then showstop=20
		if iPageCurrent>iPageCount-10 then ShowStart=ShowStart-10
		

		if ShowStop>iPageCount then ShowStop=iPageCount
		if ShowStart<1 then ShowStart=1
		
		if ShowStart>1 then
			Response.Write "<b>" + "<font size=" + "1" + "><-</b></font>"
		end if

		
		For i=ShowStart to ShowStop
			if i<>iPageCurrent then
			Response.Write " "+ "<a href=""List.asp?Page=" + cstr(i) + """>" + cstr(i) +  "</a>  "
			else
			Response.Write " " + "<font color=" + "Red" + ">[" + cstr(i) +  "]</font>  "
			End if
		next 

		if ShowStop<iPageCount then
			Response.Write "<b>" + "<font size=" + "1" + ">-></b></font>"
		end if
		
		%>
		
		</font></td>
	</tr>
<% end if%>	

</table>
</td>
</tr>


<% 
'============================================================================
End Sub 
'============================================================================
%>


<%
'============================================================================
Sub mWriteNavBar(mFirst,mLast,iPageCount,iPageCur)
'============================================================================


%>
<tr>
<td width="100%">

<table border="0" width="100%" bgcolor cellspacing="0" cellpadding="0" height="20">
  <tr>
          
          <td width="55%" valign="middle" bgcolor="#1997EE">
            <b><font color="#FFFFFF" face="VERDANA,ARIAL" size="1">
            &nbsp;PAGE <%=(iPageCur)%> OF <%=iPageCount%> PAGES &nbsp;
            </font></b>
          </td>

<td nowrap width="45%" bgcolor="#1997EE" valign="middle" align="right"><b><font color="#FFFFFF" face="VERDANA,ARIAL" size="1">

<% if mFirst=true then %>
      <a href="list.asp?Action=First"><font face="verdana,arial" size="1" color="#ffffff"> &lt;&lt; FIRST </a> |
      <a href="list.asp?Page=<%=iPageCur-1%>"><font face="verdana,arial" size="1" color="#ffffff"> &lt; PREV </a>
 <% end if %>
      <% if mLast=true and mFirst=true then Response.Write "|" %>

      <% if mLast=true then %>
      <a href="list.asp?Page=<%=iPageCur+1%>"><font face="verdana,arial" size="1" color="#ffffff"> NEXT &gt;</a> |
      <a href="list.asp?Action=Last"><font face="verdana,arial" size="1" color="#ffffff"> LAST &gt;&gt;</a>
      <% end if %>
  </font></td></tr>
</table>

</td>
<tr>

<%
'============================================================================
End Sub
'============================================================================
%>


$SAFECODE

<table border="0" cellspacing="2" cellpadding="2" width="100%">
		
<!--Start Item List -->		
		      <%DisplayItems%>
<!--End Item List -->

</table>




