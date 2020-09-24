<% Function Var_SubNavbar %>

<form method="POST" Action="search.asp?action=search&clear=true">
<p align="right"><b>
    <font color="#FFFFFF" face="Verdana" size="1">    <font color="#FFFFFF" face="Verdana" size="1"> <a href="List.asp"><font color="#ffffff">BROWSE</font></a>|<a href="new.asp"><font color="#ffffff">ADD NEW</font></a> | SEARCH : </font><input type="text" name="txtSearch" size="20">&nbsp;<input type="submit" value="Go" name="B1"> <a href="search.asp?advanced=true"> <font color="#FFFFFF" face="Verdana" size="1">Advanced..</font></a></font></b></p>
</form>

<% End Function%>




<%
'============================================================
Function WriteCatCombo(Name,Field,DisplayField,Table,TheValue)
on error resume next
'============================================================
%>
<select size="1" name="<%=Name%>">

<%
Dim mConMain,mRst,sConnection,sCommand

set mConMain=Server.CreateObject("ADODB.Connection")
set mRst=server.CreateObject("ADODB.Recordset")

sConnection=VAR_CONSTRING
mConMain.Open sConnection

sCommand="SELECT [" & Field & "],[" & displayfield  & "] From [" & table & "]"

mRst.Open sCommand,mConMain,adOpenForwardOnly 

Do while not mRst.EOF

myCap=mRst.Fields(1)

if LCASE(cstr(TheValue))= LCASE(cstr(mRst.Fields(0))) then  
%>
    <option value="<%=mRst.Fields(0)%>" selected><%=myCap%></option>
<%else%>
    <option value="<%=mRst.Fields(0)%>"><%=myCap%></option>
<%
End If
  
mRst.movenext	
Loop

mRst.Close 
mConMain.Close

set mRst=nothing
set mConmain=nothing
%>
 
 </select>  
<%
end function
%>


<%
'============================================================
Function WriteLookupCombo(Name,Query,TheValue)
on error resume next
'============================================================
%>
<select size="1" name="<%=Name%>">

<%
Dim mConMain,mRst,sConnection,sCommand

set mConMain=Server.CreateObject("ADODB.Connection")
set mRst=server.CreateObject("ADODB.Recordset")

sConnection=VAR_CONSTRING
mConMain.Open sConnection

sCommand=Query

mRst.Open sCommand,mConMain,adOpenForwardOnly 

Do while not mRst.EOF

myCap=mRst.Fields(1)

if LCASE(cstr(TheValue))= LCASE(cstr(mRst.Fields(0))) then  
%>
    <option value="<%=mRst.Fields(0)%>" selected><%=myCap%></option>
<%else%>
    <option value="<%=mRst.Fields(0)%>"><%=myCap%></option>
<%
End If
  
mRst.movenext	
Loop

mRst.Close 
mConMain.Close

set mRst=nothing
set mConmain=nothing
%>
 
 </select>  
<%
end function
%>