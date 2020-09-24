<%on error resume next%>

<% 
Dim iPageCount,FromStr

Set adoConnection = server.CreateObject("ADODB.Connection")
Set adoRecordset = server.CreateObject("ADODB.Recordset")

if trim(Request.QueryString("Item"))="" then
	Response.Redirect  Request.QueryString("From")
end if


FromStr=Request.querystring("From")
ItemStr=Request.QueryString("Item")

if FromStr="" then fromstr="list.asp"


adoConnection.Open (VAR_CONSTRING)

adoRecordset.ActiveConnection = AdoConnection
adoRecordset.PageSize=1  

ItemSQL = "Select * From [" & var_table & "] where " & Request.QueryString("Item") 

call adoRecordset.Open(ItemSQL)

iPageCount=adoRecordset.PageCount 

if iPageCount=0 then
	Response.Redirect "list.asp?" & Request.QueryString("From")
end if


%>

$SAFECODE

<p>
An Asterisk (<font color="#FF0000">*</font>) Indicates That The Field Is Required</p>

<form method="POST" action="updatecheck.asp">
<table border="0" cellpadding="2" cellspacing="2" width="100%">

<!--$Definition -->
</font>
  <tr>
    <td width="30%" height="20">&nbsp;</td>
    <td width="70%" height="20">&nbsp;</td>
  </tr>
  
</table>

<input type="submit" value="Update Changes" name="B1">   <input type="reset" value="Reset Form" name="B2"></p>
<p>
<a href="<%=FromStr%>"><< RETURN TO THE LIST </a></p>

<input type="hidden" value="<%=FromStr%>" name="FromString">
<input type="hidden" value="<%=ItemStr%>" name="ItemString">
</form>

