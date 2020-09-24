<%on error resume next%>

<%
Dim FromStr
FromStr=request.querystring("From")

if FromStr="" then fromstr="list.asp"

%>


$SAFECODE


<p>
An Asterisk (<font color="#FF0000">*</font>) Indicates That The Field Is Required</p>


<form method="POST" action="newcheck.asp">
<table border="0" cellpadding="2" cellspacing="2" width="100%">
<font face="verdana" size="1">
<!--$Definition -->
</font>
  <tr>
    <td width="30%" height="20">&nbsp;</td>
    <td width="70%" height="20">&nbsp;</td>
  </tr>
</table>
<input type="submit" value="Save Changes" name="B1">   <input type="reset" value="Reset Form" name="B2"></p>
<p>
<a href="list.asp"><< RETURN TO THE LIST </a></p>

</form>
