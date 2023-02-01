<!--#Include File="AdminCheck.asp"-->
<%
dim apid
apid=request.queryString("apid")
	
	Sqlstr="Select * From [API] where id="&apid
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
%>
<div id="BoxContents" style="width:98%;">
		<iframe id="headlogin" marginwidth="0" marginheight="0" src="<%=rs("APIurl")%>" frameborder="0" width="100%" scrolling="no" height="450" onload="this.height=this.contentWindow.document.body.scrollHeight" name="I2">
							</iframe></div>

<%	
	end if
	rs.close
%>
<!--#Include File="../Code.asp"-->