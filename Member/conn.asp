<%
         db_date="now()"
		 db_day="'d'"
		 db_minute="'n'"
		 db_true=true
		 db_false=false
connstr2="provider=microsoft.jet.oledb.4.0;data source="&Server.MapPath(SiteData)
set connn=server.CreateObject("adodb.connection")
connn.open connstr2
%>