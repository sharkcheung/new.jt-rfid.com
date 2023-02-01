<%
dim apiSET,Sqlstr
apiSET=1		'API开关 默认为关0 开1
'增加数据库表
		on error resume next
		rs.open "select * from API",conn,1,1
		if err then
		Sqlstr="create table API(id COUNTER CONSTRAINT PrimaryKey PRIMARY KEY,APIname text(255),APIurl text(255),APIb1 text(255),APIb2 text(255),APIb3 text(255))"
		Conn.Execute(Sqlstr)
		end if
		rs.close
		if err then
		response.write err.number
		err.clear
		response.end
		end if
		on error resume next
	Sqlstr="Select id,Ext_Form_Name From [Ext_FormModel]"
	Rs.Open Sqlstr,Conn,1,1
	if err then
	else
		If Not Rs.Eof Then
			While Not Rs.Eof
	%>
	<li><a href="javascript:void(0);" onClick="SetRContent('MainRight','PlusForm.asp?modeid=<%=rs("id")%>&type=1')">
	<span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><%=Rs("Ext_Form_Name")%></a>
	</li>
	<%
				Rs.MoveNext
			Wend
		end if
		rs.close
	end if
	
if apiSET=1 then
	Sqlstr="Select * From [API]"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		While Not Rs.Eof
%>
<li><a href="javascript:void(0);" onClick="SetRContent('MainRight','<%=Rs("APIurl")%>')">
<span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><%=Rs("APIname")%></a>
</li>
<%
			Rs.MoveNext
		Wend
	end if
	rs.close
end if

%>