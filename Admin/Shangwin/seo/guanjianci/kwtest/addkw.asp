<!--#Include File="../../../../AdminCheck.asp"-->
<%
	Response.Charset = "utf-8"
	dim keyword,s
	s=trim(request("s"))
	s=replace(s,"<b>","")
	s=replace(s,"</b>","")
	s=replace(s,"<B>","")
	s=replace(s,"</B>","")
	KeyWord=FKFso.FsoFileRead("../../../../KeyWord.dat")

	if instr(KeyWord,escape(s))>0 then
		response.write "<b>"&s&"</b> 关键词已存在！"
		set keyword=nothing
		response.end
	else
		if s<>"" then KeyWord=KeyWord&"%7c"&escape(s)
		Call FKFso.CreateFile("../../../../KeyWord.dat",KeyWord)
		Response.Write "<b>"&s&"</b> 成功添加到关键词库！"
	end if

%>
<!--#Include File="../../../../../Code.asp"-->