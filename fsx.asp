<% 
'防恶意点击
Dim URL,url2,zzfromip,selfhost
URL=Request.ServerVariables("Http_REFERER") 
URL2=replace(url,"www.","")

if instr(URL,"baidu")>0 or instr(URL,"google")>0 or instr(URL,"soso")>0 then 

	If DateDiff("s",Request.Cookies("oesun")("vitistime"),Now())<5000 Then 
		selfhost=lcase(request.servervariables("HTTP_HOST")) 
		Response.Write("<div style='color:#CC0000;background:#FFFF66;border:1px #FF6600 solid;font-size:12px;padding:5px;'>对不起，本网站安装了防恶意点击系统，请不要重复点击竞价链接访问本站，可以直接在浏览器地址栏中输入本站域名："&selfhost&"直接访问，不便之处请见谅，谢谢！</div>") 
		zzfromip=Request.ServerVariables("REMOTE_ADDR") 
		'Response.Write("。来源为：")
		'Response.Write(""&URL&"") 
		Response.Write("<script language=javascript>alert('对不起，本网站安装了防恶意点击系统，请不要重复点击竞价链接访问本站，可以直接在浏览器地址栏中输入本站域名："&selfhost&"访问，不便之处请见谅，谢谢！');while(true){window.history.back(-1)};</script>")
		Response.Write("<meta http-equiv=""refresh"" content=""1;URL="&URL&""">") 
		'Response.Write("<script language=javascript>while(true){window.history.back(-1)};</script>")
		Response.End
	else
		'Response.Write("正常访问，未阻击！") 
	End IF 
	
	Response.Cookies("oesun")("vitistime")=Now() 
end if

'如果地址栏带=符号参数就跳转到首页
if instr(Request.Querystring(),"=")>0 then
	response.redirect "/"
end if
%>

