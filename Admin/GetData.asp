<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名GetData.asp
'文件用途站点统计账号和客服账号获取
'版权所有企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System1") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'获取参数
Dim t,url,objXML,strd,rmtText,arrRmtdata
t=trim(Request.QueryString("act"))
strd=request.ServerVariables("HTTP_HOST")
url="http://win.qebang.net/shangwin/gongnenginfo-sy.asp?d="&strd&"&act="&t
Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
with objXML
	.open "GET",url,false
	.send(null)
	rmtText= .responseText
end with
Set objXML=nothing
if instr(rmtText,"{$$}")>0 then
	arrRmtdata=split(rmtText,"{$$}")
	select case t
	case "kfid" 
		Call FKFso.FsoLineWriteVer("../Inc/Site.asp",52,"KfUrl="""&arrRmtdata(1)&"""") 
	case "tjid" 
		Call FKFso.FsoLineWriteVer("../Inc/Site.asp",53,"TjUrl="""&arrRmtdata(1)&"""")
	end select
	response.Write arrRmtdata(0)
Else
	response.write 0
end if
%><!--#Include File="../Code.asp"-->