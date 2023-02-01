<!--#Include File="Include.asp"-->
<!--#include file="../class/Cls_OnlineUpdate.asp"-->
<%
'==========================================
'文 件 名：Index.asp
'文件用途：后台管理首页
'版权所有：深圳企帮
'==========================================
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false;">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<style type="text/css">
 body { font-size:12px;font-family:tahoma;padding:0;margin:0;TEXT-ALIGN: center;}
 a,img{border:none;}
 a:visited{color:#0866ac;}
 .updateMain {
  background-color:#efefef;
  width:960px;
  height:524px;
  margin:0 atuo;
  line-height:520px;
}
 .updateMain p{width:360px;padding:30px;line-height:180%;font-size:20px;
  position:fixed; left:50%;margin:-180px 0 0 -240px; top:50%; _position:absolute; z-index:9999;
  border:1px solid #028ce7;
  background-color:#63c3f1;color:#0866ac;font-weight:bolder;}
  iframe{
	background-color:#63c3f1;
	height:100%;
	margin-bottom:5px;
	}
  </style>
</head>

<body>
<div class='updateMain'>
	
<%
function chkUpdate()
	chkUpdate=false
	If SysVersion="" then
		SysVersion = "1.0.0"
	End If
	dim updVersionUrl,returnMsg
	updVersionUrl="http://upd-files.qebang.cn/Version.asp?r="&now()
		' chkUpdate= DateDiff("d",SysVersionTime,Date())
		' exit function
	If DateDiff("d",SysVersionTime,Date())<>0 Then	
		On Error Resume Next
		Dim Http
		Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
		Http.open "GET",updVersionUrl,False
		Http.Send()
		If Http.Readystate<>4 then
			Set Http=Nothing
			exit function
		End if
		returnMsg=BytesToBSTR(Http.responseBody,"utf-8")
		Set Http=Nothing
		If Err.number<>0 then
			Err.Clear
			exit function
		End If
		if instr(returnMsg,SysVersion)<>1 then
			chkUpdate=true
		end if
	end if
end function

	'==============================
	'函 数 名：BytesToBstr
	'作    用：转换编码函数
	'参    数：字符串Body，编码Cset
	'==============================
	Private Function BytesToBstr(Body,Cset)
		Dim Objstream
		Set Objstream = Server.CreateObject("ado"&"d"&"b.st"&"re"&"am")
		Objstream.Type = 1
		Objstream.Mode =3
		Objstream.Open
		Objstream.Write body
		Objstream.Position = 0
		Objstream.Type = 2
		Objstream.Charset = Cset
		BytesToBstr = Objstream.ReadText 
		Objstream.Close
		set Objstream = nothing
	End Function
Dim upd
upd=request("upd")
If upd="update" then
	Response.Write("<p>正在更新...请勿进行其它操作<p>")
	response.flush()
	Dim objUpdate
	Set objUpdate = New Cls_oUpdate
	With objUpdate
		.UrlVersion = "http://upd-files.qebang.cn/Version.asp?r="&now()
		.UrlUpdate = "http://upd-files.qebang.cn/"
		.UpdateLocalPath = "/"
		.LogPath="admin/LogFile/"
		.LocalVersion = SysVersion
		.doUpdate
		'response.write .info
		Call FKFso.FsoLineWriteVer("../Inc/Site.asp",50,"SysVersion="""&.LastVersion&"""") 
		Call FKFso.FsoLineWriteVer("../Inc/Site.asp",51,"SysVersionTime="""&Date()&"""")
	End With  
	Set objUpdate = Nothing
	response.clear()
	Response.Write("<p><span style='color:red'>更新完成</span>.建议重新登录软件以使程序生效,<a href='index-shangwin.asp'>直接进入</a><br></p>")
else
	If chkUpdate() Then
		Response.Write("<p>系统检测到有更新<br><iframe src='http://upd-files.qebang.cn/info.html?r="&now()&"' style='bordernone;' allowtransparency=true frameborder=no  scrolling=auto></iframe><br><a href='?upd=update'><img src='http://image001.dgcloud01.qebang.cn/website/update.png'></a></p>")
	Else
		Response.redirect "index-shangwin.asp"
	End If
End If 
%>
</div>
</body>
</html>
<!--#Include File="../Code.asp"-->