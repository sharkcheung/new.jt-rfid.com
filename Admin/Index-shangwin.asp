<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Index.asp
'文件用途：后台管理首页
'版权所有：深圳企帮
'==========================================
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title>后台管理--<%=FkSystemName%><%=FkSystemVersion%>--企帮网络荣誉出品</title>
<link href="Css/Style.css" rel="stylesheet" type="text/css" />
<SCRIPT type="text/javascript" src="../js/jquery-1.7.2.min.js"></SCRIPT>
<script type="text/javascript" src="../Js/jquery.form.min.js"></script>
<script type="text/javascript" src="../Js/function.js"></script>
<%if Bianjiqi="kinediter" then %>
<script type="text/javascript" charset="utf-8" src="/admin/dkidtioenr/kindeditor.js"></script>
<script type="text/javascript" charset="utf-8" src="/admin/dkidtioenr/lang/zh_CN.js"></script>
<%else%>
<script type="text/javascript" src="../Js/xheditor-zh-cn.min.js"></script>
<%end If%>
<!-- ymPrompt组件 -->
<script type="text/javascript" src="winskin/ymPrompt.js"></script>
<link rel="stylesheet" type="text/css" href="winskin/qq/ymPrompt.css" /> 
<!-- ymPrompt组件 -->
<script type="text/javascript">
$(document).ready(function(){
<%
If Login=False Then
	Response.Write("  tan3(""登录状态失效，请重新登录！"");")
	'Response.Write("	ShowBox(""Login.asp?Type=1&name=admin"");")
	If FKFun.GetAdminDir()="admin" Then
		'Response.Write("	alert(""系统检测到您的管理目录是默认的admin，这样不利于系统安全！\n\n建议：目录名设为6位以上、尽量复杂一些！"");")
	End If
Else
	Response.Write("	SetRContent(""UserInfo"",""Get.asp?Type=4"");")
	Response.Write("	SetRContent(""Nav"",""Get.asp?Type=1"");")
	Response.Write("	SetRContent(""MainLeft"",""Get.asp?Type=2"");")
	'Response.Write("	SetRContent(""MainRight"",""Get.asp?Type=3"");")
	Response.Write("	SetRContent(""MainRight"",""Module.asp?Type=1&MenuId=1"");")
End If
%>
	PageReSize();
});
</script>
</head>

<body oncontextmenu="return false">
<div id="AllBox">
<div id="Bodys" style="width:100%">
    <div id="PageTop">
        <div id="Top"  style="display:none">
            <div id="Logo"><a href="http://www.qebang.cn/" target="_blank" title="深圳企帮"><img src="Images/FKLogo.gif" width="136" height="32" alt="企帮LOGO" /></a></div>
            
            <div class="Cal"></div>
        </div>
        <div id="Nav">菜单
        </div><div id="UserInfo"><a href="javascript:void(0);" onClick="ShowBox('Login.asp?Type=1');" title="请您先登录！">请您先登录！</a></div>
    </div>
    <div id="PageMain">
        <div id="MainLeft">
        </div>
        <div id="MainRight">
        </div>
        <div class="Cal"></div>
    </div>
    <div id="Boxs" style="display:none">
        <div id="BoxsContent">
            <div id="BoxContent">
            </div>
        </div>
        <div id="AlphaBox" onClick="$('select').show();$('#Boxs').hide()"></div>
    </div>
</div>
</div>
</body>
</html>
<!--#Include File="../Code.asp"-->