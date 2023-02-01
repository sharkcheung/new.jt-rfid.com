<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Index.asp
'文件用途：后台管理首页
'版权所有：深圳企帮
'==========================================
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8"/>
<title>后台管理--<%=FkSystemName%><%=FkSystemVersion%>--企帮网络荣誉出品</title>

<SCRIPT type="text/javascript" src="../js/jquery-1.7.2.min.js"></SCRIPT>

<script type="text/javascript" src="../Js/jquery.form.min.js"></script>
<script type="text/javascript" src="layer/layer.js"></script>
<link href="Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
var adminpath="<%=AdminPath%>";
</script>
<%if Bianjiqi="kinediter" then %>
<script type="text/javascript" charset="utf-8" src="/<%=AdminPath%>/dkidtioenr/kindeditor.js"></script>
<script type="text/javascript" charset="utf-8" src="/<%=AdminPath%>/dkidtioenr/lang/zh_CN.js"></script>
<%elseif Bianjiqi="xheditor" then%>
<script type="text/javascript" src="../Js/xheditor-zh-cn.min.js"></script>
<%else%>
<!-- 配置文件 -->
<script type="text/javascript" src="ueditor/ueditor.config.js"></script>
<!-- 编辑器源码文件 -->
<script type="text/javascript" src="ueditor/ueditor.all.js"></script>
<!-- 实例化编辑器 -->
<%end If%>
<script type="text/javascript" src="Js/function.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$(document).on("click",".zsptlist>dt>ul>li>a>em",function(){
				$(this).toggleClass("active");
				$(this).parent("a").next("ul").slideToggle(100,function(){
					if($(this).is(":visible")){
					$(this).siblings("a").css("font-weight","bold");
					}
				if(!$(this).is(":visible")){
					$(this).siblings("a").css("font-weight","normal");
					}
					});
	})
	
<%
If Login=False Then
	Response.Write("  tan3(""登录状态失效，请重新登录！"");")
	'Response.Write("	ShowBox(""Login.asp?Type=1&name=admin"");")
	If FKFun.GetAdminDir()="admin" Then
		'Response.Write("	alert(""系统检测到您的管理目录是默认的admin，这样不利于系统安全！\n\n建议：目录名设为6位以上、尽量复杂一些！"");")
	End If
Else
	Response.Write("	SetRContent(""MainLeft"",""Get.asp?Type=2"");")
	Response.Write("	SetRContent(""MainRight"",""Module.asp?Type=1&MenuId=1"");")
End If
%>
	PageReSize();
});
</script>
<script type="text/javascript">
window.onerror = function(sMsg, sUrl, sLine) {
           var strlog="错误信息：" + sMsg + "\r\n";
           strlog+="出错文件：" + sUrl + "\r\n";
           strlog+="出错行号：" + sLine + "\r\n";
           // alert(strlog);
           return true;
    }
</script></head>

<body>
<div class="menunav">
	<div class="center">
		<a href="javascript:void(0);" class="active" onClick="SetRContent('MainLeft','Get.asp?Type=2');SetRContent('MainRight','Module.asp?Type=1&amp;MenuId=1');$('#QuickNav1').addClass('active'); $('#QuickNav2').removeClass('active');return false;" id="QuickNav1">功能设置</a><span></span>
		<a href="javascript:void(0);" onClick="SetRContent('MainLeft','Get.asp?Type=7&MenuId=1');SetRContent('MainRight','Module.asp?Type=1&amp;MenuId=1');$('#QuickNav2').addClass('active'); $('#QuickNav1').removeClass('active');return false;" id="QuickNav2">内容管理</a><span></span>
		<a href="/" target="_blank">网站预览</a>
	</div>
</div>

<div class="gnsz page">
	<div class="gnszmain">
		<dl class="zsptlist">
			<dt id="MainLeft">
			</dt>
			<dd id="MainRight" class="gnszallpage">
			</dd>
		</dl>
	</div>
</div>

<div id="Boxs" style="display:none">
        <div id="BoxsContent">
            <div id="BoxContent">
            </div>
        </div>
        <div id="AlphaBox" onClick="$('select').show();$('#Boxs').hide()"></div>
    </div>
</body>
</html>
<!--#Include File="../Code.asp"-->