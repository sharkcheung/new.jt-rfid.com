<!--#Include File="../AdminCheck.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link href="/admin/Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/Js/function.js"></script>
<%
'==========================================
'文 件 名：Weixin_Menu.asp
'文件用途：微信图文管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System2") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_menuName,Fk_menuType,Fk_menuEvent,Fk_menuStatus,Fk_menuPx,Fk_menuParent

'获取参数
Types=Clng(Request.QueryString("Type"))
id=Clng(Request.QueryString("id"))

Select Case Types
	Case 1
		Call WeixinImgTextList() '微信图文列表
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：WeixinImgTextList()
'作    用：微信图文列表
'参    数：
'==========================================
Sub WeixinImgTextList()
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
%>
<script type="text/javascript">
 function closeWin()
   {
    window.parent.ymPrompt.doHandler("error",true);
   }
</script>
<style type="text/css">
body{padding:4px;}
#ListTop,#ListContent{width:100%;margin:0 auto}
#ListContent table{border-right:0}
#ListContent table td{line-height:34px;}
p{line-height:190%;padding:20px;}
</style>
</head>
<body>
<div id="ListContent">
<p>文本内容<br><textarea name="val" id="val" cols="50" rows="6"></textarea></p>
</div>

<script type="text/javascript">
//parentWin.document.getElementById('wx_Subscribe');
var v=window.parent.document.getElementById('wx_Subscribe');
if (v.value.toLowerCase().indexOf("[wx_news")<0){
	document.getElementById("val").value=v.value;
}
</script>
</body>
</html>
<%
End Sub
%><!--#Include File="../../Code.asp"-->