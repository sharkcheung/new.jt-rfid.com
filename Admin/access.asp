<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Index.asp
'文件用途：后台管理首页
'版权所有：深圳企帮
dim Filename,Viewstyle
'==========================================
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  oncontextmenu="return false">
<head>
<meta content="IE=7" http-equiv="X-UA-Compatible" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../Js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="../Js/jquery.form.min.js"></script>
<script type="text/javascript" src="../Js/function.js"></script>
<%if Bianjiqi="kinediter" then %>
<script type="text/javascript" charset="utf-8" src="/editor/kindeditor.js"></script>
<%else%>
<script type="text/javascript" src="../Js/xheditor-zh-cn.min.js"></script>
<%end if%>
<!-- ymPrompt组件 -->
<script type="text/javascript" src="winskin/ymPrompt.js"></script>
<link rel="stylesheet" type="text/css" href="winskin/qq/ymPrompt.css" /> 
<!-- ymPrompt组件 -->
</head>

<body>
<%
		'Call FKDB.DB_Open()
		'on error resume next
	'	rs.open "select * from keywordSV",conn,1,1
	'	if err.number<1 then
	'	Sqlstr="create table keywordSV(id COUNTER CONSTRAINT PrimaryKey PRIMARY KEY,SVkeywords text(255),SVci int,SVpaiming text(255),SVb1 text(255),SVb2 text(255),SVb3 text(255))"
	'	Conn.Execute(Sqlstr)
	' 	response.write "创建成功"
	'	else
	'	response.write "表已经存在"
	'	end if
	'	rs.close
%>
<div id="chaciarea1" onmousedown="innering(1);" onclick="chakeywordspaiming('腾讯',1);" style="cursor:pointer;">asdfasdf</div>
<div id="chaciarea2" onmousedown="innering(2);" onclick="chakeywordspaiming('奥巴马',2);" style="cursor:pointer;">asdfasdf</div>
<div id="chaciarea3" onmousedown="innering(3);" onclick="chakeywordspaiming('中国',3);" style="cursor:pointer;">asdfasdf</div>

</body>
</html>
<!--#Include File="../Code.asp"-->