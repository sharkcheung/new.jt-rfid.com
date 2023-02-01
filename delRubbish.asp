<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>   
/*总容器样式*/  
html, body, tr, td, a, span, p, br, hr, h1, h2, h3, h4, h5, h6, div, img, ol, ul, li, dl, dt, dd, iframe, sub, sup, blockquto {
	margin: 0;
	padding: 0;
	font-size: 100%;
	outline: none;
}
img {
	border: none;
}
ol, ul {
	list-style: none;
}
a {
	text-decoration: none;
	color:#fff;
}
a:hover {
	color: #F60;
	text-decoration: underline;
}
body {
	font: 12px/1.5 Arial,Tahoma,\5b8b\4f53;
	color:#fff;
	/*background:url(images/body_bg.jpg);*/
}
.clear {
	clear: both;
	height: 0;
	width: 0;
	margin: 0;
	padding: 0;
	overflow: hidden;
	float: none;
	visibility: hidden;
}
#content{width:970px;height:536px;margin:0 auto;}
.m{color:#1B5E76;}
.m p{width:600px;border:#ccc solid 1px;margin:0 auto;margin-top:150px;padding:20px;line-height:190%;}
.m p .runup {background:url(images/btnrun.jpg);width:80px;height:31px;border:none;cursor:pointer;}

#box{margin:0 auto;width:970px;background: url(images/bg.jpg); height:536px;}
.main{padding-left:160px;}

.prettyprint{cursor:pointer;margin:5px 10px 10px 10px !important;padding:2px !important;word-wrap:break-word;}
h2,dl,dd{padding:2px;margin:0px;color:#0000FF;}
.pager { padding: 3px; text-align: center; color:#66C;font-size:12px; font-family:Tahoma;}   
/*分页链接样式*/  
.pager a { margin: 2px; padding:2px 5px; color: #66C; text-decoration: none; border: 1px solid #aad; }   
/*分页链接鼠标移过的样式*/  
.pager a:hover { color: #000; border: 1px solid #009; background-color:#DCDCF3; }   
/*当前页码的样式*/  
.pager span.current { font-weight: bold; margin: 0 2px; padding: 2px 5px; color: #fff; background-color: #66C; border: 1px solid #009; }   
/*不可用分页链接的样式(比如第1页时的“上一页”链接)*/  
.pager span.disabled { margin: 0 2px; padding: 2px 5px; color: #CCC; border: 1px solid #DDD; }   
/*跳转下拉菜单的样式*/  
.pager select {margin: 0px 2px -2px 2px; color:#66C;font-size:12px; font-family:Tahoma;}   
/*跳转文本框的样式*/  
.pager input {margin: 0px 2px -2px 2px; color:#66C; border: 1px solid #DDD; padding:2px; text-align:center;font-size:12px; font-family:Tahoma;}   
h2 a{margin: 2px; padding:2px 5px; color: #f00; text-decoration: none; }
</style>
<meta content="IE=7" http-equiv="X-UA-Compatible" />
<title>商赢快车垃圾清除助手V1.0 BY SharkCheung 2011/10/10</title>
</head>
<body>
<!--#include file="easp.asp"-->
<div id="content">
<h2><a href="?T=0">新闻垃圾</a> | <a href="?T=1">产品垃圾</a> <a href="?T=0&a=del">删除新闻垃圾</a> | <a href="?T=1&a=del">删除产品垃圾</a></h2>
<dl>
			<%
			Z.db.DbConn = Z.db.OpenConn(1,"Data%2f69fa10c941e0134b\data.mdb","")
		if Z.RQ("a",0)="del" THEN
			IF Z.RQ("T",1)=0 THEN
				delTable="Fk_Article"
				delName="新闻垃圾"
				delSQL="Fk_Article_Id not in(select a.Fk_Article_Id from Fk_Article a inner join Fk_Module b on a.Fk_Article_Module=b.Fk_Module_Id)"
			elseif Z.RQ("T",1)=1 THEN
				delTable="Fk_Product"
				delName="产品垃圾"
				delSQL="Fk_Product_Id not in(select a.Fk_Product_Id from Fk_Product a inner join Fk_Module b on a.Fk_Product_Module=b.Fk_Module_Id)"
			end if
			result = Z.db.DeleteRecord(delTable, delSQL)
			if result=1 then
				Z.alertUrl delName&"删除完毕！", "delRubbish.asp?T="&Z.R("T",1)
			else
				Z.alertUrl delName&"删除失败！", "delRubbish.asp?T="&Z.R("T",1)
			end if
		else
			if Z.RQ("T",1)=0 THEN
				sql="select Fk_Article_Id,Fk_Article_Title from Fk_Article where Fk_Article_Id not in(select a.Fk_Article_Id from Fk_Article a inner join Fk_Module b on a.Fk_Article_Module=b.Fk_Module_Id order by a.Fk_Article_Module asc)"
			else
				sql="select Fk_Product_Id ,Fk_Product_Title from Fk_Product where Fk_Product_Id not in(select Fk_Product_Id from Fk_Product a inner join Fk_Module b on a.Fk_Product_Module=b.Fk_Module_Id order by a.Fk_Product_Module asc)"
			end if
			Dim rs : Set rs = Z.db.GetPageRecord(1,sql)   
			'Set rs = Easp.db.GR("Fk_Article:Fk_Article_Id,Fk_Article_Title,Fk_Article_Module", "", "Fk_Article_Module Asc")
			Z.w "<dd>I D | 标 题 </dd>"
			i = 0   
			While Not rs.Eof And ( i < rs.PageSize ) 
				i=i+1
    			Z.w "<dd>"&rs(0) & " | " & rs(1) & " </dd>"  
    			rs.MoveNext()   
			wend 
			Z.db.SetPager "default", "<div class=""pager"">{first}{prev}{liststart}{list}{listend}{next}{last}  共{recordcount}条, 每页{pagesize}条, {pageindex}/{pagecount}页, 转到{jump}页</div>", Array("jump:select","jumplong:20")
			Z.WC Z.db.GetPager("")  '可以用""空参数来调用"default"样式，自定义了名称就用名称   
	    	Z.C(rs)
		end if
'		Z.db.PageSize=30
'		Set rs = Z.db.GetPageRecord(0,Array("Fk_Module:Fk_Module_Id,Fk_Module_Name", "", "Fk_Module_Id Asc"))   
'		'Set rs = Easp.db.GR("Fk_Article:Fk_Article_Id,Fk_Article_Title,Fk_Article_Module", "", "Fk_Article_Module Asc")
'		Z.db.SetPager "default", "<div class=""pager"">{first}{prev}{liststart}{list}{listend}{next}{last}  共{recordcount}条, 每页{pagesize}条, {pageindex}/{pagecount}页, 转到{jump}页</div>", Array("jump:select","jumplong:30")
'		Z.w "<dd>模 块 I D | 模 块 标 题 | </dd>"
'		i = 0   
'		While Not rs.Eof And ( i < Z.db.PageSize ) 
'			i=i+1
'    		Z.w "<dd>"&rs("Fk_Module_Id") & " | " & rs("Fk_Module_Name") & " </dd>"  
'    		rs.MoveNext()   
'		wend 
'		Z.WC Z.db.GetPager("")  '可以用""空参数来调用"default"样式，自定义了名称就用名称   
'	    Z.C(rs)
		%>
		</dl>
</div>
</body>
</html>
