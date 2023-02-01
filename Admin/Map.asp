<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_Template.asp"-->
<%
'==========================================
'文 件 名：Map.asp
'文件用途：Google地图生成拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	'Response.Write("无权限！")
	'Call FKDB.DB_Close()
	'Session.CodePage=936
	'Response.End()
End If

Dim FKTemplate
Set FKTemplate=New Cls_Template

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call MapBox() '读取SEO索引地图生成器
	Case 2
		Call MapDo() '生成地图
End Select

'==========================================
'函 数 名：MapBox()
'作    用：读取SEO索引地图生成器
'参    数：
'==========================================
Sub MapBox()
%>
<div id="BoxTop" style="width:98%;"><span>SEO索引地图</span></div>
<div id="BoxContents" style="width:98%;">
<table width="95%" border="0" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="34%" height="25" align="center"><a style="display:none;" href="javascript:void(0);" onclick="document.getElementById('Gets').src='Map.asp?Type=2';">生成SEO索引地图</a>
<input type="button" onclick="document.getElementById('Gets').src='Map.asp?Type=2';" class="Button2" name="button" id="button" value="一键生成SEO索引地图" />
</td>
        <td width="24%" align="center">&nbsp;</td>
        <td width="17%" align="center">&nbsp;</td>
        <td width="25%" align="center">&nbsp;</td>
        </tr>
    <tr>
        <td height="25" colspan="4" align="center">&nbsp;&nbsp;生成结果显示</td>
    </tr>
    <tr>
        <td height="25" colspan="4" id="Template" style="padding:10px; line-height:22px; font-size:14px;"><iframe src="Map.asp?Type=2&view=1" id="Gets" border="0" frameborder="0" width="98%" height="250px"></iframe></td>
        </tr>
</table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub

'==============================
'函 数 名：MapDo()
'作    用：生成地图
'参    数：
'==============================
Sub MapDo() 
%>
<html oncontextmenu="return false">
<STYLE> 
* {
	margin:0;
	padding:0;
	font-family:tahoma, verdana, 宋体;
}
body {
	font-size:11.5px;
	SCROLLBAR-FACE-COLOR: #e8e7e7; 
	SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; 
	SCROLLBAR-SHADOW-COLOR: #ffffff; 
	SCROLLBAR-3DLIGHT-COLOR: #cccccc; 
	SCROLLBAR-ARROW-COLOR: #03B7EC; 
	SCROLLBAR-TRACK-COLOR: #EFEFEF; 
	SCROLLBAR-DARKSHADOW-COLOR: #b2b2b2; 
	SCROLLBAR-BASE-COLOR: #000000;
	margin:10px;
	line-height:20px;
}
a {
	font-size: 11.5px;
	color: #000;
	text-decoration: none;
}
a:visited {
	color: #000;
	text-decoration: none;
}
a:hover {
	color: #000;
	text-decoration: none;
}
a:active {
	color: #000;
	text-decoration: none;
}
url{
	clear:both; width:99%;
}
</STYLE>
<%
if request("view")<>"1" then
	Dim ArticleUrl,ProductUrl,DownUrl
	Temp="<?xml version=""1.0"" encoding=""UTF-8""?>"&vbLf
	Temp=Temp&"<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">"&vbLf
	Temp=Temp&"<url>"&vbLf
	Temp=Temp&"<loc>"&SiteUrl&"/</loc>"&vbLf
	Temp=Temp&"</url>"&vbLf
	Sqlstr="Select * From [Fk_ArticleList] Order By Fk_Article_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
		If Rs("Fk_Module_Dir")<>"" Then
			ArticleUrl=Rs("Fk_Module_Dir")&"/"
		Else
			ArticleUrl="Article"&Rs("Fk_Module_Id")&"/"
		End If
		If Rs("Fk_Article_FileName")<>"" Then
			ArticleUrl=ArticleUrl&Rs("Fk_Article_FileName")&".html"
		Else
			ArticleUrl=ArticleUrl&Rs("Fk_Article_Id")&".html"
		End If
		If SiteHtml=1 Then
			ArticleUrl=ArticleUrl
		Else
			ArticleUrl=sTemp&"?"&ArticleUrl
		End If
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&SiteUrl&"/"&ArticleUrl&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select * From [Fk_ProductList] Order By Fk_Product_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
		If Rs("Fk_Module_Dir")<>"" Then
			ProductUrl=Rs("Fk_Module_Dir")&"/"
		Else
			ProductUrl="Product"&Rs("Fk_Module_Id")&"/"
		End If
		If Rs("Fk_Product_FileName")<>"" Then
			ProductUrl=ProductUrl&Rs("Fk_Product_FileName")&".html"
		Else
			ProductUrl=ProductUrl&Rs("Fk_Product_Id")&".html"
		End If
		If SiteHtml=1 Then
			ProductUrl=ProductUrl
		Else
			ProductUrl=sTemp&"?"&ProductUrl
		End If
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&SiteUrl&"/"&ProductUrl&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select * From [Fk_DownList] Order By Fk_Down_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
		If Rs("Fk_Module_Dir")<>"" Then
			DownUrl=Rs("Fk_Module_Dir")&"/"
		Else
			DownUrl="Down"&Rs("Fk_Module_Id")&"/"
		End If
		If Rs("Fk_Down_FileName")<>"" Then
			DownUrl=DownUrl&Rs("Fk_Down_FileName")&".html"
		Else
			DownUrl=DownUrl&Rs("Fk_Down_Id")&".html"
		End If
		If SiteHtml=1 Then
			DownUrl=DownUrl
		Else
			DownUrl=sTemp&"?"&DownUrl
		End If
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&SiteUrl&"/"&DownUrl&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type<>5 Order By Fk_Module_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
		DownUrl=FKTemplate.GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName"))
		DownUrl=Right(DownUrl,(Len(DownUrl)-Len(SiteDir)))
		Temp=Temp&"<url>"&vbLf
		Temp=Temp&"<loc>"&SiteUrl&"/"&DownUrl&"</loc>"&vbLf
		Temp=Temp&"</url>"&vbLf
		Rs.MoveNext
	Wend
	Rs.Close
	Temp=Temp&"</urlset>"&vbLf
	Call FKFso.CreateFile("../sitemap.xml",Temp)
	Response.Write("<p><a href=""/sitemap.xml"" target=""_blank"">SEO索引地图生成成功!</a></p>")
	temp=replace(temp,"</url>","</url><br>")
	Response.write Temp&"</html>"
else
	Response.write "</html>"
end if
End Sub
%>
<!--#Include File="../Code.asp"-->