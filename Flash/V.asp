<!--#Include File="../Include.asp"--><%
'==========================================
'文 件 名：V.asp
'文件用途：Flash轮换
'版权所有：企帮网络www.qebang.cn
'==========================================

Dim Height,Width,Menu,Module
Dim Pic,Text,Url,ArticleUrl,ProductUrl,DownUrl

'获取参数
Types=Clng(Request.QueryString("Type"))
Menu=Clng(Request.QueryString("Menu"))
Module=Clng(Request.QueryString("Module"))
Height=Request.QueryString("Height")
Width=Request.QueryString("Width")
If Height="" Then
	Height=100
Else
	Height=Clng(Height)
End If
If Width="" Then
	Width=100
Else
	Width=Clng(Width)
End If

If Types=1 Then
	If Module=0 Then
		Sqlstr="Select Top 5 * From [Fk_ArticleList] Where Fk_Article_Menu="&Menu&" And Fk_Article_Pic<>'' Order By Fk_Article_Ip desc, Px desc, Fk_Article_Id Desc"
		Temp=1
	Else
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id="&Module&""
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			If Rs("Fk_Module_Type")=1 Or Rs("Fk_Module_Type")=2 Or Rs("Fk_Module_Type")=7 Then
				Temp=Rs("Fk_Module_Type")
			Else
				Rs.Close
				Response.End()
			End If
		Else
			Rs.Close
			Response.End()
		End If
		Rs.Close
		If Temp=1 Then
			Sqlstr="Select Top 5 * From [Fk_ArticleList] Where (Fk_Article_Module="&Module&" Or Fk_Module_LevelList Like '%%,"&Module&",%%') And Fk_Article_Pic<>'' Order By Fk_Article_Ip desc, Px desc, Fk_Article_Id Desc"
		ElseIf Temp=2 Then
			Sqlstr="Select Top 5 * From [Fk_ProductList] Where (Fk_Product_Module="&Module&" Or Fk_Module_LevelList Like '%%,"&Module&",%%') And Fk_Product_Pic<>'' Order By Fk_Product_Ip desc, Px desc, Fk_Product_Id Desc"
		ElseIf Temp=7 Then
			Sqlstr="Select Top 5 * From [Fk_DownList] Where (Fk_Down_Module="&Module&" Or Fk_Down_LevelList Like '%%,"&Module&",%%') And Fk_Down_Pic<>'' Order By Fk_Down_Ip desc, Px desc, Fk_Down_Id Desc"
		End If
	End If
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
		If Temp=1 Then
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
				ArticleUrl=SiteDir&ArticleUrl
			Else
				ArticleUrl=SiteDir&"?"&ArticleUrl
			End If
			If Pic="" Then
				Pic=Rs("Fk_Article_Pic")
			Else
				Pic=Pic&"|"&Rs("Fk_Article_Pic")
			End If
			If Text="" Then
				Text=Rs("Fk_Article_Title")
			Else
				Text=Text&"|"&Rs("Fk_Article_Title")
			End If
			If Url="" Then
				Url=ArticleUrl
			Else
				Url=Url&"|"&ArticleUrl
			End If
		ElseIf Temp=2 Then
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
				ProductUrl=SiteDir&ProductUrl
			Else
				ProductUrl=SiteDir&"?"&ProductUrl
			End If
			If Pic="" Then
				Pic=Rs("Fk_Product_Pic")
			Else
				Pic=Pic&"|"&Rs("Fk_Product_Pic")
			End If
			If Text="" Then
				Text=Rs("Fk_Product_Title")
			Else
				Text=Text&"|"&Rs("Fk_Product_Title")
			End If
			If Url="" Then
				Url=ProductUrl
			Else
				Url=Url&"|"&ProductUrl
			End If
		ElseIf Temp=7 Then
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
				DownUrl=SiteDir&DownUrl
			Else
				DownUrl=SiteDir&"?"&DownUrl
			End If
			If Pic="" Then
				Pic=Rs("Fk_Down_Pic")
			Else
				Pic=Pic&"|"&Rs("Fk_Down_Pic")
			End If
			If Text="" Then
				Text=Rs("Fk_Down_Title")
			Else
				Text=Text&"|"&Rs("Fk_Down_Title")
			End If
			If Url="" Then
				Url=DownUrl
			Else
				Url=Url&"|"&DownUrl
			End If
		End If
		Rs.MoveNext
	Wend
	Rs.Close
%>
var focus_width=<%=Width%>     //场景宽
var focus_height=<%=Height%>　　//场景高
var text_height=0　　　//文字说明字高，为0时不显示文本
var swf_height = focus_height+text_height
var pics='<%=Pic%>'
var links='<%=Url%>'
var texts='<%=Text%>'
document.write('<object ID="focus_flash" classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=8,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="<%=SiteDir%>Flash/1/focus.swf"><param name="quality" value="high"><param name="bgcolor" value="#ffffff">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
document.write('<embed ID="focus_flash" src="<%=SiteDir%>Flash/1/focus.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" class=tablebody1 quality="high" width="'+ focus_width +'" height="'+ focus_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer"/>');  document.write('</object>');
<%
End If
%>
<!--#Include File="../Code.asp"-->
