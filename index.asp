<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：Index.asp
'文件用途：首页
'版权所有：企帮网络www.qebang.cn
'==========================================
'增加百度或谷歌推广来源判断



if request("bdclkid")<>"" or request("gclid")<>"" or request("sukey")<>"" or request("from")<>"" then
	response.redirect "/"
end if

if len(Site301)>0 then
	dim site301list,ubSite301,site301i,curDomain,mHost,toUrl,toUrllist
	curDomain=request.ServerVariables("HTTP_HOST")
	'site301="c.com|c.com,a.cn,a.com,www.a.cn"
	'curDomain="a.com"
	if instr(Site301,"|")>0 then	'至少要有两域名才能执行以下操作
		site301list=split(Site301,"|")    'a.com,b.com.www.c.com,www.a.com
		ubSite301=ubound(site301list)
		if instr(trim(site301list(0)),"www.")>0 then		'如果主域名带了www.
			mHost=trim(site301list(0))
		else
			mHost="www."&trim(site301list(0))
		end if 
		toUrllist=site301list(1)
		toUrl="http://"&mHost		'最终跳转链接
		if instr(","&trim(toUrllist)&",",","&trim(curDomain)&",")>0 then		'如果当前域名存在于跳转列表
			if trim(curDomain)=mHost then
			else
				Response.Status="301 Moved Permanently"
				Response.AddHeader "Location", toUrl
			end if
		else
'			if trim(curDomain)<>mHost then
'				if "www."&trim(curDomain)=mHost then
'					Response.Status="301 Moved Permanently"
'					Response.AddHeader "Location", toUrl
'				end if
'			end if
		end if
	end if
end if

'定义变量
Dim PageCode,PageCodes,PageUrl,PageType,PageFileName,CategoryDirName
Dim Fk_Article_Module,Fk_Product_Module,Fk_Down_Module
Dim FKTemplate,TemplateId
Dim Rst
Set Rst=Server.Createobject("Adodb.RecordSet")
Set FKTemplate=New Cls_Template

PageUrl=FKFun.HTMLEncode(Request.QueryString())

'判断首页
If PageUrl="" And SiteHtml=0 Then
	'抽取模板内容
	Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='index'"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		PageCode=Rs("Fk_Template_Content")
	Else
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write arrTips(0)
		Response.End()
	End If
	Rs.Close
	If SiteTest=1 Then
		PageCode=FKFso.FsoFileRead("Skin/"&SiteTemplate&"/index.html")
	End If
	'处理函数
	PageCode=FKTemplate.FileChange(PageCode)
	PageCode=FKTemplate.MoreUrlChange(PageCode)
	PageCode=Replace(PageCode,"{$SitePageNow$}","0")
	PageCode=FKTemplate.SiteChange(PageCode)
	PageCode=FKTemplate.TemplateDo(PageCode)
	PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(0))
ElseIf PageUrl="" And SiteHtml=1 Then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location", "/Index.html"
	'Response.Redirect("/Index.html")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

If SiteHtml=1 And PageUrl<>"" Then
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location", SiteDir&"html/"&PageUrl
	'Response.Redirect(SiteDir&"html/"&PageUrl)
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If
'判断类别
If Instr(PageUrl,"/")=0 And PageUrl<>"" Then
	PageFileName=Split(PageUrl,".")(0)
	If Left(PageFileName,4)="Info" And IsNumeric(Replace(PageFileName,"Info","")) Then
		Id=Clng(Replace(PageFileName,"Info",""))
		PageType=3
	ElseIf Left(PageFileName,4)="Page" And IsNumeric(Replace(PageFileName,"Page","")) Then
		Id=Clng(Replace(PageFileName,"Page",""))
		PageType=0
	ElseIf Left(PageFileName,5)="GBook" And (IsNumeric(Replace(PageFileName,"GBook","")) Or IsNumeric(Split(Replace(PageFileName,"GBook",""),"__")(0))) Then
		If Instr(Replace(PageFileName,"GBook",""),"__")>0 Then
			Id=Clng(Split(Replace(PageFileName,"GBook",""),"__")(0))
			PageNow=Clng(Split(Replace(PageFileName,"GBook",""),"__")(1))
		Else
			Id=Clng(Replace(PageFileName,"GBook",""))
			PageNow=1
		End If
		PageType=4
	ElseIf Left(PageFileName,3)="Job" And IsNumeric(Replace(PageFileName,"Job","")) Then
		Id=Clng(Replace(PageFileName,"Job",""))
		PageType=6
	Else
		If Instr(PageFileName,"__")>0 Then
			PageNow=Clng(Split(PageFileName,"__")(1))
			PageFileName=Split(PageFileName,"__")(0)
		Else
			PageNow=1
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_FileName='"&PageFileName&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			TemplateId=Rs("Fk_Module_Template")
			PageType=Rs("Fk_Module_Type")
			dim tiaozhuanurl:tiaozhuanurl=Rs("Fk_Module_Url")	'2012-2-24 增加跳转链接栏目类型的页面跳转
			If Not IsNull(Rs("Fk_Module_PageCode")) Then
				PageArr=Split(Rs("Fk_Module_PageCode"),"|--|")
			End If
			If PageType=4 Then
				PageArr=Split(Rs("Fk_Module_PageCode"),"|--|")
				If Rs("Fk_Module_PageCount")>0 Then
					PageSizes=Rs("Fk_Module_PageCount")
				End If
				CategoryDirName=PageFileName
			End If
			Id=Rs("Fk_Module_Id")
		Else
			Rs.Close
			Call FKDB.DB_Close()
			Response.redirect("error.html")
			Session.CodePage=936
			Response.End()
		End If
		Rs.Close
		'2012-2-24 增加跳转链接栏目类型的页面跳转
		If PageType=5 and tiaozhuanurl<>"" Then
			response.Redirect tiaozhuanurl
		End If
	End If
	If TemplateId="" Then
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Id
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			TemplateId=Rs("Fk_Module_Template")
			If PageType=4 Then
				PageArr=Split(Rs("Fk_Module_PageCode"),"|--|")
				If Rs("Fk_Module_PageCount")>0 Then
					PageSizes=Rs("Fk_Module_PageCount")
				End If
				CategoryDirName="GBook"&Id
			End If
		Else
			Rs.Close
			Call FKDB.DB_Close()
			Response.redirect("error.html")
			Session.CodePage=936
			Response.End()
		End If
		Rs.Close
	End If
	Select Case PageType
		Case 0
			If TemplateId=0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='page'"
			Else
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
			End If
		Case 3
			If TemplateId=0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='info'"
			Else
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
			End If
		Case 4
			If TemplateId=0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='gbook'"
			Else
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
			End If
		Case 6
			If TemplateId=0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='job'"
			Else
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
			End If
	End Select
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		PageCode=Rs("Fk_Template_Content")
		Templates=Rs("Fk_Template_Name")
	Else
		Rs.Close
		Call FKDB.DB_Close()
		Response.Write arrTips(1)
		Session.CodePage=936
		Response.End()
	End If
	Rs.Close
	If SiteTest=1 Then
		PageCode=FKFso.FsoFileRead("Skin/"&SiteTemplate&"/"&Templates&".html")
	End If
	PageCode=FKTemplate.FileChange(PageCode)
	Select Case PageType
		Case 0
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.PageChange(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
		Case 3
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.InfoChange(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
		Case 4
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.GBookChange(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
		Case 6
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.JobChange(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
	End Select
	PageCode=FKTemplate.TemplateDo(PageCode)
	If PageType=4 Then
		PageCode=Replace(PageCode,"{$GBookPage$}",FKTemplate.ShowPageCode(SiteDir&sTemp&"?"&CategoryDirName&"__{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
		PageCode=FKTemplate.PageCodeChange(PageCode)
	End If
	PageCode=FKTemplate.MoreUrlChange(PageCode)
ElseIf PageUrl<>"" Then
	CategoryDirName=Split(PageUrl,"/")(0)
	PageFileName=Split(Split(PageUrl,"/")(1),".")(0)
	If Left(CategoryDirName,7)="Article" And IsNumeric(Replace(CategoryDirName,"Article","")) Then
		Id=Clng(Replace(CategoryDirName,"Article",""))
		PageType=1
	ElseIf Left(CategoryDirName,7)="Product" And IsNumeric(Replace(CategoryDirName,"Product","")) Then
		Id=Clng(Replace(CategoryDirName,"Product",""))
		PageType=2
	ElseIf Left(CategoryDirName,4)="Down" And IsNumeric(Replace(CategoryDirName,"Down","")) Then
		Id=Clng(Replace(CategoryDirName,"Down",""))
		PageType=7
	Else
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Dir='"&CategoryDirName&"'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageType=Rs("Fk_Module_Type")
			PageArr=Split(Rs("Fk_Module_PageCode"),"|--|")
			Id=Rs("Fk_Module_Id")
		Else
			Rs.Close
			Call FKDB.DB_Close()
			Response.redirect("error.html")
			Session.CodePage=936
			Response.End()
		End If
		Rs.Close
	End If
	If Left(PageFileName,5)="Index" Then
		If Instr(PageFileName,"_") Then
			PageNow=Clng(Split(PageFileName,"_")(1))
		Else
			PageNow=1
		End If
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Id
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			TemplateId=Rs("Fk_Module_Template")
			PageArr=Split(Rs("Fk_Module_PageCode"),"|--|")
			If Rs("Fk_Module_PageCount")>0 Then
				PageSizes=Rs("Fk_Module_PageCount")
			End If
		Else
			Rs.Close
			Call FKDB.DB_Close()
			Response.redirect("error.html")
			Session.CodePage=936
			Response.End()
		End If
		Rs.Close
		Select Case PageType
			Case 1
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='articlelist'"
				Else
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
				End If
			Case 2
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='productlist'"
				Else
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
				End If
			Case 7
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='downlist'"
				Else
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
				End If
		End Select
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Templates=Rs("Fk_Template_Name")
		Else
			Rs.Close
			Call FKDB.DB_Close()
			Response.Write arrTips(1)
			Session.CodePage=936
			Response.End()
		End If
		Rs.Close
		If SiteTest=1 Then
			PageCode=FKFso.FsoFileRead("Skin/"&SiteTemplate&"/"&Templates&".html")
		End If
		PageCode=FKTemplate.FileChange(PageCode)
		Select Case PageType
			Case 1
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.ArticleListChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$ArticleCategoryPage$}",FKTemplate.ShowPageCode(SiteDir&sTemp&"?"&CategoryDirName&"/Index_{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				PageCode=FKTemplate.PageCodeChange(PageCode)
			Case 2
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.ProductListChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$ProductCategoryPage$}",FKTemplate.ShowPageCode(SiteDir&sTemp&"?"&CategoryDirName&"/Index_{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				PageCode=FKTemplate.PageCodeChange(PageCode)
			Case 7
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.DownListChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$DownCategoryPage$}",FKTemplate.ShowPageCode(SiteDir&sTemp&"?"&CategoryDirName&"/Index_{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				PageCode=FKTemplate.PageCodeChange(PageCode)
		End Select
	Else
		If IsNumeric(PageFileName) Then
			Select Case PageType
				Case 1
					Sqlstr="Select * From [Fk_Article] Where Fk_Article_Id=" & PageFileName
				Case 2
					Sqlstr="Select * From [Fk_Product] Where Fk_Product_Id=" & PageFileName
				Case 7
					Sqlstr="Select * From [Fk_Down] Where Fk_Down_Id=" & PageFileName
			End Select
		Else
			Select Case PageType
				Case 1
					Sqlstr="Select * From [Fk_Article] Where Fk_Article_FileName='"&PageFileName&"'"
				Case 2
					Sqlstr="Select * From [Fk_Product] Where Fk_Product_FileName='"&PageFileName&"'"
				Case 7
					Sqlstr="Select * From [Fk_Down] Where Fk_Down_FileName='"&PageFileName&"'"
			End Select
		End If
		Select Case PageType
			Case 1
				Rs.Open Sqlstr,Conn,1,3
				If Not Rs.Eof Then
					TemplateId=Rs("Fk_Article_Template")
					Id=Rs("Fk_Article_Id")
					Fk_Article_Module=Rs("Fk_Article_Module")
				Else
					Rs.Close
					Call FKDB.DB_Close()
					Response.redirect("error.html")
					Session.CodePage=936
					Response.End()
				End If
				Rs.Close
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Article_Module
					Rs.Open Sqlstr,Conn,1,3
					If Not Rs.Eof Then
						If Rs("Fk_Module_LowTemplate")>0 Then
							TemplateId=Rs("Fk_Module_LowTemplate")
						End If
					End If
					Rs.Close
				End If
			Case 2
				Rs.Open Sqlstr,Conn,1,3
				If Not Rs.Eof Then
					TemplateId=Rs("Fk_Product_Template")
					Id=Rs("Fk_Product_Id")
					Fk_Product_Module=Rs("Fk_Product_Module")
				Else
					Rs.Close
					Call FKDB.DB_Close()
					Response.redirect("error.html")
					Session.CodePage=936
					Response.End()
				End If
				Rs.Close
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Product_Module
					Rs.Open Sqlstr,Conn,1,3
					If Not Rs.Eof Then
						If Rs("Fk_Module_LowTemplate")>0 Then
							TemplateId=Rs("Fk_Module_LowTemplate")
						End If
					End If
					Rs.Close
				End If
			Case 7
				Rs.Open Sqlstr,Conn,1,3
				If Not Rs.Eof Then
					TemplateId=Rs("Fk_Down_Template")
					Id=Rs("Fk_Down_Id")
					Fk_Down_Module=Rs("Fk_Down_Module")
				Else
					Rs.Close
					Call FKDB.DB_Close()
					Response.redirect("error.html")
					Session.CodePage=936
					Response.End()
				End If
				Rs.Close
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Down_Module
					Rs.Open Sqlstr,Conn,1,3
					If Not Rs.Eof Then
						If Rs("Fk_Module_LowTemplate")>0 Then
							TemplateId=Rs("Fk_Module_LowTemplate")
						End If
					End If
					Rs.Close
				End If
		End Select
		Select Case PageType
			Case 1
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='article'"
				Else
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
				End If
			Case 2
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='product'"
				Else
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
				End If
			Case 7
				If TemplateId=0 Then
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='down'"
				Else
					Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
				End If
		End Select
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Templates=Rs("Fk_Template_Name")
		Else
			Rs.Close
			Call FKDB.DB_Close()
			Response.Write arrTips(1)
			Session.CodePage=936
			Response.End()
		End If
		Rs.Close
		If SiteTest=1 Then
			PageCode=FKFso.FsoFileRead("Skin/"&SiteTemplate&"/"&Templates&".html")
		End If
		PageCode=FKTemplate.FileChange(PageCode)
		Select Case PageType
			Case 1
				PageCode=Replace(PageCode,"{$SitePageNow$}",Fk_Article_Module)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.ArticleChange(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Fk_Article_Module))
			Case 2
				PageCode=Replace(PageCode,"{$SitePageNow$}",Fk_Product_Module)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.ProductChange(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Fk_Product_Module))
			Case 7
				PageCode=Replace(PageCode,"{$SitePageNow$}",Fk_Down_Module)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.DownChange(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Fk_Down_Module))
		End Select
		PageCode=FKTemplate.TemplateDo(PageCode)
	End If
	PageCode=FKTemplate.MoreUrlChange(PageCode)
End If
	response.write(PageCode)
%>
<!--#Include File="Code.asp"-->
