<!--#Include File="Cls_Template.asp"-->
<%
'==========================================
'文 件 名：Cls_HTML.asp
'文件用途：模板引擎函数类
'==========================================

Dim FKTemplate
Set FKTemplate=New Cls_Template
Dim Rst
Set Rst=Server.Createobject("Adodb.RecordSet")

Class Cls_HTML
	Private PageCode,PageCodes,Pages
	
	'==============================
	'函 数 名：CreatIndex
	'作    用：生成首页
	'参    数：
	'==============================
	Public Function CreatIndex()
		Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='index'"
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}","0")
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(0))
			PageCode=FKTemplate.MoreUrlChange(PageCode)
			Call FKFso.CreateFile("/Index.html",PageCode)
			Response.Write("<p><a href=""/Index.html"" target=""_blank"">首页生成成功</a></p>")
			Response.Flush()
			Response.Clear()
		Else
			Response.Write("首页模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatInfo
	'作    用：生成信息页
	'参    数：
	'==============================
	Public Function CreatInfo(TemplateId,ModuleFileName,ModuleName,CType)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='info'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.InfoChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
			Temp=FKTemplate.GetGoUrl(3,Id,"",ModuleFileName)
			if(instr(temp,"/html")=1) then
				Temp=".."&Temp
			end if
			Call FKFso.CreateFile(Temp,PageCode)
			If CType=0 Then
				Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			End If
		Else
			Response.Write(ModuleName&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatPage
	'作    用：生成静态页
	'参    数：
	'==============================
	Public Function CreatPage(TemplateId,ModuleFileName,ModuleName)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='page'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.PageChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
			Temp=FKTemplate.GetGoUrl(0,Id,"",ModuleFileName)
			if(instr(Temp,"/html")=1) then
				Temp=".."&Temp
			end if
			Call FKFso.CreateFile(Temp,PageCode)
			Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”生成成功</a></p>")
			Response.Flush()
			Response.Clear()
		Else
			Response.Write(ModuleName&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatGBook
	'作    用：生成留言页
	'参    数：
	'==============================
	Public Function CreatGBook(TemplateId,ModuleFileName,ModuleName)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='gbook'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			If Instr(PageCode,"{$GBookPage$}")=0 Then
				PageCode=FKTemplate.FileChange(PageCode)
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.GBookChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				Temp=FKTemplate.GetGoUrl(4,Id,"",ModuleFileName)
				if(instr(Temp,"/html")=1) then
					Temp=".."&Temp
				end if
				Call FKFso.CreateFile(Temp,PageCode)
				Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			Else
				PageCodes=PageCode
				Sqlstr="Select * From [Fk_GBook] Where Fk_GBook_Module=" & Id
				Rs.Open Sqlstr,Conn,1,3
				If Not Rs.Eof Then
					Rs.PageSize=PageSizes
					Pages=Rs.PageCount
				Else
					Pages=1
				End If
				Rs.Close
				For j=1 To Pages
					PageCode=PageCodes
					PageNow=j
					If ModuleFileName="" Then
						Temp="GBook"&Id
					Else
						Temp=ModuleFileName
					End If
					PageCode=FKTemplate.FileChange(PageCode)
					PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
					PageCode=FKTemplate.SiteChange(PageCode)
					PageCode=FKTemplate.GBookChange(PageCode)
					PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
					PageCode=FKTemplate.TemplateDo(PageCode)
					If ModuleFileName="" Then
						Temp="GBook"&Id
					Else
						Temp=ModuleFileName
					End If
					Temp="../html/"&Temp
					PageCode=Replace(PageCode,"{$GBookPage$}",FKTemplate.ShowPageCode(Temp&"__{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
					PageCode=FKTemplate.PageCodeChange(PageCode)
					If j=1 Then
						Temp=Temp&".html"
					Else
						Temp=Temp&"__"&j&".html"
					End If
					Call FKFso.CreateFile(Temp,PageCode)
					Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”第"&j&"页生成成功</a></p>")
					Response.Flush()
					Response.Clear()
				Next
			End If
		Else
			Response.Write(ModuleName&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatJob
	'作    用：生成招聘页
	'参    数：
	'==============================
	Public Function CreatJob(TemplateId,ModuleFileName,ModuleName)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='job'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.JobChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
			Temp=FKTemplate.GetGoUrl(6,Id,"",ModuleFileName)
			Temp="../html/"&Temp
			Call FKFso.CreateFile(Temp,PageCode)
			Response.Write("<p><a href="""&Temp&""" target=""_blank"">"&ModuleName&"生成成功</a></p>")
			Response.Flush()
			Response.Clear()
		Else
			Response.Write("招聘模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatSubject
	'作    用：生成专题页
	'参    数：
	'==============================
	Public Function CreatSubject(TemplateId)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='subject'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.SubjectChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			Call FKFso.CreateFile("../html/Subject"&Id&".html",PageCode)
			Response.Write("<p><a href=""../html/Subject"&Id&".html"" target=""_blank"">“"&Fk_Subject_Name&"”生成成功</a></p>")
			Response.Flush()
			Response.Clear()
		Else
			Response.Write("专题页模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatArticle
	'作    用：生成文章页
	'参    数：
	'==============================
	Public Function CreatArticle(TemplateId,ModuleId,ModuleDir,ArticleFileName,ArticleTitle,CType)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & ModuleId
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				If Rs("Fk_Module_LowTemplate")>0 Then
					TemplateId=Rs("Fk_Module_LowTemplate")
				End If
			End If
			Rs.Close
		End If
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='article'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}",ModuleId)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.ArticleChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(ModuleId))
			If ModuleDir<>"" Then
				Temp=""&ModuleDir&"/"
			Else
				Temp="Article"&ModuleId&"/"
			End If
			If ArticleFileName<>"" Then
				Temp=Temp&ArticleFileName&".html"
			Else
				Temp=Temp&Id&".html"
			End If
			Call FKFso.CreateFile("../html/"&Temp,PageCode)
			If CType=0 Then
				Response.Write("<p><a href=""../html/"&Temp&""" target=""_blank"">“"&ArticleTitle&"”生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			End If
		Else
			Response.Write(ArticleTitle&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatArticleCategory
	'作    用：生成文章分类页
	'参    数：
	'==============================
	Public Function CreatArticleCategory(TemplateId,ModuleDir,ModuleName)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='articlelist'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			PageCodes=PageCode
			Rs.Close
			Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Show=1 And (Fk_Article_Module="&Id&" Or Fk_Module_LevelList Like '%%,"&Id&",%%')"
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				Rs.PageSize=PageSizes
				Pages=Rs.PageCount
			Else
				Pages=1
			End If
			Rs.Close
			For j=1 To Pages
				PageCode=PageCodes
				PageNow=j
				PageCode=FKTemplate.FileChange(PageCode)
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.ArticleListChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				If ModuleDir<>"" Then
					Temp=ModuleDir&"/"
				Else
					Temp="Article"&Id&"/"
				End If
				Temp="../html/"&Temp
				CategoryDirName=Temp
				PageCode=Replace(PageCode,"{$ArticleCategoryPage$}",FKTemplate.ShowPageCode(Temp&"Index_{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
				PageCode=FKTemplate.PageCodeChange(PageCode)
				If j=1 Then
					Temp=Temp&"Index.html"
				Else
					Temp=Temp&"Index_"&j&".html"
				End If
				Call FKFso.CreateFile(Temp,PageCode)
				Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”第"&j&"页生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			Next
		Else
			Response.Write(ModuleName&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatProduct
	'作    用：生成产品页
	'参    数：
	'==============================
	Public Function CreatProduct(TemplateId,ModuleId,ModuleDir,ProductFileName,ProductTitle,CType)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & ModuleId
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				If Rs("Fk_Module_LowTemplate")>0 Then
					TemplateId=Rs("Fk_Module_LowTemplate")
				End If
			End If
			Rs.Close
		End If
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='product'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}",ModuleId)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.ProductChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(ModuleId))
			If ModuleDir<>"" Then
				Temp=ModuleDir&"/"
			Else
				Temp="Product"&ModuleId&"/"
			End If
			If ProductFileName<>"" Then
				Temp=Temp&ProductFileName&".html"
			Else
				Temp=Temp&Id&".html"
			End If
			Temp="../html/"&Temp
			Call FKFso.CreateFile(Temp,PageCode)
			If CType=0 Then
				Response.Write("<p><a href=""../html/"&Temp&""" target=""_blank"">“"&ProductTitle&"”生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			End If
		Else
			Response.Write(ProductTitle&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatProductCategory
	'作    用：生成产品分类页
	'参    数：
	'==============================
	Public Function CreatProductCategory(TemplateId,ModuleDir,ModuleName)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='productlist'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			PageCodes=PageCode
			Rs.Close
			Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Show=1 And (Fk_Product_Module="&Id&" Or Fk_Module_LevelList Like '%%,"&Id&",%%')"
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				Rs.PageSize=PageSizes
				Pages=Rs.PageCount
			Else
				Pages=1
			End If
			Rs.Close
			For j=1 To Pages
				PageCode=PageCodes
				PageNow=j
				PageCode=FKTemplate.FileChange(PageCode)
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.ProductListChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				If ModuleDir<>"" Then
					Temp=ModuleDir&"/"
				Else
					Temp="Product"&Id&"/"
				End If
				Temp="../html/"&Temp
				CategoryDirName=Temp
				PageCode=Replace(PageCode,"{$ProductCategoryPage$}",FKTemplate.ShowPageCode(Temp&"Index_{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
				PageCode=FKTemplate.PageCodeChange(PageCode)
				If j=1 Then
					Temp=Temp&"Index.html"
				Else
					Temp=Temp&"Index_"&j&".html"
				End If
				Call FKFso.CreateFile(Temp,PageCode)
				Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”第"&j&"页生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			Next
		Else
			Response.Write(ModuleName&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatDown
	'作    用：生成下载页
	'参    数：
	'==============================
	Public Function CreatDown(TemplateId,ModuleId,ModuleDir,DownFileName,DownTitle,CType)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & ModuleId
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				If Rs("Fk_Module_LowTemplate")>0 Then
					TemplateId=Rs("Fk_Module_LowTemplate")
				End If
			End If
			Rs.Close
		End If
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='down'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			Rs.Close
			PageCode=FKTemplate.FileChange(PageCode)
			PageCode=Replace(PageCode,"{$SitePageNow$}",ModuleId)
			PageCode=FKTemplate.SiteChange(PageCode)
			PageCode=FKTemplate.DownChange(PageCode)
			PageCode=FKTemplate.TemplateDo(PageCode)
			PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(ModuleId))
			If ModuleDir<>"" Then
				Temp=ModuleDir&"/"
			Else
				Temp="Down"&ModuleId&"/"
			End If
			If DownFileName<>"" Then
				Temp=Temp&DownFileName&".html"
			Else
				Temp=Temp&Id&".html"
			End If
			Temp="../html/"&Temp
			Call FKFso.CreateFile(Temp,PageCode)
			If CType=0 Then
				Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&DownTitle&"”生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			End If
		Else
			Response.Write(DownTitle&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
	
	'==============================
	'函 数 名：CreatDownCategory
	'作    用：生成下载分类页
	'参    数：
	'==============================
	Public Function CreatDownCategory(TemplateId,ModuleDir,ModuleName)
		If TemplateId=0 Then
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='downlist'"
		Else
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,3
		If Not Rs.Eof Then
			PageCode=Rs("Fk_Template_Content")
			PageCodes=PageCode
			Rs.Close
			Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Show=1 And (Fk_Down_Module="&Id&" Or Fk_Module_LevelList Like '%%,"&Id&",%%')"
			Rs.Open Sqlstr,Conn,1,3
			If Not Rs.Eof Then
				Rs.PageSize=PageSizes
				Pages=Rs.PageCount
			Else
				Pages=1
			End If
			Rs.Close
			For j=1 To Pages
				PageCode=PageCodes
				PageNow=j
				PageCode=FKTemplate.FileChange(PageCode)
				PageCode=Replace(PageCode,"{$SitePageNow$}",Id)
				PageCode=FKTemplate.SiteChange(PageCode)
				PageCode=FKTemplate.DownListChange(PageCode)
				PageCode=FKTemplate.TemplateDo(PageCode)
				PageCode=Replace(PageCode,"{$PageNows$}",FKTemplate.PageNows(Id))
				If ModuleDir<>"" Then
					Temp=ModuleDir&"/"
				Else
					Temp="Down"&Id&"/"
				End If
				Temp="../html/"&Temp
				CategoryDirName=Temp
				PageCode=Replace(PageCode,"{$DownCategoryPage$}",FKTemplate.ShowPageCode(Temp&"Index_{Pages}.html",PageNow,PageAll,PageSizes,PageCounts))
				PageCode=FKTemplate.PageCodeChange(PageCode)
				If j=1 Then
					Temp=Temp&"Index.html"
				Else
					Temp=Temp&"Index_"&j&".html"
				End If
				Call FKFso.CreateFile(Temp,PageCode)
				Response.Write("<p><a href="""&Temp&""" target=""_blank"">“"&ModuleName&"”第"&j&"页生成成功</a></p>")
				Response.Flush()
				Response.Clear()
			Next
		Else
			Response.Write(ModuleName&"模板获取失败！<br />")
			Rs.Close
		End If
	End Function
End Class
%>
