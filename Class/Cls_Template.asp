<%
'==========================================
'文 件 名：Cls_Template.asp
'文件用途：模板引擎函数类
'==========================================

Class Cls_Template
	Private TemplateTag,TemplatePar,TemplateBCode
	Private If1,If2
	'==============================
	'函 数 名：SiteChange
	'作    用：替换站点参数
	'参    数：
	'==============================
	Public Function SiteChange(TemplateCode)
		TemplateCode=Replace(TemplateCode,"{$SiteName$}",SiteName)
		TemplateCode=Replace(TemplateCode,"{$SiteUrl$}",SiteUrl)
		TemplateCode=Replace(TemplateCode,"{$SiteSeoTitle$}",SiteSeoTitle)
		TemplateCode=Replace(TemplateCode,"{$SiteKeyword$}",SiteKeyword)
		TemplateCode=Replace(TemplateCode,"{$SiteDescription$}",SiteDescription)
		TemplateCode=Replace(TemplateCode,"{$SiteSkin$}",SiteDir&"Skin/"&SiteTemplate&"/")
		TemplateCode=Replace(TemplateCode,"{$SiteDir$}",SiteDir)
		TemplateCode=Replace(TemplateCode,"{$SystemName$}",FkSystemNameEn)
		TemplateCode=Replace(TemplateCode,"{$SystemVersion$}",FkSystemVersion)
		TemplateCode=Replace(TemplateCode,"{$ImgCdnUrl$}",ImgCdnUrl)
		TemplateCode=Replace(TemplateCode,"{$CssCdnUrl$}",CssCdnUrl)
		TemplateCode=Replace(TemplateCode,"{$JsCdnUrl$}",JsCdnUrl)
		TemplateCode=Replace(TemplateCode,"{$FileCdnUrl$}",FileCdnUrl)
		
		TemplateCode=Replace(TemplateCode,"{$Tel$}",Tel)
		TemplateCode=Replace(TemplateCode,"{$Tel400$}",Tel400)
		TemplateCode=Replace(TemplateCode,"{$Fax$}",Fax)
		TemplateCode=Replace(TemplateCode,"{$Add$}",Add)
		TemplateCode=Replace(TemplateCode,"{$Lianxiren$}",Lianxiren)
		TemplateCode=Replace(TemplateCode,"{$Beian$}",Beian)
		TemplateCode=Replace(TemplateCode,"{$Email$}",Email)
		TemplateCode=Replace(TemplateCode,"{$Tjid$}",Tjid)
		TemplateCode=Replace(TemplateCode,"{$Kfid$}",Kfid)
		TemplateCode=Replace(TemplateCode,"{$TjUrl$}",TjUrl)
		TemplateCode=Replace(TemplateCode,"{$KfUrl$}",KfUrl)
		TemplateCode=Replace(TemplateCode,"{$SiteLogo$}",SiteLogo)
		TemplateCode=Replace(TemplateCode,"{$Sitepic1$}",Sitepic1)
		TemplateCode=Replace(TemplateCode,"{$Sitepic2$}",Sitepic2)
		TemplateCode=Replace(TemplateCode,"{$Sitepic3$}",Sitepic3)
		TemplateCode=Replace(TemplateCode,"{$Sitepic4$}",Sitepic4)
		TemplateCode=Replace(TemplateCode,"{$Sitepic5$}",Sitepic5)
		TemplateCode=Replace(TemplateCode,"{$Sitepicurl1$}",Sitepicurl1)
		TemplateCode=Replace(TemplateCode,"{$Sitepicurl2$}",Sitepicurl2)
		TemplateCode=Replace(TemplateCode,"{$Sitepicurl3$}",Sitepicurl3)
		TemplateCode=Replace(TemplateCode,"{$Sitepicurl4$}",Sitepicurl4)
		TemplateCode=Replace(TemplateCode,"{$Sitepicurl5$}",Sitepicurl5)
		TemplateCode=Replace(TemplateCode,"{$Sitepictext1$}",Sitepictext1)
		TemplateCode=Replace(TemplateCode,"{$Sitepictext2$}",Sitepictext2)
		TemplateCode=Replace(TemplateCode,"{$Sitepictext3$}",Sitepictext3)
		TemplateCode=Replace(TemplateCode,"{$Sitepictext4$}",Sitepictext4)
		TemplateCode=Replace(TemplateCode,"{$Sitepictext5$}",Sitepictext5)
		SiteChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：PageCodeChange
	'作    用：页码参数参数
	'参    数：
	'==============================
	Public Function PageCodeChange(TemplateCode)
		TemplateCode=Replace(TemplateCode,"{$PageFirst$}",PageFirst)
		TemplateCode=Replace(TemplateCode,"{$PagePrev$}",PagePrev)
		TemplateCode=Replace(TemplateCode,"{$PageNext$}",PageNext)
		TemplateCode=Replace(TemplateCode,"{$PageLast$}",PageLast)
		TemplateCode=Replace(TemplateCode,"{$PageNow$}",PageNow)
		TemplateCode=Replace(TemplateCode,"{$PageCount$}",PageCounts)
		TemplateCode=Replace(TemplateCode,"{$PageRecordCount$}",PageAll)
		TemplateCode=Replace(TemplateCode,"{$PageSize$}",PageSizes)
		PageCodeChange=TemplateCode
	End Function

	'==============================
	'函 数 名：PageChange
	'作    用：替换信息页参数
	'参    数：
	'==============================
	Public Function PageChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=0 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$PageId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$PageTitle$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$PageKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=Replace(TemplateCode,"{$PageDescription$}",Rs("Fk_Module_Description"))
		End If
		Rs.Close
		PageChange=TemplateCode
	End Function

	'==============================
	'函 数 名：JobChange
	'作    用：替换招聘页参数
	'参    数：
	'==============================
	Public Function JobChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=6 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,conn,1,1
		If Not Rs.Eof Then
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$JobId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$JobTitle$}",Rs("Fk_Module_Name"))
		End If
		Rs.Close
		JobChange=TemplateCode
	End Function

	'==============================
	'函 数 名：PageNows
	'作    用：当前位置
	'参    数：
	'==============================
	Public Function PageNows(ModuleId)
		Dim ModuleIds
		ModuleIds=ModuleId
		PageNows=""
		While ModuleIds>0
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & ModuleIds
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				ModuleIds=Rs("Fk_Module_Level")
				If Rs("Fk_Module_Type")=5 Then
					PageNows="&nbsp;&nbsp;&raquo;&nbsp;&nbsp;"&"<a href="""&Rs("Fk_Module_Url")&""" title="""&Rs("Fk_Module_Name")&""">"&Rs("Fk_Module_Name")&"</a>"&PageNows
				Else
				   '----------加判断单页栏目的当前页面地址----------------------------
					If Rs("Fk_Module_Type")=3 and Rs("Fk_Module_Url")<>"" Then
					PageNows="&nbsp;&nbsp;&raquo;&nbsp;&nbsp;"&"<a href="""&Rs("Fk_Module_Url")&""" title="""&Rs("Fk_Module_Name")&""">"&Rs("Fk_Module_Name")&"</a>"&PageNows
					else
					PageNows="&nbsp;&nbsp;&raquo;&nbsp;&nbsp;"&"<a href="""&GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName"))&""" title="""&Rs("Fk_Module_Name")&""">"&Rs("Fk_Module_Name")&"</a>"&PageNows
					end if
					'---------------------------------------------------------
				End If
				
			Else
				ModuleIds=0
			End If
			Rs.Close
		Wend
		PageNows="<a href='/'>"&arrTips(2)&"</a>"&PageNows
	End Function
	
	
	'==============================
	'函 数 名：FirstLevel
	'作    用：获取顶级栏目数据
	'参    数：
	'==============================
	Public Function FirstLevel(ModuleId)
		Dim ModuleIds,info(1),rsf
		ModuleIds=ModuleId
		set rsf=createobject("adodb.recordset")
		While ModuleIds>0
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & ModuleIds
			rsf.Open Sqlstr,Conn,1,1
			If Not rsf.Eof Then
				ModuleIds=rsf("Fk_Module_Level")
				if ModuleIds = 0 then
					info(0) = rsf("Fk_Module_Id")
					info(1) = rsf("Fk_Module_Name")
					FirstLevel=info
					exit function
				end if
			Else
				ModuleIds=0
				info(0) = ""
				info(1) = ""
				FirstLevel=info
			End If
			rsf.Close
		Wend
	End Function

	'==============================
	'函 数 名：InfoChange
	'作    用：替换信息页参数
	'参    数：
	'==============================
	Public Function InfoChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=3 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,conn,1,1
		If Not Rs.Eof Then
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$InfoId$}",Rs("Fk_Module_Id"))
			
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Module_Seotitle","Fk_Module")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Module add column Fk_Module_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Module_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Name"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$InfoTitle$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$InfoKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=Replace(TemplateCode,"{$InfoDescription$}",Rs("Fk_Module_Description"))
			Temp=trim(Rs("Fk_Module_Content")&" ")
			Temp=AddInnerLink(Temp)
			TemplateCode=Replace(TemplateCode,"{$InfoContent$}",Temp)
		End If
		Rs.Close
		InfoChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：ArticleListChange
	'作    用：替换文章列表页参数
	'参    数：
	'==============================
	Public Function ArticleListChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=1 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,conn,1,1
		If Not Rs.Eof Then
			If Rs("Fk_Module_Dir")<>"" Then
				CategoryDirName=Rs("Fk_Module_Dir")
			Else
				CategoryDirName="Article"&Rs("Fk_Module_Id")
			End If
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			if instr(TemplateCode,"{$ModuleContent$}")>0 then
				
				If Len(Rs("Fk_Module_Content"))>0 then
					Temp=Rs("Fk_Module_Content")
					Temp=AddInnerLink(Temp)
					TemplateCode=Replace(TemplateCode,"{$ModuleContent$}",Temp)
				Else
					TemplateCode=Replace(TemplateCode,"{$ModuleContent$}","")
				End if
				
			end if
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Module_Seotitle","Fk_Module")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Module add column Fk_Module_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Module_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Name"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$ArticleCategoryName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ArticleCategoryId$}",Id)
			TemplateCode=Replace(TemplateCode,"{$ArticleCategoryKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=Replace(TemplateCode,"{$ArticleCategoryDescription$}",Rs("Fk_Module_Description"))
		End If
		Rs.Close
		ArticleListChange=TemplateCode
	End Function
	

	'==============================
	'函	   数：MoreUrlChange
	'作    用：首页more链接替换
	'参    数：
	'==============================
	Public Function MoreUrlChange(TemplateCode)
		while InStr(TemplateCode,"{$HomeUrlMore(")
			Temp=clng(Split(Split(TemplateCode,"{$HomeUrlMore(")(1),")$}")(0))
			if Temp>0 then
			Sqlstr="Select Fk_Module_Url,Fk_Module_Type,Fk_Module_Dir,Fk_Module_FileName From [Fk_Module] Where Fk_Module_Id=" & Temp
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				If rs("Fk_Module_Url")<>"" then
					TemplateCode=Replace(TemplateCode,"{$HomeUrlMore("&Temp&")$}",rs("Fk_Module_Url"))
				else
					TemplateCode=Replace(TemplateCode,"{$HomeUrlMore("&Temp&")$}",GetGoUrl(Rs("Fk_Module_Type"),Temp,Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))

				End if
			else
				TemplateCode=Replace(TemplateCode,"{$HomeUrlMore("&Temp&")$}","")
			End If
			Rs.Close
			end if
		wend
		MoreUrlChange=TemplateCode
	End Function

	'==============================
	'函 数 名：ProductListChange
	'作    用：替换产品列表页参数
	'参    数：
	'==============================
	Public Function ProductListChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=2 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,conn,1,1
		If Not Rs.Eof Then
			If Rs("Fk_Module_Dir")<>"" Then
				CategoryDirName=Rs("Fk_Module_Dir")
			Else
				CategoryDirName="Product"&Rs("Fk_Module_Id")
			End If
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			
			if instr(TemplateCode,"{$ModuleContent$}")>0 then
				
				If Len(Rs("Fk_Module_Content"))>0 then
					Temp=Rs("Fk_Module_Content")
					Temp=AddInnerLink(Temp)
					TemplateCode=Replace(TemplateCode,"{$ModuleContent$}",Temp)
				Else
					TemplateCode=Replace(TemplateCode,"{$ModuleContent$}","")
				End if
				
			end if
			
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Module_Seotitle","Fk_Module")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Module add column Fk_Module_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Module_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Name"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$ProductCategoryName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ProductCategoryId$}",Id)
			TemplateCode=Replace(TemplateCode,"{$ProductCategoryKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=Replace(TemplateCode,"{$ProductCategoryDescription$}",Rs("Fk_Module_Description"))
		End If
		Rs.Close
		ProductListChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：DownListChange
	'作    用：替换下载列表页参数
	'参    数：
	'==============================
	Public Function DownListChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=7 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,conn,1,1
		If Not Rs.Eof Then
			If Rs("Fk_Module_Dir")<>"" Then
				CategoryDirName=Rs("Fk_Module_Dir")
			Else
				CategoryDirName="Down"&Rs("Fk_Module_Id")
			End If
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Module_Seotitle","Fk_Module")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Module add column Fk_Module_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Module_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$ModuleTitle$}",Rs("Fk_Module_Name"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$DownCategoryName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$DownCategoryId$}",Id)
			TemplateCode=Replace(TemplateCode,"{$DownCategoryKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=Replace(TemplateCode,"{$DownCategoryDescription$}",Rs("Fk_Module_Description"))
		End If
		Rs.Close
		DownListChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：ArticleChange
	'作    用：替换文章页参数
	'参    数：
	'==============================
	Public Function ArticleChange(TemplateCode)
		Dim ArticleUrl
		Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Article_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If Rs("Fk_Article_Field")<>"" Then
				TemplateTempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
				For Each TemplateTemp In TemplateTempArr
					TemplateCode=Replace(TemplateCode,"{$Article_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
				Next
			End If
			Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
			Rst.Open Sqlstr,Conn,1,1
			while not Rst.Eof
				TemplateCode=Replace(TemplateCode,"{$Article_"&Rst("Fk_Field_Tag")&"$}","")
				Rst.MoveNext
			Wend
			Rst.Close
			Sqlstr="Select  Fk_Article_Id,Fk_Article_Url,Fk_Module_Dir,Fk_Module_Id,Fk_Article_FileName,Fk_Article_Title From [Fk_ArticleList] Where Fk_Article_Show=1 and  Fk_Module_Id="&Rs("Fk_Module_Id")&" Order By Fk_Article_Ip desc, Px desc,Fk_Article_Id Desc"
			set Rst=Conn.execute(Sqlstr)
			dim arr_rows,ub_arr_rows,arr_i
			if not Rst.eof then
				arr_rows=Rst.getrows()
			else
				TemplateCode=Replace(TemplateCode,"{$ArticlePrevTitle$}",arrTips(4))
				TemplateCode=Replace(TemplateCode,"{$ArticlePrevUrl$}","#")
			end if
			Rst.close
			ub_arr_rows=ubound(arr_rows,2)
			if ub_arr_rows>=0 then
				for arr_i=0 to ub_arr_rows
					if arr_rows(0,arr_i)=Id then
						if arr_i=0 then
							TemplateCode=Replace(TemplateCode,"{$ArticlePrevTitle$}",arrTips(4))
							TemplateCode=Replace(TemplateCode,"{$ArticlePrevUrl$}","#")
							if arr_i>ub_arr_rows-1 then
								TemplateCode=Replace(TemplateCode,"{$ArticleNextTitle$}",arrTips(5))
								TemplateCode=Replace(TemplateCode,"{$ArticleNextUrl$}","#")
							else
								If arr_rows(1,arr_i+1)<>"" Then
									ArticleUrl=arr_rows(1,arr_i+1)
								Else
									If arr_rows(2,arr_i+1)<>"" Then
										ArticleUrl=arr_rows(2,arr_i+1)&"/"
									Else
										ArticleUrl="Article"&arr_rows(3,arr_i+1)&"/"
									End If
									If arr_rows(4,arr_i+1)<>"" Then
										ArticleUrl=ArticleUrl&arr_rows(4,arr_i+1)&".html"
									Else
										ArticleUrl=ArticleUrl&arr_rows(0,arr_i+1)&".html"
									End If
									If SiteHtml=1 and sitetemplate<>"wap" Then
										ArticleUrl="/html"&SiteDir&ArticleUrl
									Else
										ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
									End If
								End If
								TemplateCode=Replace(TemplateCode,"{$ArticleNextTitle$}",arr_rows(5,arr_i+1))
								TemplateCode=Replace(TemplateCode,"{$ArticleNextUrl$}",ArticleUrl)
							end if
						elseif arr_i=ub_arr_rows then
							TemplateCode=Replace(TemplateCode,"{$ArticleNextTitle$}",arrTips(5))
							TemplateCode=Replace(TemplateCode,"{$ArticleNextUrl$}","#")
							If arr_rows(1,arr_i-1)<>"" Then
								ArticleUrl=arr_rows(1,arr_i-1)
							Else
								If arr_rows(2,arr_i-1)<>"" Then
									ArticleUrl=arr_rows(2,arr_i-1)&"/"
								Else
									ArticleUrl="Article"&arr_rows(3,arr_i-1)&"/"
								End If
								If arr_rows(4,arr_i-1)<>"" Then
									ArticleUrl=ArticleUrl&arr_rows(4,arr_i-1)&".html"
								Else
									ArticleUrl=ArticleUrl&arr_rows(0,arr_i-1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									ArticleUrl="/html"&SiteDir&ArticleUrl
								Else
									ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$ArticlePrevTitle$}",arr_rows(5,arr_i-1))
							TemplateCode=Replace(TemplateCode,"{$ArticlePrevUrl$}",ArticleUrl)
						else
							If arr_rows(1,arr_i+1)<>"" Then
								ArticleUrl=arr_rows(1,arr_i+1)
							Else
								If arr_rows(2,arr_i+1)<>"" Then
									ArticleUrl=arr_rows(2,arr_i+1)&"/"
								Else
									ArticleUrl="Article"&arr_rows(3,arr_i+1)&"/"
								End If
								If arr_rows(4,arr_i+1)<>"" Then
									ArticleUrl=ArticleUrl&arr_rows(4,arr_i+1)&".html"
								Else
									ArticleUrl=ArticleUrl&arr_rows(0,arr_i+1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									ArticleUrl="/html"&SiteDir&ArticleUrl
								Else
									ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$ArticleNextTitle$}",arr_rows(5,arr_i+1))
							TemplateCode=Replace(TemplateCode,"{$ArticleNextUrl$}",ArticleUrl)
							
							If arr_rows(1,arr_i-1)<>"" Then
								ArticleUrl=arr_rows(1,arr_i-1)
							Else
								If arr_rows(2,arr_i-1)<>"" Then
									ArticleUrl=arr_rows(2,arr_i-1)&"/"
								Else
									ArticleUrl="Article"&arr_rows(3,arr_i-1)&"/"
								End If
								If arr_rows(4,arr_i-1)<>"" Then
									ArticleUrl=ArticleUrl&arr_rows(4,arr_i-1)&".html"
								Else
									ArticleUrl=ArticleUrl&arr_rows(0,arr_i-1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									ArticleUrl="/html"&SiteDir&ArticleUrl
								Else
									ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$ArticlePrevTitle$}",arr_rows(5,arr_i-1))
							TemplateCode=Replace(TemplateCode,"{$ArticlePrevUrl$}",ArticleUrl)
						end if
					end if
				next
			end if
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$ArticleId$}",Id)
			TemplateCode=Replace(TemplateCode,"{$ArticlePic$}",trim(Rs("Fk_Article_Pic")&" "))
			If Not IsNull(Rs("Fk_Article_PicBig")) Then
				TemplateCode=Replace(TemplateCode,"{$ArticlePicBig$}",Rs("Fk_Article_PicBig"))
			End If
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Article_Seotitle","Fk_Article")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Article add column Fk_Article_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Article_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$ArticleSeoTitle$}",Rs("Fk_Article_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$ArticleSeoTitle$}",Rs("Fk_Article_Title"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$ArticleTitle$}",Rs("Fk_Article_Title"))
			TemplateCode=Replace(TemplateCode,"{$ArticleFrom$}",Rs("Fk_Article_From"))
			TemplateCode=Replace(TemplateCode,"{$ArticleModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ArticleModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ArticleTime$}",Rs("Fk_Article_Time"))
			TemplateCode=Replace(TemplateCode,"{$ArticleKeyword$}",Trim(Rs("Fk_Article_Keyword")&" "))
			TemplateCode=Replace(TemplateCode,"{$ArticleDescription$}",Trim(Rs("Fk_Article_Description")&" "))
			TemplateCode=Replace(TemplateCode,"{$ArticleClick$}","<span id=""Click""></span>")
			
			'新功能，追加转载声明
			'2014年12月31日
			'middy241@163.com
			if CheckFields("Fk_Article_Copyright","Fk_Article")=false then
				on error resume next
				Application.Lock()
					conn.execute("alter table Fk_Article add column Fk_Article_Copyright int default 0")
					conn.execute("alter table Fk_Article add column Fk_Article_CopyrightInfo varchar(200) null")
					conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFs varchar(50) null")
					conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFt varchar(50) null")
					conn.execute("alter table Fk_Article add column Fk_Article_CopyrightCl varchar(50) null")
				Application.unLock()
				err.clear
			end if
			Temp=Rs("Fk_Article_Content")
			Temp=AddInnerLink(Temp)
			if rs("Fk_Article_Copyright")=1 then
				If Rs("Fk_Article_Url")<>"" Then
					ArticleUrl=Rs("Fk_Article_Url")
				Else
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
					If SiteHtml=1 and sitetemplate<>"wap" Then
						ArticleUrl="/html"&SiteDir&ArticleUrl
					Else
						ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
					End If
				End If
				Temp=Temp&"<div class='article_copyright_class' style='margin-top:10px;padding-top:5px;word-wrap:break-word;word-break:break-all;border-top:dashed 1px #ccc;font-size:"&Rs("Fk_Article_CopyrightFs")&";font-weight:"&Rs("Fk_Article_CopyrightFt")&";color:"&Rs("Fk_Article_CopyrightCl")&"'>"&Rs("Fk_Article_CopyrightInfo")&"</div>"
				Temp=Replace(Temp,"{$originalUrl}","http://"&Request.ServerVariables("Server_name")&ArticleUrl)
			end if
			Rs.Close
			TemplateCode=Replace(TemplateCode,"{$ArticleContent$}",Temp)
		Else
			rs.close
		End If

		ArticleChange=TemplateCode
	End Function
	
	private Function CheckFields(FieldsName,TableName)
		dim blnFlag,chkStrSql,chkStrRs
		blnFlag=False
		chkStrSql="select * from "&TableName
		Set chkStrRs=Conn.Execute(chkStrSql)
		for i = 0 to chkStrRs.Fields.Count - 1
			if lcase(chkStrRs.Fields(i).Name)=lcase(FieldsName) then
				blnFlag=True
				Exit For
			else
				blnFlag=False
			end if
		Next
		CheckFields=blnFlag
		chkStrRs.close
		set chkStrRs=nothing
	End Function
	
	'==============================
	'函 数 名：ProductChange
	'作    用：替换产品页参数
	'参    数：
	'==============================
	Public Function ProductChange(TemplateCode)
		Dim ProductUrl
		Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Product_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If Rs("Fk_Product_Field")<>"" Then
				TemplateTempArr=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
				For Each TemplateTemp In TemplateTempArr
					TemplateCode=Replace(TemplateCode,"{$Product_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
				Next
			End If
			Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
			Rst.Open Sqlstr,Conn,1,1
			while not Rst.Eof
				TemplateCode=Replace(TemplateCode,"{$Product_"&Rst("Fk_Field_Tag")&"$}","")
				Rst.MoveNext
			Wend
			Rst.Close
			Sqlstr="Select  Fk_Product_Id,Fk_Product_Url,Fk_Module_Dir,Fk_Module_Id,Fk_Product_FileName,Fk_Product_Title From [Fk_ProductList] Where Fk_Product_Show=1 and  Fk_Module_Id="&Rs("Fk_Module_Id")&" Order By Fk_Product_Ip desc, Px desc,Fk_Product_Id Desc"
			set Rst=Conn.execute(Sqlstr)
			dim arr_rows,ub_arr_rows,arr_i
			if not Rst.eof then
				arr_rows=Rst.getrows()
			else
				TemplateCode=Replace(TemplateCode,"{$ProductPrevTitle$}",arrTips(4))
				TemplateCode=Replace(TemplateCode,"{$ProductPrevUrl$}","#")
			end if
			Rst.close
			ub_arr_rows=ubound(arr_rows,2)
			if ub_arr_rows>=0 then
				for arr_i=0 to ub_arr_rows
					if arr_rows(0,arr_i)=Id then
						if arr_i=0 then
							TemplateCode=Replace(TemplateCode,"{$ProductPrevTitle$}",arrTips(4))
							TemplateCode=Replace(TemplateCode,"{$ProductPrevUrl$}","#")
							if arr_i>ub_arr_rows-1 then
								TemplateCode=Replace(TemplateCode,"{$ProductNextTitle$}",arrTips(5))
								TemplateCode=Replace(TemplateCode,"{$ProductNextUrl$}","#")
							else
								If arr_rows(1,arr_i+1)<>"" Then
									ProductUrl=arr_rows(1,arr_i+1)
								Else
									If arr_rows(2,arr_i+1)<>"" Then
										ProductUrl=arr_rows(2,arr_i+1)&"/"
									Else
										ProductUrl="Product"&arr_rows(3,arr_i+1)&"/"
									End If
									If arr_rows(4,arr_i+1)<>"" Then
										ProductUrl=ProductUrl&arr_rows(4,arr_i+1)&".html"
									Else
										ProductUrl=ProductUrl&arr_rows(0,arr_i+1)&".html"
									End If
									If SiteHtml=1 and sitetemplate<>"wap" Then
										ProductUrl="/html"&SiteDir&ProductUrl
									Else
										ProductUrl=SiteDir&sTemp&"?"&ProductUrl
									End If
								End If
								TemplateCode=Replace(TemplateCode,"{$ProductNextTitle$}",arr_rows(5,arr_i+1))
								TemplateCode=Replace(TemplateCode,"{$ProductNextUrl$}",ProductUrl)
							end if
						elseif arr_i=ub_arr_rows then
							TemplateCode=Replace(TemplateCode,"{$ProductNextTitle$}",arrTips(5))
							TemplateCode=Replace(TemplateCode,"{$ProductNextUrl$}","#")
							If arr_rows(1,arr_i-1)<>"" Then
								ProductUrl=arr_rows(1,arr_i-1)
							Else
								If arr_rows(2,arr_i-1)<>"" Then
									ProductUrl=arr_rows(2,arr_i-1)&"/"
								Else
									ProductUrl="Product"&arr_rows(3,arr_i-1)&"/"
								End If
								If arr_rows(4,arr_i-1)<>"" Then
									ProductUrl=ProductUrl&arr_rows(4,arr_i-1)&".html"
								Else
									ProductUrl=ProductUrl&arr_rows(0,arr_i-1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									ProductUrl="/html"&SiteDir&ProductUrl
								Else
									ProductUrl=SiteDir&sTemp&"?"&ProductUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$ProductPrevTitle$}",arr_rows(5,arr_i-1))
							TemplateCode=Replace(TemplateCode,"{$ProductPrevUrl$}",ProductUrl)
						else
							If arr_rows(1,arr_i+1)<>"" Then
								ProductUrl=arr_rows(1,arr_i+1)
							Else
								If arr_rows(2,arr_i+1)<>"" Then
									ProductUrl=arr_rows(2,arr_i+1)&"/"
								Else
									ProductUrl="Product"&arr_rows(3,arr_i+1)&"/"
								End If
								If arr_rows(4,arr_i+1)<>"" Then
									ProductUrl=ProductUrl&arr_rows(4,arr_i+1)&".html"
								Else
									ProductUrl=ProductUrl&arr_rows(0,arr_i+1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									ProductUrl="/html"&SiteDir&ProductUrl
								Else
									ProductUrl=SiteDir&sTemp&"?"&ProductUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$ProductNextTitle$}",arr_rows(5,arr_i+1))
							TemplateCode=Replace(TemplateCode,"{$ProductNextUrl$}",ProductUrl)
							
							If arr_rows(1,arr_i-1)<>"" Then
								ProductUrl=arr_rows(1,arr_i-1)
							Else
								If arr_rows(2,arr_i-1)<>"" Then
									ProductUrl=arr_rows(2,arr_i-1)&"/"
								Else
									ProductUrl="Product"&arr_rows(3,arr_i-1)&"/"
								End If
								If arr_rows(4,arr_i-1)<>"" Then
									ProductUrl=ProductUrl&arr_rows(4,arr_i-1)&".html"
								Else
									ProductUrl=ProductUrl&arr_rows(0,arr_i-1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									ProductUrl="/html"&SiteDir&ProductUrl
								Else
									ProductUrl=SiteDir&sTemp&"?"&ProductUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$ProductPrevTitle$}",arr_rows(5,arr_i-1))
							TemplateCode=Replace(TemplateCode,"{$ProductPrevUrl$}",ProductUrl)
						end if
					end if
				next
			end if
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$ProductId$}",Id)
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Product_Seotitle","Fk_Product")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Product add column Fk_Product_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if


			'新功能，产品关联资料下载、产品测试
			'20230201
			
			' Dim Rs2, Arr_Rows, Related_Down_List, Sqlstr1, Sqlstr2, Rs3, Related_Down_Count, Arr_Rows_Video, Related_Video_Count
			
			' Sqlstr1="Select * From [Fk_Down] Where Fk_Relation_Product_Id=" & Id
			' Set Rs2 = Conn.execute(Sqlstr1)
			' If Not Rs2.Eof Then
			' 	Arr_Rows=Rs2.getrows()
			' 	Related_Down_Count = Ubound(Arr_Rows, 2)+1
			' Else
			' 	Related_Down_Count = 0
			' End If
			' Rs2.close
			
			' Sqlstr2="Select * From [Fk_Article] Where Fk_Relation_Product_Id=" & Id
			' Set Rs3 = Conn.execute(Sqlstr2)
			' If Not Rs3.Eof Then
			' 	Arr_Rows_Video=Rs3.getrows()
			' 	Related_Video_Count = Ubound(Arr_Rows_Video, 2)+1
			' Else
			' 	Related_Video_Count = 0
			' End If
			' Rs3.close
			
			if Not IsNull(Rs("Fk_Product_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$ProductSeoTitle$}",Rs("Fk_Product_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$ProductSeoTitle$}",Rs("Fk_Product_Title"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$ProductTitle$}",Rs("Fk_Product_Title"))
			TemplateCode=Replace(TemplateCode,"{$ProductTime$}",Rs("Fk_Product_Time"))
			TemplateCode=Replace(TemplateCode,"{$ProductVideoFile$}",trim(Rs("Fk_Product_Video_File") & " "))
			TemplateCode=Replace(TemplateCode,"{$ProductDate$}",FormatDateTime(Rs("Fk_Product_Time"),2))
			TemplateCode=Replace(TemplateCode,"{$ProductPic$}",trim(Rs("Fk_Product_Pic")&" "))
			If Not IsNull(Rs("Fk_Product_PicBig")) Then
				TemplateCode=Replace(TemplateCode,"{$ProductPicBig$}",Rs("Fk_Product_PicBig"))
			End If
			TemplateCode=Replace(TemplateCode,"{$ProductModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ProductModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ProductKeyword$}",Trim(Rs("Fk_Product_Keyword")&" "))
			TemplateCode=Replace(TemplateCode,"{$ProductDescription$}",Trim(Rs("Fk_Product_Description")&" "))
			TemplateCode=Replace(TemplateCode,"{$ProductClick$}","<span id=""Click""></span>")
			Temp=Rs("Fk_Product_Content")
			
			on error resume next
			'-------------------
			'扩展字段标签替换
			'addtime: 2013-12-07
			'addby ：shark
			'edittime:2014年12月17日
			'editby: shark
			'-------------------
			dim FK_Product_SlidesImgs,FK_Product_summary,FK_Product_SlidesFirst,Fk_Product_ContentEx1,Fk_Product_ContentEx2,TempImg,SlidesImgs,Fk_Product_Detail
			FK_Product_SlidesImgs=Trim(Rs("FK_Product_SlidesImgs")&" ")
			FK_Product_summary=Trim(Rs("FK_Product_summary")&" ")
			FK_Product_SlidesFirst=Trim(Rs("FK_Product_SlidesFirst")&" ")
			Fk_Product_ContentEx1=Trim(Rs("Fk_Product_ContentEx1")&" ")
			Fk_Product_ContentEx2=Trim(Rs("Fk_Product_ContentEx2")&" ")
			Fk_Product_Detail=Trim(Rs("Fk_Product_Detail")&" ")
			
			if instr(TemplateCode,"{$ProductPicSlidesImgList$}")>0 then
				TemplateCode=Replace(TemplateCode,"{$ProductPicSlidesImgList$}",FK_Product_SlidesImgs)
			end if
			TempImg=""
			if FK_Product_SlidesImgs<>"" then
				SlidesImgs=Split(FK_Product_SlidesImgs,",")
				For j=0 to Ubound(SlidesImgs)
					TempImg=TempImg&"<li><img src="""&Trim(SlidesImgs(j))&""" /></li>"
				Next
			end if
			TemplateCode=Replace(TemplateCode,"{$ProductPicSlidesFirst$}",FK_Product_SlidesFirst)
			
			if instr(TemplateCode,"{$ProductPicSlidesImgs$}")>0 then
				TemplateCode=Replace(TemplateCode,"{$ProductPicSlidesImgs$}",TempImg)
			end if
			if FK_Product_summary="" then
				FK_Product_summary=Left(RemoveHTML(Rs("Fk_Product_Content")),SiteMini)
			end if
			TemplateCode=Replace(TemplateCode,"{$ProductPicSummary$}",FK_Product_summary)
			TemplateCode=Replace(TemplateCode,"{$Product_ContentEx1$}",Fk_Product_ContentEx1)
			TemplateCode=Replace(TemplateCode,"{$Product_ContentEx2$}",Fk_Product_ContentEx2)
			TemplateCode=Replace(TemplateCode,"{$ProductDetail$}",Fk_Product_Detail)
			
			'-------------------------------------------------------
			
			Rs.Close
			Temp=AddInnerLink(Temp)
			TemplateCode=Replace(TemplateCode,"{$ProductContent$}",Temp)
			
			
		Else
			rs.close
		End If

		ProductChange=TemplateCode
	End Function
	
	'--------------
	'函数名：AddInnerLink
	'参数：content:要自动添加关键词的内容
	'作用：自动添加关键词内链
	'说明：每篇文章仅自动添加三个内链，每个关键词仅添加一次内链，按关键词优先级别从高到低判断添加
	'增加：shark
	'时间：2013-06-05
	'--------------
	function AddInnerLink(byval content)
		dim Matches,objRegExp,strs,i,Match,rsadd
		strs=content
		if strs="" then AddInnerLink="":exit function
		Set objRegExp = New Regexp'设置配置对象
		objRegExp.Global = True'设置为全文搜索
		objRegExp.IgnoreCase = True
		objRegExp.Pattern = "(\<a[^<>]+\>.+?\<\/a\>)|(\<img[^<>]+\>)"'
		Set Matches =objRegExp.Execute(strs)'开始执行配置
		'替换正则表达式
		i=0
		Dim MyArray()
		'替换a标签和img标签，排除因这两个标签对关键词造成的干扰
		For Each Match in Matches
			ReDim Preserve MyArray(i)
			MyArray(i)=Mid(Match.Value,1,len(Match.Value))
			strs=replace(strs,Match.Value,"<"&i&">",1,1,1)
			i=i+1
		Next
		dim intKk
		intKk=0
		set rsadd=server.createobject("adodb.recordset")
		Sqlstr="Select Fk_Word_Name,Fk_Word_Url From [Fk_Word] Order By Fk_Word_level Desc"
		rsadd.Open Sqlstr,Conn,1,1
		do while not rsadd.Eof
			if instr(strs,rsadd("Fk_Word_Name"))>0 then
				'response.write Rs("Fk_Word_Name")&"<br/>"
				intKk=intKk+1
				if intKk>3 then	'排除过滤掉的关键词
					exit do
				end if
 				strs=replace(strs,rsadd("Fk_Word_Name"),"<a href="""&rsadd("Fk_Word_Url")&""" target=""_blank"" title="""&rsadd("Fk_Word_Name")&""">"&rsadd("Fk_Word_Name")&"</a>",1,1,1)
			end if
			rsadd.MoveNext
		Loop
		rsadd.close
		set rsadd=nothing
		If i>0 Then 
			'替换回去
			for i=0 to ubound(MyArray)
				strs=replace(strs,"<"&i&">",MyArray(i),1,1,1)
			Next
		End If 
		AddInnerLink=strs
	end function
	
	Function RegExpTest(patrn, strng)
		Dim regEx,RetStr, Match, Matches ' 建立变量。
		Set regEx = New RegExp ' 建立正则表达式。
		regEx.Pattern = patrn ' 设置模式。
		regEx.IgnoreCase = True ' 设置是否区分大小写。
		regEx.Global = True ' 设置全局替换。
		Set Matches = regEx.Execute(strng) ' 执行搜索。
		For Each Match in Matches ' 遍历 Matches 集合。
			RetStr = RetStr & "Match " & I & " found at position "
			RetStr = RetStr & Match.FirstIndex & ". Match Value is "
			RetStr = RetStr & Match.Value & "'.<br>" 
		Next
		RegExpTest = RetStr
	End Function 
	
	'==============================
	'函 数 名：DownChange
	'作    用：替换下载页参数
	'参    数：
	'==============================
	Public Function DownChange(TemplateCode)
		Dim DownUrl
		Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Down_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If Rs("Fk_Down_Field")<>"" Then
				TemplateTempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
				For Each TemplateTemp In TemplateTempArr
					TemplateCode=Replace(TemplateCode,"{$Down_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
				Next
			End If
			Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
			Rst.Open Sqlstr,Conn,1,1
			While Not Rst.Eof
				TemplateCode=Replace(TemplateCode,"{$Down_"&Rst("Fk_Field_Tag")&"$}","")
				Rst.MoveNext
			Wend
			Rst.Close
			Sqlstr="Select  Fk_Down_Id,Fk_Down_Url,Fk_Module_Dir,Fk_Module_Id,Fk_Down_FileName,Fk_Down_Title From [Fk_DownList] Where Fk_Down_Show=1 and  Fk_Module_Id="&Rs("Fk_Module_Id")&" Order By Fk_Down_Ip desc, Px desc,Fk_Down_Id Desc"
			set Rst=Conn.execute(Sqlstr)
			dim arr_rows,ub_arr_rows,arr_i
			if not Rst.eof then
				arr_rows=Rst.getrows()
			else
				TemplateCode=Replace(TemplateCode,"{$DownPrevTitle$}",arrTips(4))
				TemplateCode=Replace(TemplateCode,"{$DownPrevUrl$}","#")
			end if
			Rst.close
			ub_arr_rows=ubound(arr_rows,2)
			if ub_arr_rows>=0 then
				for arr_i=0 to ub_arr_rows
					if arr_rows(0,arr_i)=Id then
						if arr_i=0 then
							TemplateCode=Replace(TemplateCode,"{$DownPrevTitle$}",arrTips(4))
							TemplateCode=Replace(TemplateCode,"{$DownPrevUrl$}","#")
							if arr_i>ub_arr_rows-1 then
								TemplateCode=Replace(TemplateCode,"{$DownNextTitle$}",arrTips(5))
								TemplateCode=Replace(TemplateCode,"{$DownNextUrl$}","#")
							else
								If arr_rows(1,arr_i+1)<>"" Then
									DownUrl=arr_rows(1,arr_i+1)
								Else
									If arr_rows(2,arr_i+1)<>"" Then
										DownUrl=arr_rows(2,arr_i+1)&"/"
									Else
										DownUrl="Down"&arr_rows(3,arr_i+1)&"/"
									End If
									If arr_rows(4,arr_i+1)<>"" Then
										DownUrl=DownUrl&arr_rows(4,arr_i+1)&".html"
									Else
										DownUrl=DownUrl&arr_rows(0,arr_i+1)&".html"
									End If
									If SiteHtml=1 and sitetemplate<>"wap" Then
										DownUrl="/html"&SiteDir&DownUrl
									Else
										DownUrl=SiteDir&sTemp&"?"&DownUrl
									End If
								End If
								TemplateCode=Replace(TemplateCode,"{$DownNextTitle$}",arr_rows(5,arr_i+1))
								TemplateCode=Replace(TemplateCode,"{$DownNextUrl$}",DownUrl)
							end if
						elseif arr_i=ub_arr_rows then
							TemplateCode=Replace(TemplateCode,"{$DownNextTitle$}",arrTips(5))
							TemplateCode=Replace(TemplateCode,"{$DownNextUrl$}","#")
							If arr_rows(1,arr_i-1)<>"" Then
								DownUrl=arr_rows(1,arr_i-1)
							Else
								If arr_rows(2,arr_i-1)<>"" Then
									DownUrl=arr_rows(2,arr_i-1)&"/"
								Else
									DownUrl="Down"&arr_rows(3,arr_i-1)&"/"
								End If
								If arr_rows(4,arr_i-1)<>"" Then
									DownUrl=DownUrl&arr_rows(4,arr_i-1)&".html"
								Else
									DownUrl=DownUrl&arr_rows(0,arr_i-1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									DownUrl="/html"&SiteDir&DownUrl
								Else
									DownUrl=SiteDir&sTemp&"?"&DownUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$DownPrevTitle$}",arr_rows(5,arr_i-1))
							TemplateCode=Replace(TemplateCode,"{$DownPrevUrl$}",DownUrl)
						else
							If arr_rows(1,arr_i+1)<>"" Then
								DownUrl=arr_rows(1,arr_i+1)
							Else
								If arr_rows(2,arr_i+1)<>"" Then
									DownUrl=arr_rows(2,arr_i+1)&"/"
								Else
									DownUrl="Down"&arr_rows(3,arr_i+1)&"/"
								End If
								If arr_rows(4,arr_i+1)<>"" Then
									DownUrl=DownUrl&arr_rows(4,arr_i+1)&".html"
								Else
									DownUrl=DownUrl&arr_rows(0,arr_i+1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									DownUrl="/html"&SiteDir&DownUrl
								Else
									DownUrl=SiteDir&sTemp&"?"&DownUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$DownNextTitle$}",arr_rows(5,arr_i+1))
							TemplateCode=Replace(TemplateCode,"{$DownNextUrl$}",DownUrl)
							
							If arr_rows(1,arr_i-1)<>"" Then
								DownUrl=arr_rows(1,arr_i-1)
							Else
								If arr_rows(2,arr_i-1)<>"" Then
									DownUrl=arr_rows(2,arr_i-1)&"/"
								Else
									DownUrl="Down"&arr_rows(3,arr_i-1)&"/"
								End If
								If arr_rows(4,arr_i-1)<>"" Then
									DownUrl=DownUrl&arr_rows(4,arr_i-1)&".html"
								Else
									DownUrl=DownUrl&arr_rows(0,arr_i-1)&".html"
								End If
								If SiteHtml=1 and sitetemplate<>"wap" Then
									DownUrl="/html"&SiteDir&DownUrl
								Else
									DownUrl=SiteDir&sTemp&"?"&DownUrl
								End If
							End If
							TemplateCode=Replace(TemplateCode,"{$DownPrevTitle$}",arr_rows(5,arr_i-1))
							TemplateCode=Replace(TemplateCode,"{$DownPrevUrl$}",DownUrl)
						end if
					end if
				next
			end if
			
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$DownId$}",Id)
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Down_Seotitle","Fk_Down")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Down add column Fk_Down_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Down_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$DownSeoTitle$}",Rs("Fk_Down_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$DownSeoTitle$}",Rs("Fk_Down_Title"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$DownTitle$}",Rs("Fk_Down_Title"))
			TemplateCode=Replace(TemplateCode,"{$DownLanguage$}",Rs("Fk_Down_Language"))
			TemplateCode=Replace(TemplateCode,"{$DownSystem$}",Rs("Fk_Down_System"))
			TemplateCode=Replace(TemplateCode,"{$DownFile$}",SiteDir&"File.asp?Id="&Rs("Fk_Down_Id"))'Rs("Fk_Down_File"))
			TemplateCode=Replace(TemplateCode,"{$DownTime$}",Rs("Fk_Down_Time"))
			TemplateCode=Replace(TemplateCode,"{$DownDate$}",FormatDateTime(Rs("Fk_Down_Time"),2))
			TemplateCode=Replace(TemplateCode,"{$DownPic$}",trim(Rs("Fk_Down_Pic")&" "))
			If Not IsNull(Rs("Fk_Down_PicBig")) Then
				TemplateCode=Replace(TemplateCode,"{$DownPicBig$}",Rs("Fk_Down_PicBig"))
			End If
			TemplateCode=Replace(TemplateCode,"{$DownModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$DownModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$DownKeyword$}",Trim(Rs("Fk_Down_Keyword")&" "))
			TemplateCode=Replace(TemplateCode,"{$DownDescription$}",Trim(Rs("Fk_Down_Description")&" "))
			TemplateCode=Replace(TemplateCode,"{$DownCount$}","<span id=""Count""></span>")
			TemplateCode=Replace(TemplateCode,"{$DownClick$}","<span id=""Click""></span>")
			Temp=Rs("Fk_Down_Content")
			Rs.Close
			Temp=AddInnerLink(Temp)
			TemplateCode=Replace(TemplateCode,"{$DownContent$}",Temp)
		Else
			rs.close
		End If
		DownChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：GBookChange
	'作    用：替换留言页参数
	'参    数：
	'==============================
	Public Function GBookChange(TemplateCode)
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=4 And Fk_Module_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TemplateCode=Replace(TemplateCode,"{$MenuId$}",Rs("Fk_Module_Menu"))
			TemplateCode=Replace(TemplateCode,"{$ModuleFId$}",Rs("Fk_Module_Level"))
			TemplateCode=Replace(TemplateCode,"{$ModuleId$}",Rs("Fk_Module_Id"))
			TemplateCode=Replace(TemplateCode,"{$ModuleName$}",Rs("Fk_Module_Name"))
			TemplateCode=Replace(TemplateCode,"{$ModuleUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
			TemplateCode=Replace(TemplateCode,"{$GBookModuleId$}",Rs("Fk_Module_Id"))
			
			'新功能，追加SEO title字段
			'2017年5月22日
			'middy241@163.com
			if CheckFields("Fk_Module_Seotitle","Fk_Module")=false then
				on error resume next
				Application.Lock()
				conn.execute("alter table Fk_Module add column Fk_Module_Seotitle varchar(255) null")
				Application.unLock()
				err.clear
			end if
			
			if Not IsNull(Rs("Fk_Module_Seotitle")) then
				TemplateCode=Replace(TemplateCode,"{$GBookTitle$}",Rs("Fk_Module_Seotitle"))
			else
				TemplateCode=Replace(TemplateCode,"{$GBookTitle$}",Rs("Fk_Module_Name"))
			end if
			
			TemplateCode=Replace(TemplateCode,"{$GBookId$}",Id)
			TemplateCode=Replace(TemplateCode,"{$GBookKeyword$}",Rs("Fk_Module_Keyword"))
			TemplateCode=Replace(TemplateCode,"{$GBookDescription$}",Rs("Fk_Module_Description"))
		End If
		Rs.Close
		GBookChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：SubjectChange
	'作    用：专题页参数
	'参    数：
	'==============================
	Public Function SubjectChange(TemplateCode)
		TemplateCode=Replace(TemplateCode,"{$SubjectId$}",Id)
		TemplateCode=Replace(TemplateCode,"{$SubjectName$}",Fk_Subject_Name)
		SubjectChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：FileChange
	'作    用：替换模板模块参数
	'参    数：
	'==============================
	Public Function FileChange(TemplateCode)
		While Instr(TemplateCode,"{$File(")
			Temp=Split(Split(TemplateCode,"{$File(")(1),")$}")(0)
			Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='"&Temp&"'"
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=Replace(TemplateCode,"{$File("&Temp&")$}",Rs("Fk_Template_Content"))
			Else
				TemplateCode=Replace(TemplateCode,"{$File("&Temp&")$}","")
			End If
			Rs.Close
		Wend
		While Instr(TemplateCode,"{$Info(")
			Temp=Split(Split(TemplateCode,"{$Info(")(1),")$}")(0)
			Sqlstr="Select * From [Fk_Info] Where Fk_Info_Id="&Temp&""
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=Replace(TemplateCode,"{$Info("&Temp&")$}",Rs("Fk_Info_Content"))
			Else
				TemplateCode=Replace(TemplateCode,"{$Info("&Temp&")$}","")
			End If
			Rs.Close
		Wend
		'增加首页模块标题调用
		While Instr(TemplateCode,"{$InfoTit(")
			Temp=Split(Split(TemplateCode,"{$InfoTit(")(1),")$}")(0)
			Sqlstr="Select * From [Fk_Info] Where Fk_Info_Id="&Temp&""
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				TemplateCode=Replace(TemplateCode,"{$InfoTit("&Temp&")$}",Rs("Fk_Info_Name"))
			Else
				TemplateCode=Replace(TemplateCode,"{$InfoTit("&Temp&")$}","")
			End If
			Rs.Close
		Wend
		FileChange=TemplateCode
	End Function
	
'========================标签处理区===========================	

	'==============================
	'函 数 名：FkNav
	'作    用：菜单标签操作
	'参    数：
	'==============================
	Private Function FkNav(BCode,BPar)
		Dim NavUrl,NavI,SFor,Temp88,z
		z=1
		SFor=""
		If Instr(BCode,"{$For(Nav")>0 Then
			Temp88=Split(BCode,"{$For(Nav")(0)
			Temp88=Replace(BCode,Temp88,"")
			TempArr=Split(Temp88,"{$Next$}")
			SFor=TempArr(0)&"{$Next$}"
			i=1
			While GetCount(SFor,"{$For")<>GetCount(SFor,"{$Next$}")
				SFor=SFor&TempArr(i)&"{$Next$}"
				i=i+1
			Wend
			Temp88=Split(SFor,")$}")(0)
			SFor=Right(SFor,Len(SFor)-Len(Temp88)-3)
			SFor=Left(SFor,Len(SFor)-8)
			BCode=Replace(BCode,SFor,"{FangkaFor}")
		End If
		TempArr=Split(BPar,"/")
		While TempArr(3)<0 And TempArr(1)>0
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id="&TempArr(1)&""
			Rs.Open Sqlstr,conn,1,1
			If Not Rs.Eof Then
				TempArr(1)=Rs("Fk_Module_Level")
				TempArr(3)=TempArr(3)+1
			Else
				TempArr(1)=0
			End If
			Rs.Close
		Wend
		'response.write TempArr(3)&"<br>"
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Menu="&TempArr(0)&" And Fk_Module_Level="&TempArr(1)&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		NavI=1
		While Not Rs.Eof
			If Rs("Fk_Module_Type")=5 Then
				NavUrl=Rs("Fk_Module_Url")
			Else
				NavUrl=GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName"))
			End If
			if SiteHtml=1 then
			'response.write NavUrl
				NavUrl=NavUrl
			end if
			'单页栏目如果有转向链接，那么导航为转向链接----------------------
			If Rs("Fk_Module_Type")=3 and rs("Fk_Module_Url")<>"" Then
				NavUrl=Rs("Fk_Module_Url")
			end if
			'-----------------------------------------------------------
			FkNav=FkNav&BCode
			FkNav=Replace(FkNav,"{$ListNo$}",z)
			FkNav=Replace(FkNav,"{$NavId$}",Rs("Fk_Module_Id"))
			FkNav=Replace(FkNav,"{$NavI$}",NavI)
			FkNav=Replace(FkNav,"{$NavName$}",Rs("Fk_Module_Name"))
			FkNav=Replace(FkNav,"{$NavModulePic$}",Rs("Fk_Module_Pic")&" ")
			FkNav=Replace(FkNav,"{$NavUr$}",NavUrl)
			FkNav=Replace(FkNav,"{$NavUrl$}",NavUrl)
			FkNav=Replace(FkNav,"{$NavType$}",Rs("Fk_Module_Type"))
			FkNav=Replace(FkNav,"{$Nav_Content$}",trim(Rs("Fk_Module_Content")&" "))
			If TempArr(2)>1 Then
				FkNav=Replace(FkNav,"{$NavSub$}",FkNavs(Rs("Fk_Module_Id"),Clng(TempArr(2))-1))
				'response.write Rs("Fk_Module_Id")&"_"&Clng(TempArr(2))-1&"<br>"
			End If
			Rs.MoveNext
			z=z+1
			NavI=NavI+1
		Wend
		Rs.Close
		FkNav=Replace(FkNav,"{FangkaFor}",SFor)
	End Function

	'==============================
	'函 数 名：FkNavs
	'作    用：读取多级菜单操作
	'参    数：当前父ID GetId，还要读取级数GetCount
	'==============================
	Private Function FkNavs(GetId,GetCount)
		Dim NavUrl,Rs2
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Show=1 And Fk_Module_Level="&GetId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
		Rs2.Open Sqlstr,Conn,1,1
		If Not Rs2.Eof Then
			FkNavs="<ul>" & vbCrLf
			While Not Rs2.Eof
				If Rs2("Fk_Module_Type")=5 Then
					NavUrl=Rs2("Fk_Module_Url")
				Else
					NavUrl=GetGoUrl(Rs2("Fk_Module_Type"),Rs2("Fk_Module_Id"),Rs2("Fk_Module_Dir"),Rs2("Fk_Module_FileName"))
				End If
				'----------单页栏目如果有转向链接，那么导航为转向链接----------------------
				If Rs("Fk_Module_Type")=3 and rs("Fk_Module_Url")<>"" Then
					NavUrl=Rs("Fk_Module_Url")
				end if
				'--------------------------------------------------------------------
				FkNavs=FkNavs&"<li><a href="""&NavUrl&""" title="""&Rs2("Fk_Module_Name")&""""
				If Rs2("Fk_Module_Type")=5 Then
					FkNavs=FkNavs&" target=""_blank"""
				End If
				FkNavs=FkNavs&">"&Rs2("Fk_Module_Name")&"</a>"
				If GetCount>1 Then
					FkNavs=FkNavs&FkNavs(Rs2("Fk_Module_Id"),GetCount-1)
				End If
				FkNavs=FkNavs&"</li>" & vbCrLf
				Rs2.MoveNext
			Wend
			FkNavs=FkNavs&"</ul>" & vbCrLf
		End If
		Rs2.Close
		Set Rs2=Nothing
	End Function
	
	'==============================
	'函 数 名：FkArticleList
	'作    用：文章列表标签操作
	'参    数：
	'==============================
	Private Function FkArticleList(BCode,BPar)
		Dim ArticleUrl,ArticleTitle,z,arr1
		Dim Rst
		Set Rst=Server.Createobject("Adodb.RecordSet")
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(3)>0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(3)&""
		End If
		If TempArr(3)=0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" "
		End If
		Sqlstr=Sqlstr&" * From [Fk_ArticleList] Where Fk_Article_Show=1 And Fk_Module_Menu=" & TempArr(0)
		if not isnumeric(TempArr(1)) then
			arr1=0
		else
			arr1=TempArr(1)
		end if
		If arr1>0  Then
			Sqlstr=Sqlstr&" And (Fk_Article_Module="&TempArr(1)&" Or Fk_Module_LevelList Like '%%,"&TempArr(1)&",%%')"
		End If
		If TempArr(5)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Article_Recommend Like '%%,2,%%' or Fk_Article_Ip= '1')"
		End If
		If TempArr(6)>0 Then
			Sqlstr=Sqlstr&" And Fk_Article_Subject Like '%%,"&TempArr(6)&",%%'"
		End If
		If TempArr(4)=1 And SearchStr<>"" Then
			'Sqlstr=Sqlstr&" And Fk_Article_Title Like '%"&SearchStr&"%'"
			Sqlstr=Sqlstr&" And InStr(1,LCase(Fk_Article_Title),LCase('"&SearchStr&"'),0)<>0"
		End If
		Select Case TempArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Time Desc,Fk_Article_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Click Desc,Fk_Article_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Time Asc,Fk_Article_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Click Asc,Fk_Article_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If TempArr(4)=0 Then
				While Not Rs.Eof
					If Rs("Fk_Article_Url")<>"" Then
						ArticleUrl=Rs("Fk_Article_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							ArticleUrl="/html"&SiteDir&ArticleUrl
						Else
							ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
						End If
					End If
					ArticleTitle=Rs("Fk_Article_Title")
					If Len(ArticleTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						ArticleTitle=Left(ArticleTitle,Clng(TempArr(7)))&"..."
					End If
					FkArticleList=FkArticleList&BCode
					If Rs("Fk_Article_Color")<>"" Then
						FkArticleList=Replace(FkArticleList,"{$ArticleListTitle$}","<span style='color:"&Rs("Fk_Article_Color")&"'>"&ArticleTitle&"</span>")
					Else
						FkArticleList=Replace(FkArticleList,"{$ArticleListTitle$}",ArticleTitle)
					End If
					If Rs("Fk_Article_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkArticleList=Replace(FkArticleList,"{$ArticleList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkArticleList=Replace(FkArticleList,"{$ArticleList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkArticleList=Replace(FkArticleList,"{$ListNo$}",z)
					FkArticleList=Replace(FkArticleList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkArticleList=Replace(FkArticleList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkArticleList,"{$ModuleListUrl$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					If InStr(FkArticleList,"{$ModuleListContent$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If
					
					If InStr(FkArticleList,"{$ModuleListContentNoHtml$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkArticleList=Replace(FkArticleList,"{$ArticleListId$}",Rs("Fk_Article_Id"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListTitleAll$}",Rs("Fk_Article_Title"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListUrl$}",ArticleUrl)
					If InStr(FkArticleList,"{$ArticleListContent$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ArticleListContent$}",Left(RemoveHTML(Rs("Fk_Article_Content")),SiteMini))
					End If
					FkArticleList=Replace(FkArticleList,"{$ArticleListTime$}",Rs("Fk_Article_Time"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListModuleName$}",Rs("Fk_Module_Name"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListModuleId$}",Rs("Fk_Module_Id"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListDate$}",FormatDateTime(Rs("Fk_Article_Time"),2))
					FkArticleList=Replace(FkArticleList,"{$ArticleListYear$}",Year(Rs("Fk_Article_Time")))
					FkArticleList=Replace(FkArticleList,"{$ArticleListMonth$}",Month(Rs("Fk_Article_Time")))
					FkArticleList=Replace(FkArticleList,"{$ArticleListDay$}",Day(Rs("Fk_Article_Time")))
					FkArticleList=Replace(FkArticleList,"{$ArticleListPic$}",trim(Rs("Fk_Article_Pic")&" "))
					If Not IsNull(Rs("Fk_Article_PicBig")) Then
						FkArticleList=Replace(FkArticleList,"{$ArticleListPicBig$}",Rs("Fk_Article_PicBig"))
					End If
					FkArticleList=Replace(FkArticleList,"{$ArticleListNew$}",DateDiff("d",Rs("Fk_Article_Time"),Now()))
					FkArticleList=Replace(FkArticleList,"{$ArticleListClick$}",Rs("Fk_Article_Click"))
					Rs.MoveNext
					z=z+1
				Wend
			Else
				Rs.PageSize=PageSizes
				If PageNow>Rs.PageCount Or PageNow<=0 Then
				'If PageSizes>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				Rs.AbsolutePage=PageNow
				PageAll=Rs.RecordCount
				i=1
				z=PageSizes*(PageNow-1)+1
				While (Not Rs.Eof) And i<PageSizes+1
					If Rs("Fk_Article_Url")<>"" Then
						ArticleUrl=Rs("Fk_Article_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							ArticleUrl="/html"&SiteDir&ArticleUrl
						Else
							ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
						End If
					End If
					ArticleTitle=Rs("Fk_Article_Title")
					If Len(ArticleTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						ArticleTitle=Left(ArticleTitle,Clng(TempArr(7)))&"..."
					End If
					FkArticleList=FkArticleList&BCode
					If Rs("Fk_Article_Color")<>"" Then
						FkArticleList=Replace(FkArticleList,"{$ArticleListTitle$}","<span style='color:"&Rs("Fk_Article_Color")&"'>"&ArticleTitle&"</span>")
					Else
						FkArticleList=Replace(FkArticleList,"{$ArticleListTitle$}",ArticleTitle)
					End If
					If Rs("Fk_Article_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkArticleList=Replace(FkArticleList,"{$ArticleList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkArticleList=Replace(FkArticleList,"{$ArticleList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkArticleList=Replace(FkArticleList,"{$ListNo$}",z)
					FkArticleList=Replace(FkArticleList,"{$ListNo2$}",i)
					FkArticleList=Replace(FkArticleList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkArticleList=Replace(FkArticleList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkArticleList,"{$ModuleListUrl$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					If InStr(FkArticleList,"{$ModuleListContent$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If
					If InStr(FkArticleList,"{$ModuleListContentNoHtml$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					FkArticleList=Replace(FkArticleList,"{$ArticleListId$}",Rs("Fk_Article_Id"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListTitleAll$}",Rs("Fk_Article_Title"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListUrl$}",ArticleUrl)
					If InStr(FkArticleList,"{$ArticleListContent$}")>0 Then
						FkArticleList=Replace(FkArticleList,"{$ArticleListContent$}",Left(RemoveHTML(Rs("Fk_Article_Content")),SiteMini))
					End If
					FkArticleList=Replace(FkArticleList,"{$ArticleListTime$}",Rs("Fk_Article_Time"))
					FkArticleList=Replace(FkArticleList,"{$ArticleListDate$}",FormatDateTime(Rs("Fk_Article_Time"),2))
					FkArticleList=Replace(FkArticleList,"{$ArticleListUrl$}",ArticleUrl)
					FkArticleList=Replace(FkArticleList,"{$ArticleListPic$}",trim(Rs("Fk_Article_Pic")&" "))
					If Not IsNull(Rs("Fk_Article_PicBig")) Then
						FkArticleList=Replace(FkArticleList,"{$ArticleListPicBig$}",Rs("Fk_Article_PicBig"))
					End If
					FkArticleList=Replace(FkArticleList,"{$ArticleListYear$}",Year(Rs("Fk_Article_Time")))
					FkArticleList=Replace(FkArticleList,"{$ArticleListMonth$}",Month(Rs("Fk_Article_Time")))
					FkArticleList=Replace(FkArticleList,"{$ArticleListDay$}",Day(Rs("Fk_Article_Time")))
					FkArticleList=Replace(FkArticleList,"{$ArticleListNew$}",DateDiff("d",Rs("Fk_Article_Time"),Now()))
					FkArticleList=Replace(FkArticleList,"{$ArticleListClick$}",Rs("Fk_Article_Click"))
					Rs.MoveNext
					z=z+1
					i=i+1
				Wend
				' If PageNow>1 Then
					' PageFirst=SiteDir&"?"&CategoryDirName&"/Index.html"
					' PagePrev=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow-1)&".html"
					' If PageNow=2 Then
						' PagePrev=SiteDir&"?"&CategoryDirName&"/Index.html"
					' End If
				' Else
					' PageFirst="#"
					' PagePrev="#"
				' End If
				' If PageCounts>PageNow Then
					' PageNext=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow+1)&".html"
					' PageLast=SiteDir&"?"&CategoryDirName&"/Index_"&PageCounts&".html"
				' Else
					' PageNext="#"
					' PageLast="#"
				' End If
				' If SiteHtml=1 Then
					' PageFirst=Replace(PageFirst,"?","")
					' PagePrev=Replace(PagePrev,"?","")
					' PageNext=Replace(PageNext,"?","")
					' PageLast=Replace(PageLast,"?","")
				' Else
					' PageFirst=Replace(PageFirst,"?",sTemp&"?")
					' PagePrev=Replace(PagePrev,"?",sTemp&"?")
					' PageNext=Replace(PageNext,"?",sTemp&"?")
					' PageLast=Replace(PageLast,"?",sTemp&"?")
				' End If
			End If
		End If
		Rs.Close
	End Function

	'==============================
	'函 数 名：GetModuleContent
	'作    用：列表中获取栏目介绍内容
	'参    数：intModuleid：栏目id
	'==============================
	Private Function GetModuleContent(intModuleid)
		dim modulers
		set modulers=conn.Execute("select Fk_Module_Content from Fk_Module where Fk_Module_Id="&intModuleid)
		if not modulers.EOF Then
			GetModuleContent=trim(modulers("Fk_Module_Content")&" ")
		Else
			GetModuleContent=""
		End If
		modulers.Close
	End Function
		
	'==============================
	'函 数 名：FkProductList
	'作    用：产品列表标签操作
	'参    数：
	'==============================
	Private Function FkProductList(BCode,BPar)
		Dim ProductUrl,ProductTitle,z
		Dim Rst
		Set Rst=Server.Createobject("Adodb.RecordSet")
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(3)>0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(3)&""
		End If
		If TempArr(3)=0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" "
		End If
		Sqlstr=Sqlstr&" * From [Fk_ProductList] Where Fk_Product_Show=1 And Fk_Module_Menu=" & TempArr(0)
		If TempArr(1)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Product_Module="&TempArr(1)&" Or Fk_Module_LevelList Like '%%,"&TempArr(1)&",%%')"
		End If
		If TempArr(5)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Product_Recommend Like '%%,2,%%' or Fk_Product_Ip = '1')"
		End If
		If TempArr(6)>0 Then
			Sqlstr=Sqlstr&" And Fk_Product_Subject Like '%%,"&TempArr(6)&",%%'"
		End If
		If TempArr(4)=1 And SearchStr<>"" Then
			'Sqlstr=Sqlstr&" And Fk_Product_Title Like '%"&SearchStr&"%'"
			Sqlstr=Sqlstr&" And InStr(1,LCase(Fk_Product_Title),LCase('"&SearchStr&"'),0)<>0"
		End If
		Select Case TempArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc, Px desc, Fk_Product_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc, Px desc, Fk_Product_Time Desc,Fk_Product_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc, Px desc, Fk_Product_Click Desc,Fk_Product_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc, Px desc, Fk_Product_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc, Px desc, Fk_Product_Time Asc,Fk_Product_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc, Px desc, Fk_Product_Click Asc,Fk_Product_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If TempArr(4)=0 Then
				While Not Rs.Eof
					If Rs("Fk_Product_Url")<>"" Then
						ProductUrl=Rs("Fk_Product_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							ProductUrl="/html"&SiteDir&ProductUrl
						Else
							ProductUrl=SiteDir&sTemp&"?"&ProductUrl
						End If
					End If
					ProductTitle=Rs("Fk_Product_Title")
					If Len(ProductTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						ProductTitle=Left(ProductTitle,Clng(TempArr(7)))&"..."
					End If
					FkProductList=FkProductList&BCode
					If Rs("Fk_Product_Color")<>"" Then
						FkProductList=Replace(FkProductList,"{$ProductListTitle$}","<span style='color:"&Rs("Fk_Product_Color")&"'>"&ProductTitle&"</span>")
					Else
						FkProductList=Replace(FkProductList,"{$ProductListTitle$}",ProductTitle)
					End If
					If Rs("Fk_Product_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkProductList=Replace(FkProductList,"{$ProductList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,conn,1,1
					While Not Rst.Eof
						FkProductList=Replace(FkProductList,"{$ProductList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkProductList=Replace(FkProductList,"{$ListNo$}",z)
					FkProductList=Replace(FkProductList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkProductList=Replace(FkProductList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkProductList,"{$ModuleListUrl$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),	Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					
					If InStr(FkProductList,"{$ModuleListContent$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If
					
					If InStr(FkProductList,"{$ModuleListContentNoHtml$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkProductList=Replace(FkProductList,"{$ProductListId$}",Rs("Fk_Product_Id"))
					FkProductList=Replace(FkProductList,"{$ProductListTitleAll$}",Rs("Fk_Product_Title"))
					If InStr(FkProductList,"{$ProductListContent$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ProductListContent$}",Left(RemoveHTML(Rs("Fk_Product_Content")),SiteMini))
					End If
					FkProductList=Replace(FkProductList,"{$ProductListUrl$}",ProductUrl)
					FkProductList=Replace(FkProductList,"{$ProductListClick$}",Rs("Fk_Product_Click"))
					FkProductList=Replace(FkProductList,"{$ProductListTime$}",Rs("Fk_Product_Time"))
					FkProductList=Replace(FkProductList,"{$ProductListDate$}",FormatDateTime(Rs("Fk_Product_Time"),2))
					FkProductList=Replace(FkProductList,"{$ProductListYear$}",Year(Rs("Fk_Product_Time")))
					FkProductList=Replace(FkProductList,"{$ProductListMonth$}",Month(Rs("Fk_Product_Time")))
					FkProductList=Replace(FkProductList,"{$ProductListDay$}",Day(Rs("Fk_Product_Time")))
					FkProductList=Replace(FkProductList,"{$ProductListNew$}",DateDiff("d",Rs("Fk_Product_Time"),Now()))
					FkProductList=Replace(FkProductList,"{$ProductListPic$}",trim(Rs("Fk_Product_Pic")&" "))
					If Not IsNull(Rs("Fk_Product_PicBig")) Then
						FkProductList=Replace(FkProductList,"{$ProductListPicBig$}",Rs("Fk_Product_PicBig"))
					End If
					
					on error resume next
					'-------------------
					'扩展字段标签替换
					'time: 2013-12-07
					'add ：shark
					'-------------------
					dim FK_Product_SlidesImgs,FK_Product_summary,FK_Product_SlidesFirst,Fk_Product_ContentEx1,Fk_Product_ContentEx2
					FK_Product_SlidesImgs=Trim(Rs("FK_Product_SlidesImgs")&" ")
					FK_Product_summary=Trim(Rs("FK_Product_summary")&" ")
					FK_Product_SlidesFirst=Trim(Rs("FK_Product_SlidesFirst")&" ")
					Fk_Product_ContentEx1=Trim(Rs("Fk_Product_ContentEx1")&" ")
					Fk_Product_ContentEx2=Trim(Rs("Fk_Product_ContentEx2")&" ")
					
					if instr(FkProductList,"{$ProductListPicSlidesImgList$}")>0 then
						FkProductList=Replace(FkProductList,"{$ProductListPicSlidesImgList$}",FK_Product_SlidesImgs)
					end if
					dim TempImg,SlidesImgs
					TempImg=""
					if FK_Product_SlidesImgs<>"" then
		                SlidesImgs=Split(FK_Product_SlidesImgs,",")
						dim j
		                For j=0 to Ubound(SlidesImgs)
		                	TempImg=TempImg&"<li><img src="""&Trim(SlidesImgs(i))&""" /></li>"
		                Next
					end if
					FkProductList=Replace(FkProductList,"{$ProductListPicSlidesFirst$}",FK_Product_SlidesFirst)
					
					if instr(FkProductList,"{$ProductListPicSlidesImgs$}")>0 then
						FkProductList=Replace(FkProductList,"{$ProductListPicSlidesImgs$}",TempImg)
					end if
					if FK_Product_summary="" then
						FK_Product_summary=Left(RemoveHTML(Rs("Fk_Product_Content")),SiteMini)
					end if
					FkProductList=Replace(FkProductList,"{$ProductListPicSummary$}",FK_Product_summary)
					FkProductList=Replace(FkProductList,"{$ProductList_ContentEx1$}",Fk_Product_ContentEx1)
					FkProductList=Replace(FkProductList,"{$ProductList_ContentEx2$}",Fk_Product_ContentEx2)
			
			'-------------------------------------------------------------------------------------------
					
					
					Rs.MoveNext
					z=z+1
				Wend
			Else
				If TempArr(3)<>0 And TempArr(4)=1 Then
					PageSizes=TempArr(3)
				End If
				Rs.PageSize=PageSizes
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				Rs.AbsolutePage=PageNow
				PageAll=Rs.RecordCount
				i=1
				z=PageSizes*(PageNow-1)+1
				While (Not Rs.Eof) And i<PageSizes+1
					If Rs("Fk_Product_Url")<>"" Then
						ProductUrl=Rs("Fk_Product_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							ProductUrl="/html"&SiteDir&ProductUrl
						Else
							ProductUrl=SiteDir&sTemp&"?"&ProductUrl
						End If
					End If
					ProductTitle=Rs("Fk_Product_Title")
					If Len(ProductTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						ProductTitle=Left(ProductTitle,Clng(TempArr(7)))&"..."
					End If
					FkProductList=FkProductList&BCode
					If Rs("Fk_Product_Color")<>"" Then
						FkProductList=Replace(FkProductList,"{$ProductListTitle$}","<span style='color:"&Rs("Fk_Product_Color")&"'>"&ProductTitle&"</span>")
					Else
						FkProductList=Replace(FkProductList,"{$ProductListTitle$}",ProductTitle)
					End If
					If Rs("Fk_Product_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkProductList=Replace(FkProductList,"{$ProductList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkProductList=Replace(FkProductList,"{$ProductList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkProductList=Replace(FkProductList,"{$ListNo$}",z)
					FkProductList=Replace(FkProductList,"{$ListNo2$}",i)
					FkProductList=Replace(FkProductList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkProductList=Replace(FkProductList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkProductList,"{$ModuleListUrl$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),	Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					
					If InStr(FkProductList,"{$ModuleListContent$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If
					
					If InStr(FkProductList,"{$ModuleListContentNoHtml$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkProductList=Replace(FkProductList,"{$ProductListId$}",Rs("Fk_Product_Id"))
					FkProductList=Replace(FkProductList,"{$ProductListTitleAll$}",Rs("Fk_Product_Title"))
					If InStr(FkProductList,"{$ProductListContent$}")>0 Then
						FkProductList=Replace(FkProductList,"{$ProductListContent$}",Left(RemoveHTML(Rs("Fk_Product_Content")),SiteMini))
					End If
					FkProductList=Replace(FkProductList,"{$ProductListUrl$}",ProductUrl)
					FkProductList=Replace(FkProductList,"{$ProductListClick$}",Rs("Fk_Product_Click"))
					FkProductList=Replace(FkProductList,"{$ProductListTime$}",Rs("Fk_Product_Time"))
					FkProductList=Replace(FkProductList,"{$ProductListDate$}",FormatDateTime(Rs("Fk_Product_Time"),2))
					FkProductList=Replace(FkProductList,"{$ProductListYear$}",Year(Rs("Fk_Product_Time")))
					FkProductList=Replace(FkProductList,"{$ProductListMonth$}",Month(Rs("Fk_Product_Time")))
					FkProductList=Replace(FkProductList,"{$ProductListDay$}",Day(Rs("Fk_Product_Time")))
					FkProductList=Replace(FkProductList,"{$ProductListNew$}",DateDiff("d",Rs("Fk_Product_Time"),Now()))
					FkProductList=Replace(FkProductList,"{$ProductListPic$}",trim(Rs("Fk_Product_Pic")&" "))
					
					
					on error resume next
					'-------------------
					'扩展字段标签替换
					'addtime: 2013-12-07
					'addby ：shark
					'edittime:2014年12月17日
					'editby: shark
					'-------------------
					FK_Product_SlidesImgs=Trim(Rs("FK_Product_SlidesImgs")&" ")
					FK_Product_summary=Trim(Rs("FK_Product_summary")&" ")
					FK_Product_SlidesFirst=Trim(Rs("FK_Product_SlidesFirst")&" ")
					Fk_Product_ContentEx1=Trim(Rs("Fk_Product_ContentEx1")&" ")
					Fk_Product_ContentEx2=Trim(Rs("Fk_Product_ContentEx2")&" ")
					
					if instr(FkProductList,"{$ProductListPicSlidesImgList$}")>0 then
						FkProductList=Replace(FkProductList,"{$ProductListPicSlidesImgList$}",FK_Product_SlidesImgs)
					end if
					TempImg=""
					if FK_Product_SlidesImgs<>"" then
		                SlidesImgs=Split(FK_Product_SlidesImgs,",")
		                For j=0 to Ubound(SlidesImgs)
		                	TempImg=TempImg&"<li><img src="""&Trim(SlidesImgs(i))&""" /></li>"
		                Next
					end if
					FkProductList=Replace(FkProductList,"{$ProductListPicSlidesFirst$}",FK_Product_SlidesFirst)
					
					if instr(FkProductList,"{$ProductListPicSlidesImgs$}")>0 then
						FkProductList=Replace(FkProductList,"{$ProductListPicSlidesImgs$}",TempImg)
					end if
					if FK_Product_summary="" then
						FK_Product_summary=Left(RemoveHTML(Rs("Fk_Product_Content")),SiteMini)
					end if
					FkProductList=Replace(FkProductList,"{$ProductListPicSummary$}",FK_Product_summary)
					FkProductList=Replace(FkProductList,"{$ProductList_ContentEx1$}",Fk_Product_ContentEx1)
					FkProductList=Replace(FkProductList,"{$ProductList_ContentEx2$}",Fk_Product_ContentEx2)
			
			'-------------------------------------------------------------------------------------------
			
					If Not IsNull(Rs("Fk_Product_PicBig")) Then
						FkProductList=Replace(FkProductList,"{$ProductListPicBig$}",Rs("Fk_Product_PicBig"))
					End If
					Rs.MoveNext
					i=i+1
					z=z+1
				Wend
'				If PageNow>1 Then
'					PageFirst=SiteDir&"?"&CategoryDirName&"/Index.html"
'					PagePrev=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow-1)&".html"
'					If PageNow=2 Then
'						PagePrev=SiteDir&"?"&CategoryDirName&"/Index.html"
'					End If
'				Else
'					PageFirst="#"
'					PagePrev="#"
'				End If
'				If PageCounts>PageNow Then
'					PageNext=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow+1)&".html"
'					PageLast=SiteDir&"?"&CategoryDirName&"/Index_"&PageCounts&".html"
'				Else
'					PageNext="#"
'					PageLast="#"
'				End If
'				If SiteHtml=1 Then
'					PageFirst=Replace(PageFirst,"?","html/")
'					PagePrev=Replace(PagePrev,"?","html/")
'					PageNext=Replace(PageNext,"?","html/")
'					PageLast=Replace(PageLast,"?","html/")
'				Else
'					PageFirst=Replace(PageFirst,"?",sTemp&"?")
'					PagePrev=Replace(PagePrev,"?",sTemp&"?")
'					PageNext=Replace(PageNext,"?",sTemp&"?")
'					PageLast=Replace(PageLast,"?",sTemp&"?")
'				End If
			End If
		End If
		Rs.Close
	End Function

	'==============================
	'函 数 名：FkProductVideoList
	'作    用：下载列表标签操作
	'参    数：
	'==============================
	Private Function FkProductVideoList(BCode,BPar)
		Dim ArticleTitle,z
		Dim Rst
		Set Rst=Server.Createobject("Adodb.RecordSet")
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(3)>0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(3)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_Article] Where Fk_Article_Show=1"
		If TempArr(1)>0 Then
			Sqlstr=Sqlstr&" And Fk_Relation_Product_Id="&TempArr(1)
		End If
		If TempArr(5)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Article_Recommend Like '%%,2,%%' or Fk_Down_Ip='1')"
		End If
		If TempArr(6)>0 Then
			Sqlstr=Sqlstr&" And Fk_Article_Subject Like '%%,"&TempArr(6)&",%%'"
		End If
		If TempArr(4)=1 And SearchStr<>"" Then
			'Sqlstr=Sqlstr&" And Fk_Article_Title Like '%"&SearchStr&"%'"
			Sqlstr=Sqlstr&" And InStr(1,LCase(Fk_Article_Title),LCase('"&SearchStr&"'),0)<>0"
		End If
		Select Case TempArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Time Desc,Fk_Article_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Click Desc,Fk_Article_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Time Asc,Fk_Article_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc, Px desc, Fk_Article_Click Asc,Fk_Article_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If TempArr(4)=0 Then
				While Not Rs.Eof
					ArticleTitle=Rs("Fk_Article_Title")
					If Len(ArticleTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						ArticleTitle=Left(ArticleTitle,Clng(TempArr(7)))&"..."
					End If
					FkProductVideoList=FkProductVideoList&BCode
					If Rs("Fk_Article_Color")<>"" Then
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTitle$}","<span style='color:"&Rs("Fk_Article_Color")&"'>"&ArticleTitle&"</span>")
					Else
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTitle$}",ArticleTitle)
					End If
					If Rs("Fk_Article_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkProductVideoList=Replace(FkProductVideoList,"{$ArticleList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkProductVideoList=Replace(FkProductVideoList,"{$ListNo$}",z)
					
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListId$}",Rs("Fk_Article_Id"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTitleAll$}",Rs("Fk_Article_Title"))
					If InStr(FkProductVideoList,"{$ArticleListContent$}")>0 Then
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListContent$}",Left(RemoveHTML(Rs("Fk_Article_Content")),SiteMini))
					End If
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListClick$}",Rs("Fk_Article_Click"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTime$}",Rs("Fk_Article_Time"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListDate$}",FormatDateTime(Rs("Fk_Article_Time"),2))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListYear$}",Year(Rs("Fk_Article_Time")))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListMonth$}",Month(Rs("Fk_Article_Time")))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListDay$}",Day(Rs("Fk_Article_Time")))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListNew$}",DateDiff("d",Rs("Fk_Article_Time"),Now()))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListPic$}",trim(Rs("Fk_Article_Pic")&" "))
					Rs.MoveNext
					z=z+1
				Wend
			Else
				Rs.PageSize=PageSizes
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				Rs.AbsolutePage=PageNow
				PageAll=Rs.RecordCount
				i=1
				z=PageSizes*(PageNow-1)+1
				While (Not Rs.Eof) And i<PageSizes+1
					ArticleTitle=Rs("Fk_Article_Title")
					If Len(ArticleTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						ArticleTitle=Left(ArticleTitle,Clng(TempArr(7)))&"..."
					End If
					FkProductVideoList=FkProductVideoList&BCode
					If Rs("Fk_Article_Color")<>"" Then
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTitle$}","<span style='color:"&Rs("Fk_Article_Color")&"'>"&ArticleTitle&"</span>")
					Else
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTitle$}",ArticleTitle)
					End If
					If Rs("Fk_Article_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkProductVideoList=Replace(FkProductVideoList,"{$ArticleList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkProductVideoList=Replace(FkProductVideoList,"{$ListNo$}",z)
					FkProductVideoList=Replace(FkProductVideoList,"{$ListNo2$}",i)
					
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListId$}",Rs("Fk_Article_Id"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTitleAll$}",Rs("Fk_Article_Title"))
					If InStr(FkProductVideoList,"{$ArticleListContent$}")>0 Then
						FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListContent$}",Left(RemoveHTML(Rs("Fk_Article_Content")),SiteMini))
					End If
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListUrl$}",ArticleUrl)
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListClick$}",Rs("Fk_Article_Click"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListCount$}",Rs("Fk_Article_Count"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListTime$}",Rs("Fk_Article_Time"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListSystem$}",Rs("Fk_Article_System"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListLanguage$}",Rs("Fk_Article_Language"))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListDate$}",FormatDateTime(Rs("Fk_Article_Time"),2))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListYear$}",Year(Rs("Fk_Article_Time")))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListMonth$}",Month(Rs("Fk_Article_Time")))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListDay$}",Day(Rs("Fk_Article_Time")))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListNew$}",DateDiff("d",Rs("Fk_Article_Time"),Now()))
					FkProductVideoList=Replace(FkProductVideoList,"{$ArticleListPic$}",trim(Rs("Fk_Article_Pic")&" "))
					Rs.MoveNext
					i=i+1
					z=z+1
				Wend
			End If
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkProductDownList
	'作    用：产品关联资料列表标签操作
	'参    数：
	'==============================
	Private Function FkProductDownList(BCode,BPar)
		Dim DownUrl,DownTitle,z
		Dim Rst
		Set Rst=Server.Createobject("Adodb.RecordSet")
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(3)>0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(3)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Menu=" & TempArr(0)
		If TempArr(1)>0 Then
			Sqlstr=Sqlstr&" And Fk_Relation_Product_Id="&TempArr(1)
		End If
		If TempArr(5)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Down_Recommend Like '%%,2,%%' or Fk_Down_Ip='1')"
		End If
		If TempArr(6)>0 Then
			Sqlstr=Sqlstr&" And Fk_Down_Subject Like '%%,"&TempArr(6)&",%%'"
		End If
		If TempArr(4)=1 And SearchStr<>"" Then
			'Sqlstr=Sqlstr&" And Fk_Down_Title Like '%"&SearchStr&"%'"
			Sqlstr=Sqlstr&" And InStr(1,LCase(Fk_Down_Title),LCase('"&SearchStr&"'),0)<>0"
		End If
		Select Case TempArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Time Desc,Fk_Down_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Click Desc,Fk_Down_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Time Asc,Fk_Down_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Click Asc,Fk_Down_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If TempArr(4)=0 Then
				While Not Rs.Eof
					If Rs("Fk_Down_Url")<>"" Then
						DownUrl=Rs("Fk_Down_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							DownUrl="/html"&SiteDir&DownUrl
						Else
							DownUrl=SiteDir&sTemp&"?"&DownUrl
						End If
					End If
					DownTitle=Rs("Fk_Down_Title")
					If Len(DownTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						DownTitle=Left(DownTitle,Clng(TempArr(7)))&"..."
					End If
					FkProductDownList=FkProductDownList&BCode
					If Rs("Fk_Down_Color")<>"" Then
						FkProductDownList=Replace(FkProductDownList,"{$DownListTitle$}","<span style='color:"&Rs("Fk_Down_Color")&"'>"&DownTitle&"</span>")
					Else
						FkProductDownList=Replace(FkProductDownList,"{$DownListTitle$}",DownTitle)
					End If
					If Rs("Fk_Down_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkProductDownList=Replace(FkProductDownList,"{$DownList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkProductDownList=Replace(FkProductDownList,"{$DownList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkProductDownList=Replace(FkProductDownList,"{$ListNo$}",z)
					FkProductDownList=Replace(FkProductDownList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkProductDownList=Replace(FkProductDownList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkProductDownList,"{$ModuleListUrl$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					
					If InStr(FkProductDownList,"{$ModuleListContent$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If
					
					If InStr(FkProductDownList,"{$ModuleListContentNoHtml$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkProductDownList=Replace(FkProductDownList,"{$DownListId$}",Rs("Fk_Down_Id"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListTitleAll$}",Rs("Fk_Down_Title"))
					If InStr(FkProductDownList,"{$DownListContent$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$DownListContent$}",Left(RemoveHTML(Rs("Fk_Down_Content")),SiteMini))
					End If
					FkProductDownList=Replace(FkProductDownList,"{$DownListUrl$}",DownUrl)
					FkProductDownList=Replace(FkProductDownList,"{$DownListClick$}",Rs("Fk_Down_Click"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListCount$}",Rs("Fk_Down_Count"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListTime$}",Rs("Fk_Down_Time"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListSystem$}",Rs("Fk_Down_System"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListLanguage$}",Rs("Fk_Down_Language"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListFile$}",SiteDir&"File.asp?Id="&Rs("Fk_Down_Id"))'Rs("Fk_Down_File"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListDate$}",FormatDateTime(Rs("Fk_Down_Time"),2))
					FkProductDownList=Replace(FkProductDownList,"{$DownListYear$}",Year(Rs("Fk_Down_Time")))
					FkProductDownList=Replace(FkProductDownList,"{$DownListMonth$}",Month(Rs("Fk_Down_Time")))
					FkProductDownList=Replace(FkProductDownList,"{$DownListDay$}",Day(Rs("Fk_Down_Time")))
					FkProductDownList=Replace(FkProductDownList,"{$DownListNew$}",DateDiff("d",Rs("Fk_Down_Time"),Now()))
					FkProductDownList=Replace(FkProductDownList,"{$DownListPic$}",trim(Rs("Fk_Down_Pic")&" "))
					If Not IsNull(Rs("Fk_Down_PicBig")) Then
						FkProductDownList=Replace(FkProductDownList,"{$DownListPicBig$}",Rs("Fk_Down_PicBig"))
					End If
					Rs.MoveNext
					z=z+1
				Wend
			Else
				Rs.PageSize=PageSizes
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				Rs.AbsolutePage=PageNow
				PageAll=Rs.RecordCount
				i=1
				z=PageSizes*(PageNow-1)+1
				While (Not Rs.Eof) And i<PageSizes+1
					If Rs("Fk_Down_Url")<>"" Then
						DownUrl=Rs("Fk_Down_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							DownUrl="/html"&SiteDir&DownUrl
						Else
							DownUrl=SiteDir&sTemp&"?"&DownUrl
						End If
					End If
					DownTitle=Rs("Fk_Down_Title")
					If Len(DownTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						DownTitle=Left(DownTitle,Clng(TempArr(7)))&"..."
					End If
					FkProductDownList=FkProductDownList&BCode
					If Rs("Fk_Down_Color")<>"" Then
						FkProductDownList=Replace(FkProductDownList,"{$DownListTitle$}","<span style='color:"&Rs("Fk_Down_Color")&"'>"&DownTitle&"</span>")
					Else
						FkProductDownList=Replace(FkProductDownList,"{$DownListTitle$}",DownTitle)
					End If
					If Rs("Fk_Down_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkProductDownList=Replace(FkProductDownList,"{$DownList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkProductDownList=Replace(FkProductDownList,"{$DownList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkProductDownList=Replace(FkProductDownList,"{$ListNo$}",z)
					FkProductDownList=Replace(FkProductDownList,"{$ListNo2$}",i)
					FkProductDownList=Replace(FkProductDownList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkProductDownList=Replace(FkProductDownList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkProductDownList,"{$ModuleListUrl$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					
					If InStr(FkProductDownList,"{$ModuleListContent$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If

					If InStr(FkProductDownList,"{$ModuleListContentNoHtml$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkProductDownList=Replace(FkProductDownList,"{$DownListId$}",Rs("Fk_Down_Id"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListTitleAll$}",Rs("Fk_Down_Title"))
					If InStr(FkProductDownList,"{$DownListContent$}")>0 Then
						FkProductDownList=Replace(FkProductDownList,"{$DownListContent$}",Left(RemoveHTML(Rs("Fk_Down_Content")),SiteMini))
					End If
					FkProductDownList=Replace(FkProductDownList,"{$DownListUrl$}",DownUrl)
					FkProductDownList=Replace(FkProductDownList,"{$DownListClick$}",Rs("Fk_Down_Click"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListCount$}",Rs("Fk_Down_Count"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListTime$}",Rs("Fk_Down_Time"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListSystem$}",Rs("Fk_Down_System"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListLanguage$}",Rs("Fk_Down_Language"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListFile$}",SiteDir&"File.asp?Id="&Rs("Fk_Down_Id"))'Rs("Fk_Down_File"))
					FkProductDownList=Replace(FkProductDownList,"{$DownListDate$}",FormatDateTime(Rs("Fk_Down_Time"),2))
					FkProductDownList=Replace(FkProductDownList,"{$DownListYear$}",Year(Rs("Fk_Down_Time")))
					FkProductDownList=Replace(FkProductDownList,"{$DownListMonth$}",Month(Rs("Fk_Down_Time")))
					FkProductDownList=Replace(FkProductDownList,"{$DownListDay$}",Day(Rs("Fk_Down_Time")))
					FkProductDownList=Replace(FkProductDownList,"{$DownListNew$}",DateDiff("d",Rs("Fk_Down_Time"),Now()))
					FkProductDownList=Replace(FkProductDownList,"{$DownListPic$}",trim(Rs("Fk_Down_Pic")&" "))
					If Not IsNull(Rs("Fk_Down_PicBig")) Then
						FkProductDownList=Replace(FkProductDownList,"{$DownListPicBig$}",Rs("Fk_Down_PicBig"))
					End If
					Rs.MoveNext
					i=i+1
					z=z+1
				Wend
				If PageNow>1 Then
					PageFirst=SiteDir&"?"&CategoryDirName&"/Index.html"
					PagePrev=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow-1)&".html"
					If PageNow=2 Then
						PagePrev=SiteDir&"?"&CategoryDirName&"/Index.html"
					End If
				Else
					PageFirst="#"
					PagePrev="#"
				End If
				If PageCounts>PageNow Then
					PageNext=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow+1)&".html"
					PageLast=SiteDir&"?"&CategoryDirName&"/Index_"&PageCounts&".html"
				Else
					PageNext="#"
					PageLast="#"
				End If
				If SiteHtml=1 and sitetemplate<>"wap" Then
					PageFirst=Replace(PageFirst,"?","html/")
					PagePrev=Replace(PagePrev,"?","html/")
					PageNext=Replace(PageNext,"?","html/")
					PageLast=Replace(PageLast,"?","html/")
				Else
					PageFirst=Replace(PageFirst,"?",sTemp&"?")
					PagePrev=Replace(PagePrev,"?",sTemp&"?")
					PageNext=Replace(PageNext,"?",sTemp&"?")
					PageLast=Replace(PageLast,"?",sTemp&"?")
				End If
			End If
		End If
		Rs.Close
	End Function

	'==============================
	'函 数 名：FkDownList
	'作    用：下载列表标签操作
	'参    数：
	'==============================
	Private Function FkDownList(BCode,BPar)
		Dim DownUrl,DownTitle,z
		Dim Rst
		Set Rst=Server.Createobject("Adodb.RecordSet")
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(3)>0 And TempArr(4)=0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(3)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_DownList] Where Fk_Down_Show=1 And Fk_Module_Menu=" & TempArr(0)
		If TempArr(1)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Down_Module="&TempArr(1)&" Or Fk_Module_LevelList Like '%%,"&TempArr(1)&",%%')"
		End If
		If TempArr(5)>0 Then
			Sqlstr=Sqlstr&" And (Fk_Down_Recommend Like '%%,2,%%' or Fk_Down_Ip='1')"
		End If
		If TempArr(6)>0 Then
			Sqlstr=Sqlstr&" And Fk_Down_Subject Like '%%,"&TempArr(6)&",%%'"
		End If
		If TempArr(4)=1 And SearchStr<>"" Then
			'Sqlstr=Sqlstr&" And Fk_Down_Title Like '%"&SearchStr&"%'"
			Sqlstr=Sqlstr&" And InStr(1,LCase(Fk_Down_Title),LCase('"&SearchStr&"'),0)<>0"
		End If
		Select Case TempArr(2)
			Case 0
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Id Desc"
			Case 1
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Time Desc,Fk_Down_Id Desc"
			Case 2
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Click Desc,Fk_Down_Id Desc"
			Case 3
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Id Asc"
			Case 4
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Time Asc,Fk_Down_Id Desc"
			Case 5
				Sqlstr=Sqlstr&" Order By Fk_Down_Ip desc, Px desc, Fk_Down_Click Asc,Fk_Down_Id Desc"
		End Select
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If TempArr(4)=0 Then
				While Not Rs.Eof
					If Rs("Fk_Down_Url")<>"" Then
						DownUrl=Rs("Fk_Down_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							DownUrl="/html"&SiteDir&DownUrl
						Else
							DownUrl=SiteDir&sTemp&"?"&DownUrl
						End If
					End If
					DownTitle=Rs("Fk_Down_Title")
					If Len(DownTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						DownTitle=Left(DownTitle,Clng(TempArr(7)))&"..."
					End If
					FkDownList=FkDownList&BCode
					If Rs("Fk_Down_Color")<>"" Then
						FkDownList=Replace(FkDownList,"{$DownListTitle$}","<span style='color:"&Rs("Fk_Down_Color")&"'>"&DownTitle&"</span>")
					Else
						FkDownList=Replace(FkDownList,"{$DownListTitle$}",DownTitle)
					End If
					If Rs("Fk_Down_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkDownList=Replace(FkDownList,"{$DownList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkDownList=Replace(FkDownList,"{$DownList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkDownList=Replace(FkDownList,"{$ListNo$}",z)
					FkDownList=Replace(FkDownList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkDownList=Replace(FkDownList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkDownList,"{$ModuleListUrl$}")>0 Then
						FkDownList=Replace(FkDownList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					
					If InStr(FkDownList,"{$ModuleListContent$}")>0 Then
						FkDownList=Replace(FkDownList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If
					
					If InStr(FkDownList,"{$ModuleListContentNoHtml$}")>0 Then
						FkDownList=Replace(FkDownList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkDownList=Replace(FkDownList,"{$DownListId$}",Rs("Fk_Down_Id"))
					FkDownList=Replace(FkDownList,"{$DownListTitleAll$}",Rs("Fk_Down_Title"))
					If InStr(FkDownList,"{$DownListContent$}")>0 Then
						FkDownList=Replace(FkDownList,"{$DownListContent$}",Left(RemoveHTML(Rs("Fk_Down_Content")),SiteMini))
					End If
					FkDownList=Replace(FkDownList,"{$DownListUrl$}",DownUrl)
					FkDownList=Replace(FkDownList,"{$DownListClick$}",Rs("Fk_Down_Click"))
					FkDownList=Replace(FkDownList,"{$DownListCount$}",Rs("Fk_Down_Count"))
					FkDownList=Replace(FkDownList,"{$DownListTime$}",Rs("Fk_Down_Time"))
					FkDownList=Replace(FkDownList,"{$DownListSystem$}",Rs("Fk_Down_System"))
					FkDownList=Replace(FkDownList,"{$DownListLanguage$}",Rs("Fk_Down_Language"))
					FkDownList=Replace(FkDownList,"{$DownListFile$}",SiteDir&"File.asp?Id="&Rs("Fk_Down_Id"))'Rs("Fk_Down_File"))
					FkDownList=Replace(FkDownList,"{$DownListDate$}",FormatDateTime(Rs("Fk_Down_Time"),2))
					FkDownList=Replace(FkDownList,"{$DownListYear$}",Year(Rs("Fk_Down_Time")))
					FkDownList=Replace(FkDownList,"{$DownListMonth$}",Month(Rs("Fk_Down_Time")))
					FkDownList=Replace(FkDownList,"{$DownListDay$}",Day(Rs("Fk_Down_Time")))
					FkDownList=Replace(FkDownList,"{$DownListNew$}",DateDiff("d",Rs("Fk_Down_Time"),Now()))
					FkDownList=Replace(FkDownList,"{$DownListPic$}",trim(Rs("Fk_Down_Pic")&" "))
					If Not IsNull(Rs("Fk_Down_PicBig")) Then
						FkDownList=Replace(FkDownList,"{$DownListPicBig$}",Rs("Fk_Down_PicBig"))
					End If
					Rs.MoveNext
					z=z+1
				Wend
			Else
				Rs.PageSize=PageSizes
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				Rs.AbsolutePage=PageNow
				PageAll=Rs.RecordCount
				i=1
				z=PageSizes*(PageNow-1)+1
				While (Not Rs.Eof) And i<PageSizes+1
					If Rs("Fk_Down_Url")<>"" Then
						DownUrl=Rs("Fk_Down_Url")
					Else
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
						If SiteHtml=1 and sitetemplate<>"wap" Then
							DownUrl="/html"&SiteDir&DownUrl
						Else
							DownUrl=SiteDir&sTemp&"?"&DownUrl
						End If
					End If
					DownTitle=Rs("Fk_Down_Title")
					If Len(DownTitle)>Clng(TempArr(7)) And Clng(TempArr(7))>0 Then
						DownTitle=Left(DownTitle,Clng(TempArr(7)))&"..."
					End If
					FkDownList=FkDownList&BCode
					If Rs("Fk_Down_Color")<>"" Then
						FkDownList=Replace(FkDownList,"{$DownListTitle$}","<span style='color:"&Rs("Fk_Down_Color")&"'>"&DownTitle&"</span>")
					Else
						FkDownList=Replace(FkDownList,"{$DownListTitle$}",DownTitle)
					End If
					If Rs("Fk_Down_Field")<>"" Then
						TemplateTempArr=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
						For Each TemplateTemp In TemplateTempArr
							FkDownList=Replace(FkDownList,"{$DownList_"&Split(TemplateTemp,"|-Fangka_Field-|")(0)&"$}",Split(TemplateTemp,"|-Fangka_Field-|")(1))
						Next
					End If
					Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
					Rst.Open Sqlstr,Conn,1,1
					While Not Rst.Eof
						FkDownList=Replace(FkDownList,"{$DownList_"&Rst("Fk_Field_Tag")&"$}","")
						Rst.MoveNext
					Wend
					Rst.Close
					FkDownList=Replace(FkDownList,"{$ListNo$}",z)
					FkDownList=Replace(FkDownList,"{$ListNo2$}",i)
					FkDownList=Replace(FkDownList,"{$ModuleListId$}",Rs("Fk_Module_Id"))
					FkDownList=Replace(FkDownList,"{$ModuleListName$}",Rs("Fk_Module_Name"))
					If InStr(FkDownList,"{$ModuleListUrl$}")>0 Then
						FkDownList=Replace(FkDownList,"{$ModuleListUrl$}",GetGoUrl(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"),Rs("Fk_Module_Dir"),Rs("Fk_Module_FileName")))
					End If
					
					If InStr(FkDownList,"{$ModuleListContent$}")>0 Then
						FkDownList=Replace(FkDownList,"{$ModuleListContent$}",GetModuleContent(Rs("Fk_Module_Id")))
					End If

					If InStr(FkDownList,"{$ModuleListContentNoHtml$}")>0 Then
						FkDownList=Replace(FkDownList,"{$ModuleListContentNoHtml$}",RemoveHTML(GetModuleContent(Rs("Fk_Module_Id"))))
					End If
					
					FkDownList=Replace(FkDownList,"{$DownListId$}",Rs("Fk_Down_Id"))
					FkDownList=Replace(FkDownList,"{$DownListTitleAll$}",Rs("Fk_Down_Title"))
					If InStr(FkDownList,"{$DownListContent$}")>0 Then
						FkDownList=Replace(FkDownList,"{$DownListContent$}",Left(RemoveHTML(Rs("Fk_Down_Content")),SiteMini))
					End If
					FkDownList=Replace(FkDownList,"{$DownListUrl$}",DownUrl)
					FkDownList=Replace(FkDownList,"{$DownListClick$}",Rs("Fk_Down_Click"))
					FkDownList=Replace(FkDownList,"{$DownListCount$}",Rs("Fk_Down_Count"))
					FkDownList=Replace(FkDownList,"{$DownListTime$}",Rs("Fk_Down_Time"))
					FkDownList=Replace(FkDownList,"{$DownListSystem$}",Rs("Fk_Down_System"))
					FkDownList=Replace(FkDownList,"{$DownListLanguage$}",Rs("Fk_Down_Language"))
					FkDownList=Replace(FkDownList,"{$DownListFile$}",SiteDir&"File.asp?Id="&Rs("Fk_Down_Id"))'Rs("Fk_Down_File"))
					FkDownList=Replace(FkDownList,"{$DownListDate$}",FormatDateTime(Rs("Fk_Down_Time"),2))
					FkDownList=Replace(FkDownList,"{$DownListYear$}",Year(Rs("Fk_Down_Time")))
					FkDownList=Replace(FkDownList,"{$DownListMonth$}",Month(Rs("Fk_Down_Time")))
					FkDownList=Replace(FkDownList,"{$DownListDay$}",Day(Rs("Fk_Down_Time")))
					FkDownList=Replace(FkDownList,"{$DownListNew$}",DateDiff("d",Rs("Fk_Down_Time"),Now()))
					FkDownList=Replace(FkDownList,"{$DownListPic$}",trim(Rs("Fk_Down_Pic")&" "))
					If Not IsNull(Rs("Fk_Down_PicBig")) Then
						FkDownList=Replace(FkDownList,"{$DownListPicBig$}",Rs("Fk_Down_PicBig"))
					End If
					Rs.MoveNext
					i=i+1
					z=z+1
				Wend
				If PageNow>1 Then
					PageFirst=SiteDir&"?"&CategoryDirName&"/Index.html"
					PagePrev=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow-1)&".html"
					If PageNow=2 Then
						PagePrev=SiteDir&"?"&CategoryDirName&"/Index.html"
					End If
				Else
					PageFirst="#"
					PagePrev="#"
				End If
				If PageCounts>PageNow Then
					PageNext=SiteDir&"?"&CategoryDirName&"/Index_"&(PageNow+1)&".html"
					PageLast=SiteDir&"?"&CategoryDirName&"/Index_"&PageCounts&".html"
				Else
					PageNext="#"
					PageLast="#"
				End If
				If SiteHtml=1 and sitetemplate<>"wap" Then
					PageFirst=Replace(PageFirst,"?","html/")
					PagePrev=Replace(PagePrev,"?","html/")
					PageNext=Replace(PageNext,"?","html/")
					PageLast=Replace(PageLast,"?","html/")
				Else
					PageFirst=Replace(PageFirst,"?",sTemp&"?")
					PagePrev=Replace(PagePrev,"?",sTemp&"?")
					PageNext=Replace(PageNext,"?",sTemp&"?")
					PageLast=Replace(PageLast,"?",sTemp&"?")
				End If
			End If
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkFriendsList
	'作    用：友情链接列表标签操作
	'参    数：
	'==============================
	Private Function FkFriendsList(BCode,BPar)
		Dim z
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(2)>0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(2)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_Friends] Where 1=1"
		If TempArr(0)>0 Then
			Sqlstr=Sqlstr&" And Fk_Friends_FriendsType="&TempArr(0)&""
		End If
		If TempArr(1)=1 Then
			Sqlstr=Sqlstr&" And Fk_Friends_ShowType=1"
		Else
			Sqlstr=Sqlstr&" And Fk_Friends_ShowType=2"
		End If
		Sqlstr=Sqlstr&" Order by Fk_Friends_Id Asc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			While Not Rs.Eof
				FkFriendsList=FkFriendsList&BCode
				FkFriendsList=Replace(FkFriendsList,"{$ListNo$}",z)
				FkFriendsList=Replace(FkFriendsList,"{$FriendsName$}",Rs("Fk_Friends_Name"))
				FkFriendsList=Replace(FkFriendsList,"{$FriendsUrl$}",Rs("Fk_Friends_Url"))
				FkFriendsList=Replace(FkFriendsList,"{$FriendsAbout$}",Rs("Fk_Friends_About"))
				If Rs("Fk_Friends_Logo")<>"" Then
					FkFriendsList=Replace(FkFriendsList,"{$FriendsLogo$}",Rs("Fk_Friends_Logo"))
				End If
				Rs.MoveNext
				z=z+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkJobList
	'作    用：招聘列表标签操作
	'参    数：
	'==============================
	Private Function FkJobList(BCode,BPar)
		Dim z
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(0)>0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(0)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_Job] Where 1=1"
		If TempArr(1)=1 Then
			Sqlstr=Sqlstr&" And DateAdd('d',Fk_Job_Date,Fk_Job_Time)<=#"&Now()&"#"
		End If
		If TempArr(1)=2 Then
			Sqlstr=Sqlstr&" And DateAdd('d',Fk_Job_Date,Fk_Job_Time)>#"&Now()&"#"
		End If
		Sqlstr=Sqlstr&" Order By Fk_Job_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			While Not Rs.Eof
				FkJobList=FkJobList&BCode
				FkJobList=Replace(FkJobList,"{$ListNo$}",z)
				FkJobList=Replace(FkJobList,"{$JobName$}",Rs("Fk_Job_Name"))
				FkJobList=Replace(FkJobList,"{$JobCount$}",Rs("Fk_Job_Count"))
				FkJobList=Replace(FkJobList,"{$JobAbout$}",Rs("Fk_Job_About"))
				FkJobList=Replace(FkJobList,"{$JobArea$}",Rs("Fk_Job_Area"))
				If Rs("Fk_Job_Date")=0 Then
					FkJobList=Replace(FkJobList,"{$JobDate$}",arrTips(6))
				Else
					FkJobList=Replace(FkJobList,"{$JobDate$}",Rs("Fk_Job_Date")&arrTips(7))
				End If
				FkJobList=Replace(FkJobList,"{$JobTime$}",Rs("Fk_Job_Time"))
				Rs.MoveNext
				z=z+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkSubjectList
	'作    用：专题列表标签操作
	'参    数：
	'==============================
	Private Function FkSubjectList(BCode,BPar)
		Dim SubjectUrl,z
		z=1
		TempArr=Split(BPar,"/")
		Sqlstr="Select"
		If TempArr(0)>0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(0)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_Subject] Where 1=1"
		Sqlstr=Sqlstr&" Order By Fk_Subject_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			While Not Rs.Eof
				SubjectUrl="Subject.asp?Id=" & Rs("Fk_Subject_Id")
				FkSubjectList=FkSubjectList&BCode
				FkSubjectList=Replace(FkSubjectList,"{$ListNo$}",z)
				FkSubjectList=Replace(FkSubjectList,"{$SubjectListName$}",Rs("Fk_Subject_Name"))
				FkSubjectList=Replace(FkSubjectList,"{$SubjectListPic$}",Rs("Fk_Subject_Pic"))
				FkSubjectList=Replace(FkSubjectList,"{$SubjectListUrl$}",SubjectUrl)
				Rs.MoveNext
				z=z+1
			Wend
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：FkGBookList
	'作    用：留言列表标签操作
	'参    数：
	'==============================
	Private Function FkGBookList(BCode,BPar)
		Dim z
		z=1
		TempArr=Split(BPar,"/")
		If TempArr(0)>0 Then
			Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=4 And Fk_Module_Id=" & TempArr(0)
			Rs.Open Sqlstr,Conn,1,1
			If Not Rs.Eof Then
				If Rs("Fk_Module_FileName")="" Then
					CategoryDirName="GBook"&Rs("Fk_Module_Id")
				Else
					CategoryDirName=Rs("Fk_Module_FileName")
				End If
			End If
			Rs.Close
		End If
		Sqlstr="Select"
		If TempArr(1)>0 And TempArr(3)=0 Then
			Sqlstr=Sqlstr&" Top "&TempArr(1)&""
		End If
		Sqlstr=Sqlstr&" * From [Fk_GBook] Where Fk_GBook_Module=" & TempArr(0)
		If TempArr(2)=1 Then
			Sqlstr=Sqlstr&" And Fk_GBook_ReContent<>''"
		End If
		Sqlstr=Sqlstr&" Order By Fk_GBook_Id Desc"
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			If TempArr(3)=0 Then
				While Not Rs.Eof
					FkGBookList=FkGBookList&BCode
					FkGBookList=Replace(FkGBookList,"{$ListNo$}",z)
					FkGBookList=Replace(FkGBookList,"{$GBookListTitle$}",Rs("Fk_GBook_Title"))
					FkGBookList=Replace(FkGBookList,"{$GBookListName$}",Rs("Fk_GBook_Name"))
					FkGBookList=Replace(FkGBookList,"{$GBookListContent$}",Rs("Fk_GBook_Content"))
					FkGBookList=Replace(FkGBookList,"{$GBookListTime$}",Rs("Fk_GBook_Time"))
					If Rs("Fk_GBook_ReContent")<>"" Then
						FkGBookList=Replace(FkGBookList,"{$GBookListReContent$}",Rs("Fk_GBook_ReContent"))
						FkGBookList=Replace(FkGBookList,"{$GBookListReTime$}",Rs("Fk_GBook_ReTime"))
					Else
						FkGBookList=Replace(FkGBookList,"{$GBookListReContent$}",arrTips(8))
						FkGBookList=Replace(FkGBookList,"{$GBookListReTime$}","")
					End If
					Rs.MoveNext
					z=z+1
				Wend
			Else
				Rs.PageSize=PageSizes
				If PageNow>Rs.PageCount Or PageNow<=0 Then
					PageNow=1
				End If
				PageCounts=Rs.PageCount
				Rs.AbsolutePage=PageNow
				PageAll=Rs.RecordCount
				i=1
				z=PageSizes*(PageNow-1)+1
				While (Not Rs.Eof) And i<PageSizes+1
					FkGBookList=FkGBookList&BCode
					FkGBookList=Replace(FkGBookList,"{$ListNo$}",z)
					FkGBookList=Replace(FkGBookList,"{$ListNo2$}",i)
					FkGBookList=Replace(FkGBookList,"{$GBookListTitle$}",Rs("Fk_GBook_Title"))
					FkGBookList=Replace(FkGBookList,"{$GBookListName$}",Rs("Fk_GBook_Name"))
					FkGBookList=Replace(FkGBookList,"{$GBookListContent$}",Rs("Fk_GBook_Content"))
					FkGBookList=Replace(FkGBookList,"{$GBookListTime$}",Rs("Fk_GBook_Time"))
					If Rs("Fk_GBook_ReContent")<>"" Then
						FkGBookList=Replace(FkGBookList,"{$GBookListReContent$}",Rs("Fk_GBook_ReContent"))
						FkGBookList=Replace(FkGBookList,"{$GBookListReTime$}",Rs("Fk_GBook_ReTime"))
					Else
						FkGBookList=Replace(FkGBookList,"{$GBookListReContent$}",arrTips(8))
						FkGBookList=Replace(FkGBookList,"{$GBookListReTime$}","")
					End If
					Rs.MoveNext
					i=i+1
					z=z+1
				Wend
				If PageNow>1 Then
					PageFirst=SiteDir&"?"&CategoryDirName&".html"
					PagePrev=SiteDir&"?"&CategoryDirName&"__"&(PageNow-1)&".html"
					If PageNow=2 Then
						PagePrev=SiteDir&"?"&CategoryDirName&".html"
					End If
				Else
					PageFirst="#"
					PagePrev="#"
				End If
				If PageCounts>PageNow Then
					PageNext=SiteDir&"?"&CategoryDirName&"__"&(PageNow+1)&".html"
					PageLast=SiteDir&"?"&CategoryDirName&"__"&PageCounts&".html"
				Else
					PageNext="#"
					PageLast="#"
				End If
				If SiteHtml=1 and sitetemplate<>"wap" Then
					PageFirst=Replace(PageFirst,"?","html/")
					PagePrev=Replace(PagePrev,"?","html/")
					PageNext=Replace(PageNext,"?","html/")
					PageLast=Replace(PageLast,"?","html/")
				Else
					PageFirst=Replace(PageFirst,"?",sTemp&"?")
					PagePrev=Replace(PagePrev,"?",sTemp&"?")
					PageNext=Replace(PageNext,"?",sTemp&"?")
					PageLast=Replace(PageLast,"?",sTemp&"?")
				End If
			End If
		End If
		Rs.Close
	End Function
	
	'==============================
	'函 数 名：GetGoUrl
	'作    用：获取内容操作链接
	'参    数：模块类型ModuleType，模块ID ModuleId
	'==============================
	Public Function GetGoUrl(ModuleType,ModuleId,ModuleDir,ModuleFileName)
		If SiteHtml=0 or SiteTemplate="wap" Then
			Select Case ModuleType
				Case 0
					GetGoUrl=SiteDir&sTemp&"?Page"&ModuleId&".html"
				Case 1
					GetGoUrl=SiteDir&sTemp&"?Article"&ModuleId&"/Index.html"
				Case 2
					GetGoUrl=SiteDir&sTemp&"?Product"&ModuleId&"/Index.html"
				Case 3
					GetGoUrl=SiteDir&sTemp&"?Info"&ModuleId&".html"
				Case 4
					GetGoUrl=SiteDir&sTemp&"?GBook"&ModuleId&".html"
				Case 6
					GetGoUrl=SiteDir&sTemp&"?Job"&ModuleId&".html"
				Case 7
					GetGoUrl=SiteDir&sTemp&"?Down"&ModuleId&"/Index.html"
				Case Else
					GetGoUrl="#"
			End Select
			If (ModuleType=1 Or ModuleType=2 Or ModuleType=7) And ModuleDir<>"" Then
					GetGoUrl=SiteDir&sTemp&"?"&ModuleDir&"/Index.html"
			End If
			If (ModuleType=0 Or ModuleType=3 Or ModuleType=4) And ModuleFileName<>"" Then
					GetGoUrl=SiteDir&sTemp&"?"&ModuleFileName&".html"
			End If
		ElseIf SiteHtml=1 Then
			Select Case ModuleType
				Case 0
					GetGoUrl=SiteDir&"html/Page"&ModuleId&".html"
				Case 1
					GetGoUrl=SiteDir&"html/Article"&ModuleId&"/index.html"
				Case 2
					GetGoUrl=SiteDir&"html/Product"&ModuleId&"/index.html"
				Case 3
					GetGoUrl=SiteDir&"html/Info"&ModuleId&".html"
				Case 4
					GetGoUrl=SiteDir&"html/GBook"&ModuleId&".html"
				Case 6
					GetGoUrl=SiteDir&"html/Job"&ModuleId&".html"
				Case 7
					GetGoUrl=SiteDir&"html/Down"&ModuleId&"/index.html"
				Case Else
					GetGoUrl="#"
			End Select
			If (ModuleType=1 Or ModuleType=2 Or ModuleType=7) And ModuleDir<>"" Then
					GetGoUrl=SiteDir&"html/"&ModuleDir&"/index.html"
			End If
			If (ModuleType=0 Or ModuleType=3 Or ModuleType=4) And ModuleFileName<>"" Then
					GetGoUrl=SiteDir&"html/"&ModuleFileName&".html"
			End If
		End If
	End Function
		
	'==============================
	'函 数 名：ShowPageCode
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPageCode(PageUrl,Nows,AllCount,Sizes,AllPage)
		PageUrl=replace(PageUrl,"../","/")
		If Nows>1 Then
			ShowPageCode="<a href="""&Replace(PageUrl,"{Pages}",1)&""" title="""&PageArr(0)&""">"&PageArr(0)&"</a>"
			ShowPageCode=ShowPageCode&"&nbsp;"
			ShowPageCode=ShowPageCode&"<a href="""&Replace(PageUrl,"{Pages}",Nows-1)&""" title="""&PageArr(1)&""">"&PageArr(1)&"</a>"
			If SiteHtml=1 And SearchStr="" and sitetemplate<>"wap" Then
				ShowPageCode=Replace(ShowPageCode,"__1.",".")
				ShowPageCode=Replace(ShowPageCode,"_1.",".")
			End If
		Else
			ShowPageCode=ShowPageCode&""&PageArr(0)&""
			ShowPageCode=ShowPageCode&"&nbsp;"
			ShowPageCode=ShowPageCode&""&PageArr(1)&""
		End If
		ShowPageCode=ShowPageCode&"&nbsp;"
		If AllPage>Nows Then
			ShowPageCode=ShowPageCode&"<a href="""&Replace(PageUrl,"{Pages}",Nows+1)&""" title="""&PageArr(2)&""">"&PageArr(2)&"</a>"
			ShowPageCode=ShowPageCode&"&nbsp;"
			ShowPageCode=ShowPageCode&"<a href="""&Replace(PageUrl,"{Pages}",AllPage)&""" title="""&PageArr(3)&""">"&PageArr(3)&"</a>"
		Else
			ShowPageCode=ShowPageCode&""&PageArr(2)&""
			ShowPageCode=ShowPageCode&"&nbsp;"
			ShowPageCode=ShowPageCode&""&PageArr(3)&""
		End If
		TempArr=Split(PageUrl,"{Pages}")
		ShowPageCode=ShowPageCode&"&nbsp;"&Sizes&""&PageArr(4)&"&nbsp;"&PageArr(5)&""&AllPage&""&PageArr(6)&""&AllCount&""&PageArr(7)&"&nbsp;"&PageArr(8)&""&Nows&""&PageArr(9)&"&nbsp;"
		ShowPageCode=ShowPageCode&"<select name=""Change_Page"" id=""Change_Page"" onChange=""window.location.href=this.options[this.selectedIndex].value"">"
		For i=1 To AllPage
			If i=1 Then
				If i=Nows Then
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"_{Pages}","")&""" selected=""selected"">"&PageArr(10)&""&i&""&PageArr(11)&"</option>"
				Else
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"_{Pages}","")&""">"&PageArr(10)&""&i&""&PageArr(11)&"</option>"
				End If
			Else
				If i=Nows Then
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"{Pages}",i)&""" selected=""selected"">"&PageArr(10)&""&i&""&PageArr(11)&"</option>"
				Else
					ShowPageCode=ShowPageCode&"<option value="""&Replace(PageUrl,"{Pages}",i)&""">"&PageArr(10)&""&i&""&PageArr(11)&"</option>"
				End If
			End If
		Next
      	ShowPageCode=ShowPageCode&"</select>"
	End Function
	
	'==============================
	'函 数 名：RemoveHTML
	'作    用：过滤HTML
	'参    数：
	'==============================
	Private Function RemoveHTML(strHTML)
		Dim objRegExp, Match, Matches 
		Set objRegExp = New Regexp 
		objRegExp.IgnoreCase = True 
		objRegExp.Global = True 
		'取闭合的<> 
		objRegExp.Pattern = "<.+?>" 
		'进行匹配 
		Set Matches = objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		'取特殊字符
		objRegExp.Pattern = "\&.+?;" 
		'进行匹配 
		Set Matches = objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		RemoveHTML=strHTML 
		Set objRegExp = Nothing 
	End Function

'========================模板引擎区===========================
	'==============================
	'函 数 名：TemplateDo
	'作    用：获取优先处理函数
	'参    数：
	'==============================
	Public Function TemplateDo(TemplateCode)
		Dim ForI,IfI
		ForI=Instr(TemplateCode,"{$For")
		IfI=Instr(TemplateCode,"{$If")
		If ForI=0 And IfI=0 Then
			TemplateDo=TemplateCode
			Exit Function
		End If
		If ForI>0 And IfI>0 Then
			If ForI<IfI Then
				TemplateCode=TemplateFor(TemplateCode)
			Else
				TemplateCode=TemplateIf(TemplateCode)
			End If
		ElseIf ForI>0 Then
			TemplateCode=TemplateFor(TemplateCode)
		ElseIf IfI>0 Then
			TemplateCode=TemplateIf(TemplateCode)
		ELse
			TemplateDo=TemplateCode
			Exit Function
		End If
		Call TemplateDo(TemplateCode)
		TemplateDo=TemplateCode
	End Function

	'==============================
	'函 数 名：TemplateFor
	'作    用：处理For
	'参    数：
	'==============================	
	Private Function TemplateFor(TemplateCode)
		Temp=GetFor(TemplateCode)
		TemplateTag=Split(Split(Temp,"{$For(")(1),",")(0)
		TemplatePar=Split(Split(Temp,",")(1),")")(0)
		TemplateBCode=Right(Temp,Len(Temp)-Len("{$For("&TemplateTag&","&TemplatePar&")$}"))
		TemplateBCode=Left(TemplateBCode,Len(TemplateBCode)-8)
		Select Case TemplateTag
			Case "Nav"
				TemplateFor=Replace(TemplateCode,Temp,FkNav(TemplateBCode,TemplatePar))
			Case "ArticleList"
				TemplateFor=Replace(TemplateCode,Temp,FkArticleList(TemplateBCode,TemplatePar))
			Case "ProductList"
				TemplateFor=Replace(TemplateCode,Temp,FkProductList(TemplateBCode,TemplatePar))
			Case "ProductVideoList"
				TemplateFor=Replace(TemplateCode,Temp,FkProductVideoList(TemplateBCode,TemplatePar))
			Case "ProductDownList"
				TemplateFor=Replace(TemplateCode,Temp,FkProductDownList(TemplateBCode,TemplatePar))
			Case "DownList"
				TemplateFor=Replace(TemplateCode,Temp,FkDownList(TemplateBCode,TemplatePar))
			Case "FriendsList"
				TemplateFor=Replace(TemplateCode,Temp,FkFriendsList(TemplateBCode,TemplatePar))
			Case "JobList"
				TemplateFor=Replace(TemplateCode,Temp,FkJobList(TemplateBCode,TemplatePar))
			Case "SubjectList"
				TemplateFor=Replace(TemplateCode,Temp,FkSubjectList(TemplateBCode,TemplatePar))
			Case "GBookList"
				TemplateFor=Replace(TemplateCode,Temp,FkGBookList(TemplateBCode,TemplatePar))
			Case Else
				TemplateFor=Replace(TemplateCode,Temp,"")
		End Select
	End Function

	'==============================
	'函 数 名：TemplateIf
	'作    用：处理If
	'参    数：
	'==============================	
	Private Function TemplateIf(TemplateCode)
		Dim Check1,Check2
		Dim t_TempArr,t_Temp,tempIf,myChange,myTemp
		Temp=GetIf(TemplateCode)
		myTemp=Temp
		myChange=0
		'截取IF嵌套
		If GetCount(Temp,"{$If")>1 Then
			myChange=1
			t_Temp=Right(Temp,Len(Temp)-5)
			t_Temp=Left(t_Temp,Len(t_Temp)-10)
			t_TempArr=Split(FKFun.RegExpTest("\{\$If\((.|\n)*?\{\$End If\$\}",t_Temp),"|-_-|")
			For Each t_Temp In t_TempArr
				Temp=Replace(Temp,t_Temp,"{FangkaIF}")
			Next
		End If
		TemplatePar=Split(Split(Temp,"{$If(")(1),")")(0)
		If1=GetIfOne(Temp,"{$If("&TemplatePar&")$}")
		If2=Replace(Temp,If1,"")
		If2=Replace(If2,"{$If("&TemplatePar&")$}","")
		If2=Left(If2,Len(If2)-10)
		If2=Right(If2,Len(If2)-8)
		If If2="{$Null$}" Then
			If2=""
		End If
		If myChange=1 Then
			Temp=MyTemp
			MyTemp=If1&"|-_无聊的间隔_-|"&If2
			For Each t_Temp In t_TempArr
				MyTemp=Replace(MyTemp,"{FangkaIF}",t_Temp,1,1)
			Next
			If1=Split(MyTemp,"|-_无聊的间隔_-|")(0)
			If2=Split(MyTemp,"|-_无聊的间隔_-|")(1)
		End If
		TempArr=Split(TemplatePar,",")
		If IsNumeric(TempArr(0)) And IsNumeric(TempArr(1)) Then
			Check1=CDBl(TempArr(0))
			Check2=CDBl(TempArr(1))
		Else
			Check1=TempArr(0)
			Check2=TempArr(1)
		End If
		Select Case TempArr(2)
			Case ">"
				If TempArr(0)>TempArr(1) Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "<"
				If TempArr(0)<TempArr(1) Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "="
				If TempArr(0)=TempArr(1) Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "<>"
				If TempArr(0)<>TempArr(1) Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case "<="
				If TempArr(0)<=TempArr(1) Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case ">="
				If TempArr(0)>=TempArr(1) Then
					TemplateIf=Replace(TemplateCode,Temp,If1)
				Else
					TemplateIf=Replace(TemplateCode,Temp,If2)
				End If
			Case Else
				TemplateIf=Replace(TemplateCode,Temp,"")
		End Select
		
	End Function
	

	'==============================
	'函 数 名：SearchPageChange
	'作    用：替换搜索模块页码参数
	'参    数：
	'==============================
	Public Function SearchPageChange(TemplateCode)
		Dim TempMUrl
		PageFirst=""
		PagePrev=""
		PageNext=""
		PageLast=""
		TempMUrl=SiteDir&"Search/Index.asp?SearchStr="&Server.URLEncode(SearchStr)&"&SearchType="&SearchType&"&SearchTemplate="&Server.URLEncode(SearchTemplate)&"&SearchField="&Server.URLEncode(SearchField)&"&SearchFieldList="&Server.URLEncode(SearchFieldList)&"&Page="
		If PageCounts>1 Then
			If PageNow>1 Then
				PageFirst=TempMUrl&"1"
				If PageNow=2 Then
					PagePrev=TempMUrl&"1"
				Else
					PagePrev=TempMUrl&(PageNow-1)
				End If
			End If
			If PageNow<PageLast Then
				PageNext=TempMUrl&(PageNow+1)
				PageLast=TempMUrl&PageLast
			End If
		End If
		TempArr=Split(FKFun.RegExpTest("\{\$SearchPageCode\(.*?\)\$\}",TemplateCode),"|-_-|")
		For Each Temp In TempArr
			If Temp<>"" Then
				TemplateCode=ReplaceTag(TemplateCode,Temp,CheckPageCode(Split(Split(Temp,"(")(1),")")(0),TempMUrl&"{Pages}"))
			End If
		Next
		TemplateCode=PageCodeChange(TemplateCode)
		SearchPageChange=TemplateCode
	End Function

	'==============================
	'函 数 名：GetFor
	'作    用：获取For字符串
	'参    数：
	'==============================	
	Private Function GetFor(TemplateCode)
		Temp=Split(TemplateCode,"{$For")(0)
		Temp=Replace(TemplateCode,Temp,"")
		TempArr=Split(Temp,"{$Next$}")
		GetFor=TempArr(0)&"{$Next$}"
		i=1
		While GetCount(GetFor,"{$For")<>GetCount(GetFor,"{$Next$}")
			GetFor=GetFor&TempArr(i)&"{$Next$}"
			i=i+1
		Wend
	End Function

	'==============================
	'函 数 名：GetIf
	'作    用：获取If字符串
	'参    数：
	'==============================	
	Private Function GetIf(TemplateCode)
		Temp=Split(TemplateCode,"{$If")(0)
		Temp=Replace(TemplateCode,Temp,"")
		TempArr=Split(Temp,"{$End If$}")
		GetIf=TempArr(0)&"{$End If$}"
		i=1
		While GetCount(GetIf,"{$If")<>GetCount(GetIf,"{$End If$}")
			GetIf=GetIf&TempArr(i)&"{$End If$}"
			i=i+1
		Wend
	End Function

	'==============================
	'函 数 名：GetIfOne
	'作    用：获取If字符串Else前
	'参    数：
	'==============================	
	Private Function GetIfOne(TemplateCode,IfCode)
		TempArr=Split(TemplateCode,"{$Else$}")
		GetIfOne=Replace(TempArr(0),IfCode,"")
		i=1
		While GetCount(GetIfOne,"{$If")<>GetCount(GetIfOne,"{$End If$}")
			GetIfOne=GetIfOne&TempArr(i)
			i=i+1
		Wend
	End Function

	'==============================
	'函 数 名：GetCount
	'作    用：判断字符串中相同字符的个数
	'参    数：
	'==============================	
	Private Function GetCount(Strs,Word)
		Dim N1,N2,N3
		N1=Len(Strs)
		N2=Len(Replace(Strs,Word,""))
		N3=Len(Word)
		GetCount=Clng(((N1-N2)/N3))
	End Function 
	
	'==============================
	'函 数 名：ReplaceTag
	'作    用：替换标签
	'参    数：
	'==============================
	Private Function ReplaceTag(MyStr,MyTag,MyChang)
		If IsNull(MyChang) Then
			ReplaceTag=Replace(MyStr,MyTag,"")
		Else
			ReplaceTag=Replace(MyStr,MyTag,MyChang)
		End If
	End Function
	
	'==============================
	'函 数 名：SearchChange
	'作    用：搜索页参数
	'参    数：
	'==============================
	Public Function SearchChange(TemplateCode)
		TemplateCode=ReplaceTag(TemplateCode,"{$SearchStr$}",SearchStr)
		TemplateCode=ReplaceTag(TemplateCode,"{$SearchType$}",SearchType)
		SearchChange=TemplateCode
	End Function
	
	'==============================
	'函 数 名：GetTemplate
	'作    用：获取模板代码
	'参    数：
	's_FileName      获取的模板文件
	's_TemplateId    获取类型
	's_IsIndex       是否菜单首页
	's_MenuTepmlate  菜单模板目录
	'==============================
	Public Function GetTemplate(s_FileName,s_TemplateId,s_IsIndex,s_MenuTepmlate)
		Dim TempFileName
		If s_MenuTepmlate<>"" Then
			s_MenuTepmlate=s_MenuTepmlate&"/"
		End If
		If s_IsIndex=1 And s_MenuTepmlate<>"" Then  '子菜单中的首页模块
			If Fk_Site_SkinTest=1 Then
				GetTemplate=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&s_MenuTepmlate&"index.html")
				Exit Function
			End If
			Sqlstr="Select Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Name='"&s_MenuTepmlate&"index'"
		ElseIf s_TemplateId=0 Then  '默认模板
			If Fk_Site_SkinTest=1 Then
				GetTemplate=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&s_MenuTepmlate&s_FileName&".html")
				Exit Function
			End If
			Sqlstr="Select Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Name='"&s_MenuTepmlate&s_FileName&"'"
		Else  '自定义模板
			Sqlstr="Select Fk_Template_Name,Fk_Template_Content From [Fk_Template] Where Fk_Template_Id=" & s_TemplateId
		End If
		Rs.Open Sqlstr,Conn,1,1
		If Not Rs.Eof Then
			TempFileName=Rs("Fk_Template_Name")
			GetTemplate=Rs("Fk_Template_Content")
		Else
			Call FKFun.ShowErr("模板未找到！",0)
		End If
		Rs.Close
		If Fk_Site_SkinTest=1 Then
			GetTemplate=FKFso.FsoFileRead(FileDir&"Skin/"&Fk_Site_Template&"/"&TempFileName&".html")
		End If
	End Function
	
	'==============================
	'函 数 名：GetHtmlSuffix
	'作    用：获取生成后缀
	'参    数：
	'==============================
	Public Function GetHtmlSuffix()
		Select Case Fk_Site_HtmlSuffix
			Case 0
				GetHtmlSuffix=".html"
			Case 1
				GetHtmlSuffix=".htm"
			Case 2
				GetHtmlSuffix=".shtml"
			Case 3
				GetHtmlSuffix=".xml"
		End Select
	End Function
	
	'==============================
	'函 数 名：ReChangeField
	'作    用：多余标签清理
	'参    数：
	'TemplateCode  要处理的字符串
	'==============================
	Public Function ReChangeField(TemplateCode)
		TemplateCode=FKFun.ReplaceTest("\{\$aaaaaaaaaaaaaaaaaabbbbbbGBookList.*?\$\}","",TemplateCode)
		TemplateCode=FKFun.ReplaceTest("\{\$aaaaaaaaaaaaaaaaaabbbbbbField\_.*?\$\}","",TemplateCode)
		TemplateCode=FKFun.ReplaceTest("\{\$aaaaaaaaaaaaaaaaaabbbbbbFieldList\_.*?\$\}","",TemplateCode)
		ReChangeField=TemplateCode
	End Function
End Class
%>
