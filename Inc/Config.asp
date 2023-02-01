<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>

<%
Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
Response.Expires=-999
Session.Timeout=999
Dim StartTime,EndTime
StartTime=Timer()
'==========================================
'文 件 名：Config.asp
'文件用途：系统配置
'版权所有：深圳企帮
'==========================================
%>

<!--#Include File="../Class/Cls_DB.asp"-->
<!--#Include File="../Class/Cls_Fso.asp"-->
<!--#Include File="../Class/Cls_Fun.asp"-->
<%

'定义页面常量
Dim Id,i,j,PageErr,Types,Temp,TempArr,Login,Fso,F,objAdoStream,PageArr,TemplateTempArr,TemplateTemp
Dim Conn,Rs,Sqlstr
dim Fk_Site_PageSize,TempPageSize,Fk_Site_SkinTest
Dim SearchStr,SearchType,SearchTemplate,SearchField,SearchFieldList,Fk_Site_Sign,Fk_Site_PageSign,Fk_Site_SysHidden,Fk_Site_Html,Fk_Site_HtmlType,Fk_Site_HtmlSuffix
Dim PageNow,PageCounts,PageSizes,PageAll
Dim SiteName,SiteUrl,SiteKeyword,SiteDescription,SiteOpen,SiteTemplate,SiteHtml,SiteData,SiteDir,SiteToPinyin,FetionNum,FetionPass,SiteQQ,SiteNoTrash,SiteMini,SiteDelWord,SiteDBDir,SiteTest,SiteFlash,Tel,Tel400,Fax,Beian,Lianxiren,Email,Add,Kfid,Tjid,SiteLogo,Sitepic1,Sitepic2,Sitepic3,Sitepic4,Sitepic5,Sitepicurl1,Sitepicurl2,Sitepicurl3,Sitepicurl4,Sitepicurl5,Sitepictext1,Sitepictext2,Sitepictext3,Sitepictext4,Sitepictext5,Bianjiqi,SysVersion,SysVersionTime,KfUrl,TjUrl,Site301,SiteSeoTitle,MiniAppId,MiniAppKey,AdminPath,NewAdminPath,ImgCdnUrl,CssCdnUrl,JsCdnUrl,FileCdnUrl,isCDN,wapUrl,PublicModuleID
Dim FKDB,FKFun,FKFso
Dim FKModuleId,FKModuleName
Dim Templates,PageFirst,PageNext,PagePrev,PageLast
Dim sTemp
Dim adminlx

'置默认值
Set FKDB=New Cls_DB
Set FKFun=New Cls_Fun
Set FKFso=New Cls_Fso
Login=False
FKModuleId=Split("3,1,2,7,4,5",",")
FKModuleName=Split("单页栏目,新闻栏目,图文栏目,下载栏目,留言栏目,转向链接",",")
%>
<!--#Include File="Site.asp"-->
<!--#Include File="Conn.asp"-->
<!--#Include File="System.asp"-->
<%
If SiteFlash=0 Then
	sTemp=""
Else
	sTemp="Index.asp"
End If
If len(NewAdminPath)=0 Then
	NewAdminPath="Admin-new"
End If
If len(AdminPath)=0 Then
	AdminPath="Admin"
End If
%>

