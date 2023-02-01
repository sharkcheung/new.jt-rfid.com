<!--#Include File="../../Inc/Config.asp"-->
<!--#Include File="../../inc/md5.asp"-->
<!--#Include File="../config.asp"-->
<script language="jscript" runat="server">    
Array.prototype.get = function(x) { 
	return this[x]; 
};    
function parseJSON(strJSON) { 
	return eval("(" + strJSON + ")"); 
}    
</script>

<%Response.Addheader "Content-Type","text/json; charset=utf-8"
'==========================================
'文 件 名：api/publishArticle/Index.asp
'文件用途：对接臻至软文发布新闻推送接口
'==========================================
Dim Fk_Article_Title,Fk_Article_Content,Fk_Article_Click,Fk_Article_Show,Fk_Article_Time,Fk_Article_Pic,Fk_Article_PicBig,Fk_Article_Template,Fk_Article_FileName,Fk_Article_Subject,Fk_Article_Recommend,Fk_Article_Keyword,Fk_Article_Description,Fk_Article_From,Fk_Article_Color,Fk_Article_Url,Fk_Article_Field,Fk_Article_onTop,Fk_Article_px,Fk_Article_Seotitle
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Dir,Fk_Article_Module
Dim Temp2,KeyWordlist,kwdrs,ki,host
dim appendFrom,getpostjson
dim Fk_Article_Copyright,Fk_Article_CopyrightInfo,Fk_Article_CopyrightFs,Fk_Article_CopyrightFt,Fk_Article_CopyrightCl,ArticleUrl

function bytes2bstr(vin)
	dim bytesstream,stringreturn
	set bytesstream = server.CreateObject("adodb.stream")
	bytesstream.type = 2
	bytesstream.open
	bytesstream.writeText vin
	bytesstream.position = 0
	bytesstream.charset = "utf-8"'或者gb2312
	bytesstream.position = 2
	stringreturn = bytesstream.readtext
	bytesstream.close
	set bytesstream = nothing
	bytes2bstr = stringreturn
end function


'==============================
'函 数 名：ShowString
'作    用：判断字符串长度
'参    数：
'需进行判断的文本CheckStr
'限定最短ShortLen
'限定最长LongLen
'验证类型CheckType（0两头限制，1限制最短，2限制最长）
'过短提示LongStr
'过长提示LongStr，
'==============================
sub checkString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)

	If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
		response.Write("{""success"":false,""message"":"""&ShortErr&"""}")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
		response.Write("{""success"":false,""message"":"""&LongErr&"""}")
		Call FKDB.DB_Close()
		Response.End()
	End If
End sub

'==============================
'函 数 名：StringLength
'作    用：判断字符串长度
'参    数：需进行判断的文本Txt
'==============================
Function StringLength(Txt)
	dim x,y,ii
	Txt=Trim(Txt)
	x=Len(Txt)
	y=0
	For ii = 1 To x
		If Asc(Mid(Txt,ii,1))<=2 or Asc(Mid(Txt,ii,1))>255 Then
			y=y + 2
		Else
			y=y + 1
		End If
	Next
	StringLength=y
End Function

Function getIP()
	Dim sIPAddress, sHTTP_X_FORWARDED_FOR
	 sHTTP_X_FORWARDED_FOR = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	If sHTTP_X_FORWARDED_FOR = "" Or InStr(sHTTP_X_FORWARDED_FOR, "unknown") > 0 Then
		sIPAddress = Request.ServerVariables("REMOTE_ADDR")
	ElseIf InStr(sHTTP_X_FORWARDED_FOR, ",") > 0 Then
		sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ",") -1)
	ElseIf InStr(sHTTP_X_FORWARDED_FOR, ";") > 0 Then
		sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ";") -1)
	Else
		sIPAddress = sHTTP_X_FORWARDED_FOR
	End If
	getIP = Trim(Mid(sIPAddress, 1, 15))
End Function

dim readjson,json,code,n,domainname,curtime,strcode,requestip
requestip = getIP()
if instr(whiteIPList,requestip)=0 then
	response.Write("{""success"":false,""message"":""no permission""}")
	response.end() 
end if

code = Request.ServerVariables("HTTP_CODE")

if IsEmpty(code) then
	response.Write("{""success"":false,""message"":""code is null""}")
	response.end() 
end if

getpostjson=Request.TotalBytes '得到字节数
if getpostjson=0 then
	response.Write("{""success"":false,""message"":""post data is null""}")
	response.End()
end if

domainname = Request.ServerVariables("SERVER_NAME")
curtime = year(date)&right("0"&month(date),2)&right("0"&day(date),2)


strcode = md5(domainname&curtime,32)

if code<>strcode then
	response.Write("{""success"":false,""message"":""校验失败""}")
	response.End()
end if

readjson=Request.BinaryRead(getpostjson) '二进制方式来读取客户端使用POST传送方法所传递的数据
json = bytes2bstr(readjson) '二进制转化
dim jsons
set jsons = parseJSON(json)
on error resume next
Fk_Article_Title = FKFun.HTMLEncode(Trim(jsons.title))
Fk_Article_Seotitle = Fk_Article_Title
Fk_Article_Keyword = FKFun.HTMLEncode(Trim(jsons.keyword))
Fk_Article_Pic = FKFun.HTMLEncode(Trim(jsons.articleimg))
Fk_Article_PicBig = Fk_Article_Pic
Fk_Article_Content = jsons.content
if Err.Number <> 0 then 
	Err.clear
end if

Fk_Module_Id=PublicModuleID
Fk_Article_Color=""
Fk_Article_Description=""
Fk_Article_Url=""
Fk_Article_From="本站"

Fk_Article_FileName=""
Fk_Article_Recommend=",0,"
Fk_Article_Subject=",0,"
Fk_Article_Template=0
Fk_Article_Show=1
Fk_Article_onTop=0
Fk_Article_px=0
dim rnd_num
RANDOMIZE
rnd_num=INT(100*RND)+1
Fk_Article_click=rnd_num
Fk_Article_Copyright=0
Fk_Article_CopyrightInfo=""
Fk_Article_CopyrightFs=""
Fk_Article_CopyrightFt=""
Fk_Article_CopyrightCl=""


Call checkString(Fk_Article_Title,1,255,0,"请输入内容标题！","内容标题不能大于255个字符！")
Call checkString(Fk_Article_From,1,50,0,"请输入内容来源！","内容来源不能大于50个字符！")
Call FKFun.ShowNum(Fk_Article_px,"排序必须为数字！")
Call checkString(Fk_Article_Seotitle,0,255,2,"请输入内容SEO标题！","内容SEO标题不能大于255个字符！")
Call checkString(Fk_Article_Keyword,0,255,2,"请输入内容SEO关键词！","内容SEO关键词不能大于255个字符！")
Call checkString(Fk_Article_Description,0,255,2,"请输入内容SEO描述！","内容SEO描述不能大于255个字符！")
Call checkString(Fk_Article_Url,0,255,2,"请输入内容转向链接！","内容转向链接不能大于255个字符！")
If Fk_Article_Url="" Then
	Call checkString(Fk_Article_Content,10,1,1,"请输入内容内容，不少于10个字符！","内容内容不能大于1个字符！")
End If
Call checkString(Fk_Article_Pic,0,255,2,"请输入内容题图路径！","内容题图小图路径不能大于255个字符！")
Call checkString(Fk_Article_PicBig,0,255,2,"请输入内容题图路径！","内容题图大图路径不能大于255个字符！")
Call checkString(Fk_Article_FileName,0,100,2,"生成文件名不符合标准！","内容文件名不能大于100个字符！")
'Call checkString(Fk_Article_FileName,2,1,1,"生成文件名不能为空","生成内容不能大于1个字符！")
Call FKFun.ShowNum(Fk_Article_Template,"请选择模板！")
Call FKFun.ShowNum(Fk_Article_Show,"请选择内容是否显示！")
Call FKFun.ShowNum(Fk_Article_click,"请输入正确的点击量！")
Call FKFun.ShowNum(Fk_Article_Copyright,"参数错误，请刷新页面！")
Call checkString(Fk_Article_CopyrightInfo,0,200,0,"请输入转载声明！","转载声明内容不能大于200个字符！")
Call checkString(Fk_Article_CopyrightFs,0,50,0,"请选择字体大小！","字体大小不能大于50个字符！")
Call checkString(Fk_Article_CopyrightFt,0,50,0,"请选择是否加粗！","粗体样式不能大于50个字符！")
Call checkString(Fk_Article_CopyrightCl,0,50,0,"请选择字体颜色！","字体颜色不能大于50个字符！")

Call FKDB.DB_Open()
Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
Rs.Open Sqlstr,Conn,1,1
If Not Rs.Eof Then
	Fk_Module_Menu=Rs("Fk_Module_Menu")
Else
	response.Write("{""success"":false,""message"":""栏目不存在！""}")
	Rs.Close
	Call FKDB.DB_Close()
	Response.End()
End If
Rs.Close
If SiteDelWord=1 Then
	TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("../../admin/DelWord.dat")))," ")
	For Each Temp In TempArr
		If Temp<>"" Then
			Fk_Article_Content=Replace(Fk_Article_Content,Temp,"**")
			Fk_Article_Title=Replace(Fk_Article_Title,Temp,"**")
			Fk_Article_Seotitle=Replace(Fk_Article_Seotitle,Temp,"**")
			Fk_Article_Keyword=Replace(Fk_Article_Keyword,Temp,"**")
		End If
	Next
End If

Function CheckFields(FieldsName,TableName)
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
End Function

'新功能，追加转载声明
'2014年12月31日
'middy241@163.com
if CheckFields("Fk_Article_Copyright","Fk_Article")=false then
	conn.execute("alter table Fk_Article add column Fk_Article_Copyright int default 0")
	conn.execute("alter table Fk_Article add column Fk_Article_CopyrightInfo varchar(200) null")
	conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFs varchar(50) null")
	conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFt varchar(50) null")
	conn.execute("alter table Fk_Article add column Fk_Article_CopyrightCl varchar(50) null")
end if

Sqlstr="Select * From [Fk_Article] Where Fk_Article_Module="&Fk_Module_Id&" And (Fk_Article_Title='"&Fk_Article_Title&"'"
If Fk_Article_FileName<>"" Then
	Sqlstr=Sqlstr&" Or Fk_Article_FileName='"&Fk_Article_FileName&"'"
End If
Sqlstr=Sqlstr&")"
Rs.Open Sqlstr,Conn,1,3
If Rs.Eof Then
	Application.Lock()
	Rs.AddNew()
	Rs("Fk_Article_Title")=Fk_Article_Title
	Rs("Fk_Article_Color")=Fk_Article_Color
	Rs("Fk_Article_From")=Fk_Article_From
	Rs("Fk_Article_Seotitle")=Fk_Article_Seotitle
	Rs("Fk_Article_Keyword")=Fk_Article_Keyword
	Rs("Fk_Article_Field")=Fk_Article_Field
	Rs("Fk_Article_Description")=Fk_Article_Description
	Rs("Fk_Article_Url")=Fk_Article_Url
	Rs("Fk_Article_Show")=Fk_Article_Show
	Rs("Fk_Article_click")=Fk_Article_click
	Rs("Fk_Article_Pic")=Fk_Article_Pic
	Rs("Fk_Article_PicBig")=Fk_Article_PicBig
	Rs("Fk_Article_Content")=Fk_Article_Content
	Rs("Fk_Article_Recommend")=Fk_Article_Recommend
	Rs("Fk_Article_Subject")=Fk_Article_Subject
	Rs("Fk_Article_Module")=Fk_Module_Id
	Rs("Fk_Article_Menu")=Fk_Module_Menu
	Rs("Fk_Article_FileName")=Fk_Article_FileName
	Rs("Fk_Article_Template")=Fk_Article_Template
	Rs("Fk_Article_Ip")=Fk_Article_onTop
	
	Rs("Fk_Article_Copyright")=Fk_Article_Copyright
	Rs("Fk_Article_CopyrightInfo")=Fk_Article_CopyrightInfo
	Rs("Fk_Article_CopyrightFs")=Fk_Article_CopyrightFs
	Rs("Fk_Article_CopyrightFt")=Fk_Article_CopyrightFt
	Rs("Fk_Article_CopyrightCl")=Fk_Article_CopyrightCl
	
	Rs("Px")=Fk_Article_px
	Rs.Update()
	Application.UnLock()
	
	Fk_Article_Id = Rs("Fk_Article_Id")
	set mrs=conn.execute("select Fk_Module_Dir from Fk_Module where Fk_Module_Id="&Fk_Module_Id)
	if not mrs.eof then
		Fk_Module_Dir=mrs("Fk_Module_Dir")
	end if
	mrs.close
	set mrs=nothing
	If Fk_Module_Dir<>"" Then
		ArticleUrl=Fk_Module_Dir&"/"
	Else
		ArticleUrl="Article"&Fk_Module_Id&"/"
	End If
	If Fk_Article_FileName<>"" Then
		ArticleUrl=ArticleUrl&Fk_Article_FileName&".html"
	Else
		ArticleUrl=ArticleUrl&Rs("Fk_Article_Id")&".html"
	End If
	If SiteHtml=1 and sitetemplate<>"wap" Then
		ArticleUrl="/html"&SiteDir&ArticleUrl
	Else
		ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
	End If
	
	response.Write("{""success"":true,""message"":""内容发布成功！"",""articleid"":"&Rs("Fk_Article_Id")&",""articleurl"":""http://"&domainname&ArticleUrl&"""}")
	'插入日志
	dim log_content,log_ip,log_user
	log_content="添加内容：【"&Fk_Article_Title&"】"
	log_user="zhenzhi"
	
	log_ip=FKFun.getIP()
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
Else
	response.Write("{""success"":false,""message"":""该内容标题已经被占用，请重新填写！""}")
End If
Rs.Close
%>
<!--#Include File="../../Code.asp"-->
