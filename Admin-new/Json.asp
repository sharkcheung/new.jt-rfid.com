<!--#Include File="CheckToken.asp"-->
<%
'==========================================
'文 件 名：Json.asp
'文件用途：生成远程请求需要的json数据
'版权所有：企帮网络www.qebang.cn
'==========================================
'验证token

dim data,Sql
	
dim callback,op

op=request("op")
callback=request("callback")
if op="get_newslist" then
	' if callback ="" then
		' response.end
	' end if
	Sql="Select Fk_Module_id,Fk_Module_Name From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]=0 order by Fk_Module_order desc,Fk_Module_id"
	data = newsclasslist
	data = server.htmlencode(data) 
	response.write callback&"({""newslist"":"""&data&"""})"
elseif op="collectDo" then
	dim strTitle,strContent,strModuleid
	strTitle=FKFun.HTMLEncode(Trim(Request("title")))
	strContent=FKFun.HTMLEncode(Trim(Request("content")))
	strModuleid=Trim(Request("moduleid"))
	if strModuleid = "" then strModuleid = 0
	strModuleid = cint(strModuleid)
	if not isnumeric(strModuleid) then strModuleid = 0
	
	if len(strTitle)=0 or len(strContent)=0 or len(strModuleid)=0 then
		response.write "参数不正确！" & len(strTitle) & "--" & len(strContent) & "--" & (strModuleid-1)
	else
		response.write sql
		sql = "select top 1 * from [FK_Article] where [Fk_Article_Title]='"&strTitle&"'"
		rs.Open sql,conn,1,3
		if rs.recordcount=0 then
			rs.addnew
			rs("Fk_Article_Title")=strTitle
			rs("Fk_Article_Content")=strContent
			rs("Fk_Article_From")="互联网"
			rs("Fk_Article_Module")=strModuleid
			rs("Fk_Article_Menu")=1
			rs("Fk_Article_Recommend")=",0,"
			rs("Fk_Article_Subject")=",0,"
			rs.update
			rs.close
			response.write "数据采集成功！"
		else
			rs.close
			response.write "已存在此标题的内容！"
		end if
	end if
end if

'转换中文为unicode
function URLEncoding(vstrIn)
    dim i
    dim strReturn,ThisChr,innerCode,Hight8,Low8
    strReturn = ""
    for i = 1 to Len(vstrIn)
        ThisChr = Mid(vStrIn,i,1)
        If Abs(Asc(ThisChr)) < &HFF then
            strReturn = strReturn & ThisChr
        else
            innerCode = Asc(ThisChr)
            If innerCode < 0 then
                innerCode = innerCode + &H10000
            end If
            Hight8 = (innerCode  and &HFF00)\ &HFF
            Low8 = innerCode and &HFF
            strReturn = strReturn & "%" & Hex(Hight8) &  "%" & Hex(Low8)
        end If
    next
    URLEncoding = strReturn
end function

' 可以外部调用的公共方法 
Public Function GetJSON(strSql,Root)
	Dim returnStr 
	Dim i 
	Dim oneRecord 
	
	' 获取数据 
	Rs.open strSql,conn,1,1 
	' 生成JSON字符串 
	If Rs.eof=false And Rs.Bof=false Then 
		'returnStr="{ "&Chr(13)& Chr(9) & Chr(9) & Root & ":{ "& Chr(13) & Chr(9) & Chr(9) &Chr(9) & Chr(9) &"records:[ " & Chr(13)
		returnStr="{ "&Chr(13)& Chr(9) & Chr(9) & """" & Root & """" & ": "& Chr(13) & Chr(9) & Chr(9) &Chr(9) & Chr(9) &"[ " & Chr(13)
		
		While(Not Rs.Eof)
			' ------- 
			oneRecord= Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "{ " 
			
			For i=0 To Rs.Fields.Count -1 
				oneRecord=oneRecord & Chr(34) & Rs.Fields(i).Name&Chr(34) &":" 
				oneRecord=oneRecord & Chr(34) & Rs.Fields(i).Value&Chr(34) &"," 
			Next 
		'去除记录最后一个字段后的"," 
		oneRecord=Left(oneRecord,InStrRev(oneRecord,",")-1) 
		oneRecord=oneRecord & "}," & Chr(13)
		'------------ 
		returnStr=returnStr & oneRecord 
		Rs.MoveNext 
		Wend 
		' 去除所有记录数组后的"," 
		returnStr=Left(returnStr,InStrRev(returnStr,",")-1) & Chr(13)
		'returnStr=returnStr & Chr(9) & Chr(9) &Chr(9) & Chr(9) &"]" & Chr(13) & Chr(9) & Chr(9) & "}" &Chr(13) & "}" 
		 returnStr=returnStr & Chr(9) & Chr(9) &Chr(9) & Chr(9) &"]" & Chr(13) & Chr(9) & Chr(9) & "}" &Chr(13) & "" 
	End If
	GetJSON=returnStr 
End Function 

function newsclasslist() '获取新闻类列表函数
	dim newsclass,iii,sql,rs2,sql2,Fk_Module_Name,Fk_Module_id,Fk_Module_Name2,Fk_Module_id2
	newsclass="<select size='1' name='D1'>"
	iii=1
	Sql="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]=0 "
	Rs.Open Sql,Conn,1,1
	do until rs.EOF
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_id=Rs("Fk_Module_id")
		newsclass=newsclass&"<option value='"&Fk_Module_id&"'>"&Fk_Module_Name&"↓</option>"
		
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Sql2="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]="&Fk_Module_id
		Rs2.Open Sql2,Conn,1,1
		do until rs2.EOF
		Fk_Module_Name2=Rs2("Fk_Module_Name")
		Fk_Module_id2=Rs2("Fk_Module_id")
		newsclass=newsclass&"<option value='"&Fk_Module_id2&"'>　"&Fk_Module_Name2&"</option>"
		rs2.MoveNext
		iii=iii+1
		loop
		rs2.close
		Set Rs2=Nothing
		
	rs.MoveNext
	loop
	    
	rs.close
	Set Rs=Nothing
	
	newsclass=newsclass&"</select>"
	newsclasslist=newsclass
end function

Response.Cookies("FkAdminName")	=strMobile
Response.Cookies("Usertype")	=strUsertype
Response.Cookies("token")		=strToken
Response.Cookies("strkfurl")	=strkfurl
Response.Cookies("strtjurl")	=strtjurl
Response.Cookies("FkAdminPass")	=Md5(Md5(strToken,32),16)
Response.Cookies("FkAdminIp")	=Request.ServerVariables("REMOTE_ADDR")
Response.Cookies("FkAdminTime")	=Now()
Response.Cookies("FkAdminName").Expires=#May 10,2030#
Response.Cookies("FkAdminPass").Expires=#May 10,2030#

%><!--#Include File="../Code.asp"-->