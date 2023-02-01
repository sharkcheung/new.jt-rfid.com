<!--#Include File="CheckToken.asp"-->
<%
'==========================================
'文 件 名：Json.asp
'文件用途：生成远程请求需要的json数据
'版权所有：企帮网络www.qebang.cn
'==========================================
'验证token

dim data,Sql,strSql
dim Fk_Word_Id,Fk_Word_level,Fk_Word_url,Fk_Word_Name
	
dim op

op=request("op")
if op="get_newslist" then
	response.write ShowModuleSelect(1,"")
	' Sql="Select Fk_Module_id,Fk_Module_Name From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]=0 order by Fk_Module_order desc,Fk_Module_id"
	' data = GetJSON(Sql,"newslist")
	' data = server.htmlencode(data) 
	' response.write callback&"("&data&")"
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
elseif op = "get_innerlist" then
	Sqlstr="Select * From [Fk_Word] Order By Fk_Word_level desc,Fk_Word_Id Asc"
	response.write GetJSON(Sqlstr,"list")

elseif op="inner_edit" then
	Fk_Word_Name=FKFun.HTMLEncode(Trim(Request("Fk_Word_Name")))
	Fk_Word_level=FKFun.HTMLEncode(Trim(Request("Fk_Word_level")))
	Fk_Word_url=FKFun.HTMLEncode(Trim(Request("Fk_Word_url")))
	Fk_Word_Id=Trim(Request("Fk_Word_Id"))
	if Fk_Word_Id = "" then Fk_Word_Id = 0
	Fk_Word_Id = cint(Fk_Word_Id)
	if not isnumeric(Fk_Word_Id) then Fk_Word_Id = 0
	
	if len(Fk_Word_level)=0 or len(Fk_Word_Name)=0 or len(Fk_Word_url)=0 then
		response.write "参数不正确！"
	else		
		if Fk_Word_Id>0 then
			sql = "select top 1 * from [Fk_Word] where [Fk_Word_Id]="&Fk_Word_Id&""
			rs.Open sql,conn,1,3
			if rs.eof then
				rs.close
				response.write "修改的内容不存在"
			else
				rs("Fk_Word_level")=Fk_Word_level
				rs("Fk_Word_Name")=Fk_Word_Name
				rs("Fk_Word_url")=Fk_Word_url
				rs.update
				rs.close
				response.write "修改成功"
			end if
		else
			sql = "select top 1 * from [Fk_Word]"
			rs.Open sql,conn,1,3
			rs.addnew
			rs("Fk_Word_level")=Fk_Word_level
			rs("Fk_Word_Name")=Fk_Word_Name
			rs("Fk_Word_url")=Fk_Word_url
			rs.update
			rs.close
			response.write "添加成功"
		end if
	end if
elseif op="inner_del" then
	Fk_Word_Id=Trim(Request("Fk_Word_Id"))
	if Fk_Word_Id = "" then Fk_Word_Id = 0
	Fk_Word_Id = cint(Fk_Word_Id)
	if not isnumeric(Fk_Word_Id) then Fk_Word_Id = 0
	
	if Fk_Word_Id>0 then
		sql = "delete from [Fk_Word] where [Fk_Word_Id]="&Fk_Word_Id&""
		conn.execute(sql)
		response.write "删除成功"
	else
		response.write "参数不正确！"
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
	set Rs =conn.execute(strSql)
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

'==============================
'函 数 名：ShowModuleSelect
'作    用：输出ModuleSelect列表
'参    数：要输出的菜单MenuIds
'==============================
Public Function ShowModuleSelect(MenuIds,AutoId)
	Call ShowModuleSelectM(MenuIds,0,"",AutoId)
End Function
Public Function ShowModuleSelectM(MenuIds,LevelId,TitleBack,AutoId)
	Dim Rs2,TitleBacks
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	If LevelId=0 Then
		TitleBack="pid=""0"""
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Type=1 and Fk_Module_Menu="&MenuIds&" And Fk_Module_Level="&LevelId&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs2.Open Sqlstr,Conn,1,3
	While Not Rs2.Eof
	%><data id="<%=Rs2("Fk_Module_Id")%>" <%=TitleBack%>><%=Rs2("Fk_Module_Name")%></data><%
		If LevelId=0 Then
			TitleBacks="pid="""&Rs2("Fk_Module_Id")&""""
		Else
			TitleBacks=TitleBack
		End If
		Call ShowModuleSelectM(MenuIds,Rs2("Fk_Module_Id"),TitleBacks,AutoId)
		Rs2.MoveNext
	Wend
	Rs2.Close
	Set Rs2=Nothing
End Function

function newsclasslist(Fk_Module_Level,strlist) '获取新闻类列表函数
	dim newsclass,iii,sql,rs2,sql2,Fk_Module_Name,Fk_Module_id,Fk_Module_Name2,Fk_Module_id2,list
	iii=1
	Sql="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]="&Fk_Module_Level
	set rs2 = conn.execute(sql)
	do until rs2.EOF
		Fk_Module_Name=rs2("Fk_Module_Name")
		Fk_Module_id=rs2("Fk_Module_id")
		strlist=strlist&","&Fk_Module_id&":"&Fk_Module_Name
		' if Rs("Fk_Module_Level")<>0 then
			strlist =  newsclasslist(Fk_Module_id,strlist)
			response.write strlist
		' end if
		' Set Rs2=Server.Createobject("Adodb.RecordSet")
		' Sql2="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]="&Fk_Module_id
		' Rs2.Open Sql2,Conn,1,1
		' do until rs2.EOF
		' Fk_Module_Name2=Rs2("Fk_Module_Name")
		' Fk_Module_id2=Rs2("Fk_Module_id")
		' newsclass=newsclass&"|"&Fk_Module_id2&":"&Fk_Module_Name2
		' rs2.MoveNext
		' iii=iii+1
		' loop
		' rs2.close
		' Set Rs2=Nothing
		
	rs2.MoveNext
	loop
	    
	rs2.close
	Set rs2=Nothing
	' response.write newsclasslist
	' newsclasslist=list
end function

%><!--#Include File="../Code.asp"-->