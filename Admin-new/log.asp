<!--#Include File="AdminCheck.asp"-->
<%
response.charset="utf-8"
session.codepage=65001
'==========================================
'文 件 名：Word.asp
'文件用途：关键词链接管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System11") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call LogList() '日志列表
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：LogList()
'作    用：日志列表
'参    数：
'==========================================
Sub LogList()
	on error resume next
	rs.open "select id from newTB_log",conn,1,1
	if err.number<>0 then
		Sqlstr="create table newTB_log(id integer identity(1,1) primary key,log_content varchar(255),log_time date default now(),log_ip varchar(20),log_user varchar(30))"
		Conn.Execute(Sqlstr)
		err.clear
	end if
	rs.close
	Session("NowPage")=FkFun.GetNowUrl()
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	Dim SearchStr,d1,SearchStr1,ssj
	ssj=Trim(Request.QueryString("ssj"))
	SearchStr1=URLDecode(Trim(Request.QueryString("SearchStr")))
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	d1=FkFun.HTMLEncode(Trim(Request.QueryString("un")))
	'response.write ssj
%>
<div id="BoxTop" style="width:98%;"><span>工作记录</span></div>
<div id="ListNav" style="width:98%; margin-left:8px;border-left:1px solid #7998B7;border-right:1px solid #7998B7;">
    <ul>
        <li>账号：<select name="D1" id="D1" style="vertical-align:middle;" class="Input">
      <option value="0" <%if d1="0" then response.write "selected"%>>请选择账号</option>
      <option value="-1" <%if d1="-1" then response.write "selected"%>>客服账号</option>
<%set rs= conn.execute("select * from Fk_Admin where Fk_Admin_Limit=1")
if not rs.eof then
do while not rs.eof%>
	<option value="<%=rs("Fk_Admin_LoginName")%>" <%if d1=trim(rs("Fk_Admin_LoginName")) then response.write "selected"%>><%=rs("Fk_Admin_LoginName")%></option>
<%rs.movenext
loop
end if
rs.close
%>
</select><input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle;"/>&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','log.asp?Type=1&SearchStr='+encodeURIComponent(document.all.SearchStr.value)+'&un='+document.all.D1.value+'&ssj='+$('.ssj:checked').val());" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/> <input type="radio" id="ssj1" name="ssj" value="1" class="Input ssj" style="vertical-align:middle;border:0;background:none;" <%if ssj="1" or ssj="" then response.write "checked"%> onclick="SetRContent('MainRight','log.asp?Type=1&SearchStr='+encodeURIComponent(document.all.SearchStr.value)+'&un='+document.all.D1.value+'&ssj='+this.value);"><label for="ssj1" style="vertical-align:middle;">不限时间</label> <input type="radio" id="ssj2" name="ssj" value="2" onclick="SetRContent('MainRight','log.asp?Type=1&SearchStr='+encodeURIComponent(document.all.SearchStr.value)+'&un='+document.all.D1.value+'&ssj='+this.value);" <%if ssj="2" then response.write "checked"%>  class="Input ssj" style="vertical-align:middle;border:0;background:none;"><label for="ssj2" style="vertical-align:middle;">今天</label> <input type="radio" id="ssj3" name="ssj" value="3" onclick="SetRContent('MainRight','log.asp?Type=1&SearchStr='+encodeURIComponent(document.all.SearchStr.value)+'&un='+document.all.D1.value+'&ssj='+this.value);" <%if ssj="3" then response.write "checked"%>  class="Input ssj" style="vertical-align:middle;border:0;background:none;"><label for="ssj3" style="vertical-align:middle;">昨天</label></li>
    </ul>
</div>
<div id="ListContent" style="width:98%;margin-left:8px;background-color:#FFF;border-left:1px solid #7998B7;border-right:1px solid #7998B7;border-bottom:1px solid #D7E0E7;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">操作账号</td>
            <td align="center" class="ListTdTop">操作记录</td>
            <td align="center" class="ListTdTop">操作时间</td>
            <td align="center" class="ListTdTop">操作IP</td>
        </tr>
	<%dim wherestr
	Sqlstr="Select n.*,f.Fk_Admin_Limit From [newTB_log] as n left join Fk_Admin f on n.log_user=f.Fk_Admin_LoginName where 1=1"
	If len(SearchStr) Then
		Sqlstr=Sqlstr&" And log_content Like '%%"&SearchStr&"%%'"
	End If
	if len(d1)<>0 then
		if d1<>0 then
			if d1="-1" then
				Sqlstr=Sqlstr&" And f.Fk_Admin_Limit<>1"
			else
				Sqlstr=Sqlstr&" And n.log_user='"&d1&"'"
			end if
		end if
	end if
	if len(ssj)<>0 then
		if ssj="2" then
			Sqlstr=Sqlstr&" And datediff('d',n.log_time,now())=0"
		elseif ssj="3" then
			Sqlstr=Sqlstr&" And datediff('d',n.log_time,now())=1"
		end if
	end if
	Sqlstr=Sqlstr&" Order By log_time desc"
	'response.write sqlstr
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
%>
        <tr>
            <td height="20" align="center"><%=Rs("id")%></td>
            <td align="center"><%if d1="-1" or rs("Fk_Admin_Limit")<>1 then 
				response.write "客服人员"
				else
				response.write Rs("log_user")
				end if%></td>
            <td align="left" style=" width: 570px;word-break: break-all;word-wrap: break-word;"><%=Rs("log_content")%></td>
            <td align="center"><%=Rs("log_time")%></td>
            <td align="center"><%=Rs("log_ip")%></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" align="center" colspan="5">&nbsp;<%Call FKFun.ShowPageCode("log.asp?Type=1&un="&d1&"&SearchStr="&server.urlencode(SearchStr1)&"&ssj="&ssj&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="5 align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
</div>
<div id="ListBottom">

</div>

<%
End Sub

Function URLDecode(enStr)
   dim deStr,strSpecial
   dim c,i,v
     deStr=""
     strSpecial="!""#$%&'()*+,.-_/:;<=>?@[\]^`{|}~%"
     for i=1 to len(enStr)
       c=Mid(enStr,i,1)
       if c="%" then
         v=eval("&h"+Mid(enStr,i+1,2))
         if inStr(strSpecial,chr(v))>0 then
           deStr=deStr&chr(v)
           i=i+2
         else
           v=eval("&h"+ Mid(enStr,i+1,2) + Mid(enStr,i+4,2))
           deStr=deStr & chr(v)
           i=i+5
         end if
       else
         if c="+" then
           deStr=deStr&" "
         else
           deStr=deStr&c
         end if
       end if
     next
     URLDecode=deStr
End function
%><!--#Include File="../Code.asp"-->
