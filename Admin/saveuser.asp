<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Inc/Md5.asp"--><%
dim userid,action
action=request.QueryString("action")
userid=request.QueryString("id")
if userid="" then userid=request("userid")
select case action
case "save"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from u_members where id="&userid,conn,1,3
if trim(request("userpassword"))<>"" then rs("m_upass")=Md5(md5(trim(request("userpassword")),32),16)
rs("m_uname")=trim(request("userzhenshiname"))
rs("m_uemail")=trim(request("useremail"))
'rs("m_question")=trim(request("quesion"))
if trim(request("answer"))<>"" then rs("m_answer")=trim(request("answer"))
'rs("sfz")=trim(request("sfz"))
rs("m_usex")=trim(request("shousex"))
rs("m_uage")=trim(request("nianling"))
'rs("szsheng")=trim(request("hukouprovince"))
'rs("szshi")=trim(request("hukoucapital"))
'rs("szxian")=trim(request("hukoucity"))
rs("m_uaddress")=trim(request("shouhuodizhi"))
rs("m_utel")=trim(request("usertel"))
rs("m_umobile")=trim(request("usermobile"))
rs("m_uzip")=trim(request("youbian"))
rs("m_uQQ")=trim(request("qq"))
'rs("m_uWeb")=trim(request("homepage"))
rs("content")=trim(request("content"))
'if trim(request("vipdate"))<>"" then
'    rs("vipdate")=trim(request("vipdate"))
'end if

if trim(request("yucun"))<>"" then
rs("yucun")=trim(request("yucun"))
else
rs("yucun")=0
end if

'rs("reglx")=trim(request("reglx"))

rs.Update
rs.Close
set rs=nothing
response.Write "修改会员信息操作成功!"
case "del"
conn.execute "delete from u_members where id in ("&userid&") "
conn.execute "delete from u_order where u_id in ("&userid&")"
'conn.execute "delete from BJX_action_jp where userid in ("&userid&")"
'conn.execute "delete from BJX_history where userid in ("&userid&")"
'response.Redirect request.servervariables("http_referer")
response.Write "会员删除成功!"
end select
%>
