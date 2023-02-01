<!--#Include File="Include.asp"-->
<!--#Include File="../member/func2.asp"-->
<%
dim action,dingdan,username,oid,ors,p_id,pro_ids,p_rs,p_name,target,p_link,zongji,feiyong,frs,SongName,prs,old_zhuangtai
action=request.QueryString("action")
oid=request.QueryString("oid")
dingdan=request.QueryString("dan")
username=request.QueryString("username")
select case action
case "save"
if request("zhuangtai")<>"" then
'	set rs=server.CreateObject("adodb.recordset")
'	rs.Open "select pro_paytype from u_order where id="&oid&"",conn,1,3
'	do while not rs.EOF
'	old_zhuangtai=rs("pro_paytype")
'		rs("pro_paytype")=request("zhuangtai")
'		rs.Update
'		rs.MoveNext
'	loop
'	rs.Close
'	set rs=nothing
   conn.execute "update u_order set pro_paystatu="&request("zhuangtai")&" where id="&oid&" "
end if
'if request("zhuangtai")=4 then
'fhsj=now()
'conn.execute "update u_order set fhsj=date() where dingdan='"&dingdan&"' "
'end if

if cint(request("zhuangtai"))=5 and old_zhuangtai<>5 then
'	jifen=0
'	ifhuyuanka=0
'		set rs2=server.CreateObject("adodb.recordset")
'		rs2.Open "select vipid from BJX_sys",conn,1,1
'		vipid=rs2("vipid")
'		rs2.close
'		set rs2=nothing
'	set rs=server.CreateObject("adodb.recordset")
'	rs.Open "select bookcount,bookid from u_order where dingdan='"&dingdan&"'",conn,1,1
'	while not rs.eof
'		set rs2=server.CreateObject("adodb.recordset")
'		rs2.Open "select bookid,yeshu from BJX_goods where bookid="&rs("bookid"),conn,1,1
'		jifen=jifen+rs("bookcount")*rs2("yeshu")
'		rs2.close
'		set rs2=nothing
'		
'		if rs("bookid")=cint(vipid) then 
'			ifhuyuanka=1
'		end if
'		
'		rs.MoveNext
'	wend
'	rs.Close
'	'response.write ifhuyuanka&"'"&vipid
'	'response.end
'	set rs=server.CreateObject("adodb.recordset")
'	rs.Open "select jifen,reglx,vipdate from bjx_User where username='"&username&"'",conn,1,3
'	rs("jifen")=rs("jifen")+jifen
'	if ifhuyuanka=1 then 
'		rs("reglx")=2
'		if rs("vipdate")<>"" then 
'		if rs("vipdate")<date then
'		rs("vipdate")=date+365
'		else
'		rs("vipdate")=rs("vipdate")+365
'		end if
'		else
'		rs("vipdate")=date+365
'		end if
'	end if
'	rs.Update
'	rs.Close
'	set rs=nothing
'	
'	if ifhuyuanka=1 then 
'		response.Write "<script language=javascript>alert('&#1524;&#812;&#1976;&#307;&#633;&#891;&ucirc;:"&jifen&"&#15486;&#763;&#1329;&#1010;&#1150;&#938;&#1406;VIP&ucirc;');history.go(-1);<'/script>"
'	else
'		response.Write "<script language=javascript>alert('&#1524;&#812;&#1976;&#307;&#633;&#891;&ucirc;:"&jifen&"');history.go(-1);</'script>"
'	end if
else
	response.Write "<script language=javascript>alert('订单状态修改成功！');window.location.href='viewdingdan.asp?oid="&oid&"&username="&username&"&dan="&dingdan&"'</script>"
end if


case "del"
'删除订单
set rs=server.createobject("adodb.recordset")
rs.open "select pro_paystatu from u_order where id="&oid&" ",conn,1,1
if rs("pro_paystatu")>7 then
rs.close


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_order where id="&oid&"",conn,1,1
while not rs.eof
	set rs_s=server.CreateObject("adodb.recordset")
	rs_s.open "select * from BJX_goods where bookid="&rs("bookid"),conn,1,3
	rs_s("kucun")=rs_s("kucun")+rs("pro_num")
	rs_s("chengjiaocount")=rs_s("chengjiaocount")-rs("pro_num")
	rs_s.update
	rs_s.close
	set rs_s=nothing
rs.movenext
wend
rs.close

'z_jifen=0
'set rs=server.CreateObject("adodb.recordset")
'rs.open "select * from u_order_jp where  dingdan='"&dingdan&"'",conn,1,1
'while not rs.eof
'z_jifen=z_jifen+rs("jifen")
'rs.movenext
'wend
'rs.close
'set rs=server.CreateObject("adodb.recordset")
'rs.open "select * from u_members where id="&username&"",conn,1,3
'rs("jifen")=rs("jifen")+z_jifen
'rs.update
'rs.close
'set rs=nothing

else
rs.close
set rs=nothing
end if
conn.execute "delete from u_order where id="&oid&" "
response.Write "<script language=javascript>alert('订单已删除！');</script>"
'response.Write "<script language=javascript>alert('订单已删除！');window.location.href='editdingdan.asp';</script>"

end select

%>