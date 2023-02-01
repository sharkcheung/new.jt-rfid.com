<!--#include file = "inc2.asp" -->
<%
if session("u_id")="" then
response.Redirect "../"
response.End
end if
%>
<%dim dingdan,action
action=request.QueryString("action")
dingdan=request.QueryString("dan")
select case action
case "save"
if request("zhuangtai")<>"" then
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select pro_paystatu from u_order where order_id='"&dingdan&"'",connn,1,3
	do while not rs.EOF
	old_zhuangtai=rs("pro_paystatu")
		rs("pro_paystatu")=request("zhuangtai")
		rs.Update
		rs.MoveNext
	loop
	rs.Close
	set rs=nothing
end if

if cint(request("zhuangtai"))=5 and old_zhuangtai<>5 then
	jifen=0
	ifhuyuanka=0
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select pro_num,bookid from u_order where order_id='"&dingdan&"'",connn,1,1
	while not rs.eof
		set rs2=server.CreateObject("adodb.recordset")
		rs2.Open "select bookid,yeshu from BJX_goods where bookid="&rs("bookid"),connn,1,1
		jifen=jifen+rs("pro_num")*rs2("yeshu")
		rs2.close
		set rs2=nothing
		rs.MoveNext
	wend
	rs.Close
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select jifen,reglx,vipdate from u_members where m_uid='"&session("u_id")&"'",connn,1,3
	rs("jifen")=rs("jifen")+jifen
	if ifhuyuanka=1 then 
		rs("reglx")=2
		if rs("vipdate")<>"" then 
		if rs("vipdate")<date then
		rs("vipdate")=date+365
		else
		rs("vipdate")=rs("vipdate")+365
		end if
		else
		rs("vipdate")=date+365
		end if
	end if
	rs.Update
	rs.Close
	set rs=nothing
	
	if ifhuyuanka=1 then 
		response.Write "<script language=javascript>alert('订单状态修改成功！您本次购物获得积分:"&jifen&"，你本次购买了会员卡，恭喜你现在已经成为本站的VIP用户！！');history.go(-1);</script>"
	else
		response.Write "<script language=javascript>alert('订单状态修改成功！您本次购物获得积分:"&jifen&"');history.go(-1);</script>"
	end if
else
	response.Write "<script language=javascript>alert('订单状态修改成功！');history.go(-1);</script>"
end if

case "del"
set rs=server.CreateObject("adodb.recordset")
rs.open "select username,order_id from u_order where order_id='"&dingdan&"' " ,connn,1,1
'先判断此订单是不是操作人的
if request.Cookies("bjx")("username")<>trim(rs("username")) then
response.Write "您无权删除此订单!"
response.End
end if
'删除前要返还库存及积分
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_order where  order_id='"&dingdan&"' and zhuangtai=1",connn,1,1
while not rs.eof
	set rs_s=server.CreateObject("adodb.recordset")
	rs_s.open "select * from BJX_goods where bookid="&rs("bookid"),connn,1,3
	rs_s("kucun")=rs_s("kucun")+rs("bookcount")
	rs_s("chengjiaocount")=rs_s("chengjiaocount")-rs("bookcount")
	rs_s.update
	rs_s.close
	set rs_s=nothing
rs.movenext
wend
rs.close

z_jifen=0
'set rs=server.CreateObject("adodb.recordset")
'rs.open "select * from u_order_jp where  dingdan='"&dingdan&"'",conn,1,1
'while not rs.eof
'z_jifen=z_jifen+rs("jifen")
'rs.movenext
'wend
'rs.close
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_members where m_uid='"&session("u_id")&"'",connn,1,3
rs("jifen")=rs("jifen")+z_jifen
rs.update
rs.close
set rs=nothing

'只能是订了未付款时删除
connn.execute "delete from u_order where order_id='"&dingdan&"' and pro_paystatu=1"
connn.execute "delete from u_order_jp where order_id='"&dingdan&"'"
response.Write "<script language=javascript>alert('订单删除成功！');window.close();</script>"

case "star"
'给订单评分
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_order where order_id='"&dingdan&"'",connn,1,3
rs("star")=request("star")
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('您的评分已成功提交！');history.go(-1);</script>"

case "pingjia"
'给订单一个评价
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_order where order_id='"&dingdan&"'",connn,1,3
rs("pingjia")=HTMLEncode(trim(request("pingjia")))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('您的评分已成功提交！');history.go(-1);</script>"
end select

%>