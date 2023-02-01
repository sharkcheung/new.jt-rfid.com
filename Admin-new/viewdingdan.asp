<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../member/func2.asp"-->
<%
'dim dingdan,username,oid,ors,p_id,pro_ids,p_rs,p_name,target,p_link,zongji,feiyong,frs,SongName,prs
dim action,dingdan,username,oid,ors,p_id,pro_ids,p_rs,p_name,target,p_link,zongji,feiyong,frs,SongName,prs,old_zhuangtai,iiiii,jjjjj,bbbbb

action=request.QueryString("action")
if action="" then

oid=request.QueryString("oid")
dingdan=request.QueryString("dan")
username=request.QueryString("username")
if InStr(username,"'")>0 then
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
set ors=conn.execute("select * from u_order where id="&oid&"")
if not ors.eof then
   p_id=ors("p_id")
   pro_ids=ors("pro_ids")
end if
ors.close
set ors=nothing
if p_id=0 then
	if instr(pro_ids,"|")>0 then
		p_id=split(pro_ids,"|")
		for iiiii=0 to ubound(p_id)-1
		   bbbbb=split(p_id(iiiii),",")
		   for j=0 to ubound(bbbbb)-1
			  p_name=p_name&p_price(bbbbb(j),"Fk_Product_Title")&" x "&bbbbb(1)
			  if iiiii<>ubound(p_id)-1 then p_name=p_name&"+"
		   next
		next
	else
		p_name=""
	end if
   p_link="#"
   target="_self"
else
   p_name=p_price(p_id,"Fk_Product_Title")
   target="_blank"
   p_link="../product.asp?bookid="&server.URLEncode(Encrypt(p_id))&""
end if
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_order where u_id="&username&" and id="&oid&" ",conn,1,1
if rs.eof and rs.bof then
response.write "<p align=center><font color=red>此订单中有商品已被管理员删除，无法进行正确计算。<br>订单操作取消，请手动删除此订单！</font></p>"
response.write "<input type=button name=Submit3 value=删除订单 onClick=""location.href='savedingdan.asp?action=del&id="&oid&"&username="&username&"'""> </center>"
response.End
end if
'shjiaid=rs("shjiaid")
%>

<form id="SystemSet" name="SystemSet"  method="post" action="viewdingdan.asp?oid=<%=oid%>&dan=<%=dingdan%>&action=save&username=<%=username%>" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%"border="0" align="center" cellpadding="0" cellspacing="0" class="tableBorder">
                          <tr > 
                            <td colspan="2" align="center" style="background-color:#daf1ff; padding-left:10px;"> 
                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <th width="89%" align="left">
									订单号为：<b><%=dingdan%></b> ，详细资料如下：
                                    </th>
                                    <th width="11%" align="center">
									<input class="Button" type="button" name="Submit4" value="打 印" onClick="javascript:window.print()">
                                    </th>
                                  </tr>
                                </table>
                            </td>
                          </tr>
                          <tr> 
                            <td width="12%"  style='PADDING-LEFT: 10px'>
							订单状态：</td>
                            <td width="88%" style="border-bottom:0;"> 
                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                    <td > 
                                      <%zhuang()%><br>
									  &nbsp;<font color="#FF0000">这里只有用户付款后，管理员才可以进行继续操作！</font></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                          <tr> 
                            <td width="12%"  style='PADDING-LEFT: 10px; border-right:1px dashed #ceecff'>
							商品列表：</td>
                            <td style="border-bottom:0;"> 
                              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
                                <tr> 
                                  <td width="65%" align="center"  class="ListTdTop">商品名称</td>
                                  <td width="13%" align="center"  class="ListTdTop"><div align="center">
									  订购数量</div></td>
                                  <td width="22%" align="center"  class="ListTdTop"><div align="center">
									  金额小计</div></td>
                                </tr>
                                <tr> 
                                  <td  style='PADDING-LEFT: 5px'><%=p_name%></td>
                                  <td > 
                                    <div align="center"><%=rs("pro_num")%></div>
                                  </td>
                                  <td > 
                                    <div align="center"><%=rs("pro_price")&"元"%></div>
                                  </td>
                                </tr>
                                <%zongji=rs("pro_price")
								feiyong=rs("pro_fee_type")
								set frs=conn.execute("select SongFei,SongName from Iheeo_Delivery where SongKey="&feiyong&"")
								feiyong=frs("SongFei")
								SongName=frs("SongName")
								frs.close
								set frs=nothing%>
                                <tr> 
                                  <td  class="ListTdTop" colspan="3" > 
                                    <div align="right">订单总额：<%=zongji%>元＋费用：<%=feiyong%>元　　共计：<%=zongji+feiyong%>元 
										&nbsp;&nbsp;&nbsp;&nbsp;</div>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr><!--
                          <tr> 
                            <td width="12%"  style='PADDING-LEFT: 10px'>
							订单星级：</td>
                            <td  style='PADDING-LEFT: 10px'><img src="../images/level<%'=rs("pro_star")%>.gif"> 
                            　</td>
                          </tr>-->
                          <%'set snsn=server.CreateObject("adodb.recordset")
	'snsn.open "select * from u_order_jp where username="&username&" and dingdan='"&dingdan&"'",conn,1,1
	'if snsn.recordcount>0 then%>
<!--                          <tr bgcolor="#6a7f9a"> 
                            <td width="13%" style='PADDING-LEFT: 10px' >奖品列表：</td>
                            <td > 
                              <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" >
                                <tr> 
                                <td align="center">奖品名称</td>
                                <td align="center">所用积分</td>
                                </tr>
	<%
	'while not snsn.eof%>
                                <tr> 
                                  <td > 
                                    <div align="center">
                                      <%
	'set snsn1=server.CreateObject("adodb.recordset")
'	snsn1.open "select * from BJX_jiangpin where bookid="&snsn("bookid"),conn,1,1
'	if snsn1.recordcount=1 then
'	response.write snsn1("bookname")
'	end if
'	snsn1.close
'	set snsn1=nothing%>
                                    </div></td>
                                  <td align="center" ><%'=snsn("jifen")%></td>
                                </tr>
	<%
	'snsn.movenext
	'wend%>
                              </table>
                            </td>
                          </tr>-->
                          <%'end if
'snsn.close
'set snsn=nothing%>
                          <tr> 
                            <td  style='PADDING-LEFT: 10px'>
							收货人姓名：</td>
                            <td  style='PADDING-LEFT: 10px'><%=trim(rs("pro_contact"))%></td>
                          </tr>
                          <tr> 
                            <td  style='PADDING-LEFT: 10px'>
							收货地址：</td>
                            <td  style='PADDING-LEFT: 10px'><%=trim(rs("pro_add"))%></td>
                          </tr>
                          <tr> 
                            <td width="12%"  style='PADDING-LEFT: 10px'>
							送货方式：</td>
                            <td  style='PADDING-LEFT: 10px'>
<%=SongName%>
                            　</td>
                          </tr>
                          <tr> 
                            <td  style='PADDING-LEFT: 10px'>邮 
							编：</td>
                            <td  style='PADDING-LEFT: 10px'><%=trim(rs("pro_post"))%></td>
                          </tr>
                          <tr> 
                            <td  style='PADDING-LEFT: 10px'>
							联系电话：</td>
                            <td  style='PADDING-LEFT: 10px'><%=trim(rs("pro_tel"))%></td>
                          </tr>
                          <tr> 
                            <td  style='PADDING-LEFT: 10px'>
							支付方式：</td>
                            <td  style='PADDING-LEFT: 10px'><%dim rs2
          '///支付方式
		  select case int(rs("pro_paytype"))
		     case 99999
			 response.write "银行汇款"
			 case 100000
			 response.Write "货到付款"
		     case else
          set rs2=server.CreateObject("adodb.recordset")
          rs2.open "select * from Iheeo_Pay where PayKey="&int(rs("pro_paytype")),conn,1,1
		  if rs2.eof and rs2.bof then
		  response.write "方式已被删除"
		  else
          response.Write trim(rs2("PayName"))
          end if
		  rs2.Close
          set rs2=nothing
         end select
          %>
</td>
                          </tr>
                          <tr> 
                            <td  style='PADDING-LEFT: 10px'>
							用户留言：</td>
                            <td  style='PADDING-LEFT: 10px'><%=trim(rs("pro_message"))%></td>
                          </tr>
                          <tr> 
                            <td height="20"  style='PADDING-LEFT: 10px'>
							下单日期：</td>
                            <td height="20"  style='PADDING-LEFT: 10px'><%=rs("order_time")%></td>
                          </tr>
                          <tr style="display:none;">
						  <td height="30"  style='PADDING-LEFT: 10px'>
						  　</td>
                            <td  style='PADDING-LEFT: 10px'>
							<input type="button" name="Submit3" class="Button" value="删除订单" onClick="if(confirm('您确定要删除吗?')) location.href='savedingdan.asp?action=del&oid=<%=oid%>&username=<%=username%>';else return;">
                            </td>
                          </tr>
</table>
<%sub zhuang()
select case rs("pro_paystatu")
case 0%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理 
→ 
<%if rs("pro_paytype")<>100000 then %>
<input name="checkbox" type="checkbox" id="checkbox" value="checkbox" DISABLED>用户已付款 
→ 
<input type="checkbox" name="zhuangtai" value="2">服务商已经收到款 → 
<input type="checkbox" name="checkbox3" value="checkbox" DISABLED>服务商已经发货 → 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>用户已经收到货 
<%else%>
<input name="zhuangtai" type="checkbox" id="zhuangtai" value="2">订单已确认 → 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>订单已完成 
<%end if%>
<%case 1%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理 
→ 
<input name="checkbox" type="checkbox" id="checkbox" value="adf" checked DISABLED>用户已付款 
→ 
<input type="checkbox" name="zhuangtai" value="2" >服务商已经收到款 → 
<input type="checkbox" name="checkbox" value="checkbox" DISABLED>服务商已经发货 → 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>用户已经收到货
<%case 2%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理 
→ 
<%if rs("pro_paytype")<>100000 then %>
<input name="checkbox" type="checkbox" id="checkbox" value="checkbox" checked DISABLED>用户已付款 
→ 
<input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>服务商已经收到款 
→ 
<input type="checkbox" name="zhuangtai" value="3" >服务商已经发货 → 
<input type="checkbox" name="checkbox4" value="checkbox" DISABLED>用户已经收到货
<%else%>
<input name="checkbox" type="checkbox" DISABLED value="checkbox" checked>订单已确认 → 
<input type="checkbox" name="zhuangtai" id="checkbox" value="4">订单已完成 
<%end if%>
<%case 3%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理 
→ 
<input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>用户已付款 
→ 
<input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>服务商已经收到款 
→ 
<input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>服务商已经发货 
→ 
<input type="checkbox" name="checkbox" value="checkbox" DISABLED>用户已经收到货
<%case 4%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理 
→ 
<%if rs("pro_paytype")<>100000 then %>
<input name="checkbox" type="checkbox" id="checkbox" value="2" checked DISABLED>用户付款 
→ 
<input type="checkbox" name="checkbox2" value="checkbox" checked DISABLED>服务商已经收到款 
→ 
<input type="checkbox" name="checkbox3" value="checkbox" checked DISABLED>服务商已经发货 
→ 
<input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>用户已经收到货 
→ 
<input type="checkbox" name="zhuangtai" id="checkbox" value="5">完成评价
<%else%>
<input name="zhuangtai" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>订单已完成 
→ 
<input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>订单已完成 
→ 
<input type="checkbox" name="zhuangtai" id="checkbox" value="5">完成评价
<%end if%>
<%case 5%>
<input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理 
→ 
<input name="zhuangtai" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>订单已完成 
→ 
<input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>订单已完成 
<%
end select
end sub%>

</div>
<div id="BoxBottom" style="width:93%;" class="tcbtm">
 <input type="submit" onclick="layer.closeAll();$('select').show();Sends('SystemSet','viewdingdan.asp?action=del&oid=<%=oid%>&username=<%=username%>',0,'',0,0,'','');" class="Button" name="button" id="button" value="删除订单" />
 <input type="submit" onclick="Sends('SystemSet','viewdingdan.asp?oid=<%=oid%>&dan=<%=dingdan%>&action=save&username=<%=username%>',0,'',0,0,'','');" class="Button" name="button" id="button" value="修改订单" />
  <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
else
%>
<%
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
	Response.Write ("订单状态修改成功！")
end if


case "del"
'删除订单
set rs=server.createobject("adodb.recordset")
rs.open "select pro_paystatu from u_order where id="&oid&"",conn,1,1
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

response.Write "订单已删除！请关闭该窗口后刷新订单列表！"
'response.Write "<script language=javascript>alert('订单已删除！');window.location.href='editdingdan.asp';</script>"

end select

%>
<%end if%>