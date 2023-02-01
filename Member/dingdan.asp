<!--#include file = "inc2.asp" -->
<!--#include file="func.asp"-->
<%
if session("u_id")="" then
response.Redirect "../"
response.End
end if
%>
<html><head><title><%=webname%>--订单详细资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/style.css" rel="stylesheet" type="text/css">
<style type="text/css">
   *{font-size:12px;}
   td{padding:4px;}
</style>
<body leftmargin="0" rightmargin="0" topmargin="0" marginwidth="0" marginheight="0" oncontextmenu="return false">
<%dim dingdan
dingdan=request.QueryString("dan")
oid=request.QueryString("oid")
'response.Write dingdan&"_"&oid
'response.end
sql="select * from u_order where id="&oid&""
set rs=connn.execute(sql)
if rs.eof and rs.bof then
response.write "<p align=center>此订单中有商品已被管理员删除，无法进行正确计算!<br>订单取消，请通知管理员或重新下订单!</p>"
response.End
end if
pay_statu=rs("pro_paystatu")
product_id=rs("p_id")
if product_id=0 then
a=split(rs("pro_ids"),"|")
for i=0 to ubound(a)-1
   b=split(a(i),",")
   for j=0 to ubound(b)-1
      ab=ab&p_price(b(j),"Fk_Product_Title")&" x "&b(1)
	  if i<>ubound(a)-1 then ab=ab&"+"
   next
next
product_list_show=ab
else
product_list_show=p_price(product_id,"Fk_Product_Title")&" x "&rs("pro_num")
end if
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> <td height="10"></td></tr>
<tr>
<td> 
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#cccccc" align="center">
                          <tr><td colspan="2">　<strong><font color="#ffffff">订购数量订单号为：<font color="#FF0000"><%=dingdan%></font>，详细资料如下：</font></strong></td>
                          </tr><tr bgcolor="#FFFFFF"> 
                            <td width="15%" align="right">订单状态：</td>
                            <td> 
                              <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
                                <tr> 
                                  <form name="form1" method="post" action="savedingdan.asp?dan=<%=dingdan%>&action=save">
                                    <td> 
                                      <%zhuang()%>
                                      <br>
					<%if pay_statu=5 then %><%else
					response.write "<font color=red><b>订单工作流程全部完成</b></font>"
					end if%>
                                    </td>
                                  </form>
                                </tr>
                              </table>
                            </td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">产品列表：</td>
                            <td> 
                              <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#cccccc">
                                <tr align="center"> 
                                  <td width="54%"><strong><font color="#ffffff">产品</font></strong></td>
                                  <td width="24%"><strong><font color="#ffffff">数量</font></strong></td>
                                  <td width="22%"><strong><font color="#ffffff">金额小计</font></strong></td>
                                </tr>
                                <%zongji=0%>
                                <tr bgcolor="#FFFFFF"> 
                                  <td height="22" rowspan="2" valign="top" style='PADDING-LEFT: 5px'><%=product_list_show%></td>
                                  <td> 
                                    <div align="center"><%=rs("pro_num")%></div>
                                  </td>
                                  <td> 
                                    <div align="center"><%=rs("pro_price")&"元"%></div>
                                  </td>
                                </tr>
		<%zongji=rs("pro_num")*rs("pro_price")+zongji
		feiyong=rs("pro_fee")%>
                                <tr bgcolor="#FFFFFF"> 
                                  <td colspan="4" height="22"> 
                                    <div align="right">订单总额：<%=zongji%>元＋费用：<%=feiyong%>元　　共计：<%=zongji+feiyong%>元 
                                      &nbsp;&nbsp;&nbsp;&nbsp;</div>
                                  </td>
                                </tr>
                              </table>
                            </td>
                          </tr><%if rs("pro_paystatu")=5 then %>
                          <%end if%>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">收货人姓名：</td>
                            <td style="PADDING-LEFT: 12px"><%=trim(rs("pro_contact"))%></td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">收货地址：</td>
                            <td style="PADDING-LEFT: 12px"><%=trim(rs("pro_add"))%></td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">邮编：</td>
                            <td style="PADDING-LEFT: 12px"><%=trim(rs("pro_post"))%></td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">联系电话：</td>
                            <td style="PADDING-LEFT: 12px"><%=trim(rs("pro_tel"))%>　</td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">支付方式：</td>
                            <td style="PADDING-LEFT: 12px"> 
                              <%select case int(rs("pro_paytype"))
		     case 99999
			 response.write "银行汇款"
			 case 100000
			 response.Write "货到付款"
		     case else
		  sql2="select * from Iheeo_Pay where PayKey="&int(rs("pro_paytype"))
          set rs2=connn.execute(sql2)
		  if rs2.eof and rs2.bof then
		  response.write "方式已被删除"
		  else
          response.Write trim(rs2("PayName"))
          end if
		  rs2.Close
          set rs2=nothing
         end select%>
                            </td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">您的留言：</td>
                            <td style="PADDING-LEFT: 12px"><%=trim(rs("pro_message"))%>　</td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                            <td align="right">下单日期：</td>
                            <td style="PADDING-LEFT: 12px"><%=rs("order_time")%>　</td>
                          </tr>
                          <tr bgcolor="#FFFFFF"> 
                          <td height="40" colspan="2" align="center"> 
<%if rs("pro_paystatu")=1 then%>
<input class="go-wenbenkuang" type="submit" name="submit" value="删除订单" onClick="if(confirm('您确定要删除吗?')) location.href='savedingdan.asp?action=del&dan=<%=dingdan%>';else return;">
<%end if%>
<input class=go-wenbenkuang onClick=javascript:window.close() type=reset value=关闭窗口 name=submit>
</td>
</tr>
</table>
</td>
</tr>
</table>
<%
sub zhuang()
   select case rs("pro_paytype")
      case 100000%>
	     <input name="checkbox" type="checkbox" DISABLED id="checkbox" value="checkbox" checked>未作任何处理<span style='font-family:Wingdings;'>à</span> 
         <input name="checkbox3" type="checkbox" DISABLED value="checkbox" checked>订单已确认<span style='font-family:Wingdings;'>à</span> 
         <input type="checkbox" name="checkbox4" value="checkbox" checked DISABLED>订单已完成【case 6】
      <%case else
       select case rs("pro_paystatu")
          case 0%>
		     订单已经成立 请您在有效期内尽快付款　付款方式查询<br>
			 <input type="hidden" name="zhuangtai" value="1"><input name="submit" type="image" src="images/ding_1.gif" border="0" value=" 确认已经付款 ">
          <%case 1%>
		     货款已经汇出 等待服务商查款　服务商联系方式<br><img border="0" src="images/ding_2.gif">
          <%case 2%>
		     款已经收到 等待服务商发货　服务商联系方式<br><img border="0" src="images/ding_3.gif">
          <%case 3%>
		     服务商已经发货<br>
			 <input type="hidden" name="zhuangtai" value=4><input name="submit" type="image" src="images/ding_4.gif" border="0" value=" 确认已经收货">
          <%case 4%>如果您收到货<br><img border="0" src="images/ding_5.gif">
          <%case 5
	   end select
	end select
end sub
rs.close
set rs=nothing%>
</body>
</html>