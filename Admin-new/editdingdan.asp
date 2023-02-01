<!--#Include File="AdminCheck.asp"-->

<%dim zhuangtai,namekey,selectm,selectbookid,currentPage,totalPut,oos,orderid
Types=Clng(Request.QueryString("Type"))
Select Case Types
	Case 1
		Call orderList() '管理员列表
	Case 2
		Call orderEditForm() '修改用户表单
	Case 3
		Call orderEditDo() '执行修改用户
	Case 4
		Call orderDelDo() '执行删除管理员
	Case Else
		Call orderList() '管理员列表
End Select
sub orderDelDo()
'//删除订单
orderid=request("orderid")
if orderid<>"" then
	on error resume next 
	conn.execute "delete from u_order where id in ("&orderid&")"
	if err then
		response.Write "删除订单失败"
	else
		response.write "删除订单成功"
	end if
else
	response.write "未选择要删除的订单"
end if
end sub

sub orderList()
	Session("NowPage")=FkFun.GetNowUrl()
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
namekey=trim(request("namekey"))
zhuangtai=request("zhuangtai")
if zhuangtai="" then zhuangtai=999
%>

<div id="ListTop">
<div class="gnsztopbtn">
	<h3>订单管理</h3><select name="select" onChange="SetRContent('MainRight','editdingdan.asp?zhuangtai='+this.options[this.selectedIndex].value);" ><base target=Right> 
	<option value="999"  <%if zhuangtai=999 then response.Write "selected" else response.Write "" end if%>>全部订单状态</option>
	<option value="0" <%if zhuangtai=0 then response.Write "selected" else response.Write "" end if%>>未作任何处理</option>
	<option value="1" <%if zhuangtai=1 then response.Write "selected" else response.Write "" end if%>>用户已经划出款</option>
	<option value="2" <%if zhuangtai=2 then response.Write "selected" else response.Write "" end if%>>服务商已经收到款</option>
	<option value="3" <%if zhuangtai=3 then response.Write "selected" else response.Write "" end if%>>服务商已经发货</option>
	<option value="4" <%if zhuangtai=4 then response.Write "selected" else response.Write "" end if%>>用户已经收到货</option>
	</select>
    <a class="no3" href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新订单</a>
    <a class="zfsz" href="javascript:void(0);" onclick="SetRContent('MainRight','zhifu.asp?Type=1');return false">支付设置</a>
    <a class="shdz" href="javascript:void(0);" onclick="SetRContent('MainRight','songhuo.asp?Type=1');return false">送货设置</a>
</div>

    
</div>

<div id="ListContent">
<form name="form1" method="post" action="">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr><td valign="top"> 
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="tableBorder">
	<tr > 
	<th width="21%" align="center" class="ListTdTop">订单号</th>
	<th width="15%" align="center" class="ListTdTop">下单会员</th>
	<th width="24%" align="center" class="ListTdTop">订货人姓名</th>
	<th width="13%" align="center" class="ListTdTop">付款方式</th>
	<th width="15%" align="center" class="ListTdTop">订单状态</th>
	<th width="12%" align="center" class="ListTdTop">操作</th>
	</tr>
<%set Rs=server.CreateObject("adodb.recordset")
   'if namekey="" then	   '//按状态查询
      if zhuangtai="" or zhuangtai=999 then
         Rs.open "select * from u_order order by order_time desc",conn,1,1
      else
         Rs.open "select * from u_order where pro_paystatu="&zhuangtai&" order by order_time",conn,1,1
      end if
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
		  %><tr> 
          <td align="left" style="padding-left:10px;"><input name="orderid" type="checkbox" value="<%=rs("id")%>"> <a href="javascript:void(0);" onclick="ShowBox('viewdingdan.asp?oid=<%=rs("id")%>&dan=<%=rs("order_id")%>&username=<%=trim(rs("u_id"))%>');"><%=trim(rs("order_id"))%></a></td>
          <td align="center"><%if rs("u_id")<>0 then 
		  set oos=conn.execute("select m_uid from u_members where id="&rs("u_id")&"")
		  response.Write oos("m_uid")
		  oos.close
		  set oos=nothing
		  else
		  response.Write "非会员用户"
		  end if%></td>
          <td align="center"><%=trim(rs("pro_contact"))%><input type="hidden" name="pid" value="<%=rs("p_id")%>"></td>
          <td align="center">
              <%dim rs2
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
          <td align="center">
              <%
		  select case rs("pro_paystatu")
	case "0"
	response.write "<span class=""pay_css1"">未作任何处理</span>"
	case "1"
	response.write "<span class=""pay_css2"">用户已付款</span>"
	case "2"
	response.write "<span class=""pay_css3"">服务商已经收到款</span>"
	case "3"
	response.write "<span class=""pay_css4"">服务商已经发货</span>"
	case "4"
	response.write "<span class=""pay_css5"">用户已经收到货</span>"
	case "5"
	response.write "<span class=""pay_css6"">用户已经完成评价</span>"
	end select%>
	</td>
    <td align="center">
	<input type="button" onclick="ShowBox('viewdingdan.asp?oid=<%=rs("id")%>&dan=<%=rs("order_id")%>&username=<%=trim(rs("u_id"))%>','查看/处理订单');" class="Button" name="button" id="button" value="详 细" />
    <input style="display:none" type="checkbox" name="selectbookid" value="<%=rs("id")%>"></td>
	</tr>
        <%
			Rs.MoveNext
			i=i+1
		Wend%>
	<tr > 
<td style=" padding-left:10px;" height="30" colspan="6" align="left"> 
<input style="vertical-align:middle" type="checkbox" name="checkbox" value="Check All" onClick="var checkboxs=document.getElementsByName('orderid');for (var i=0;i<checkboxs.length;i++) {var e=checkboxs[i];e.checked=!e.checked;}">&nbsp;&nbsp;全选
&nbsp;&nbsp;<input style="vertical-align:middle" type="button" class="Button" name="Submit" value="删 除" onClick="var str='';$('input[name=orderid]').each(function(){if(this.checked){if(str==''){str=$(this).val();}else{str+=','+$(this).val()}}});DelIt('您确认要删除，此操作不可逆！','editdingdan.asp?Type=4&orderid='+str,'MainRight','<%=Session("NowPage")%>');">
</td>
</tr>
        <tr>
            <td height="30" colspan="6" style="text-align:center; border-bottom:0;">&nbsp;<%Call FKFun.ShowPageCode("editdingdan.asp?Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
		<%
	Else
%>

<%
	End If
	Rs.Close
%>
      </table>
                              
</td>
</tr>
</table>
</form>
</div>

<!--<table width="98%"border="1" align="center" cellpadding="5" cellspacing="0" bordercolor="#B6CDD5"class="tableBorder">
                          <tr> 
                            <td align="center" class="t2" >定单查询</td></tr>
                          <tr> 
                            <td height="50" > 
                              <table width="80%" border="0" align="center" cellpadding="1" cellspacing="1">
                                <tr> 
                                  <form name="form1" method="post" action="editdingdan.asp">
                                    <td align="center">按下单用户查询 
                                        <input name="namekey" type="text" id="namekey" value="请输入用户名" size="14" onFocus="this.value=''">
                                        &nbsp; 
                                        <select name="zhuangtai" id="zhuangtai">
                                          <option value="" selected>--选择查询状态--</option>
                                          <option value="" >全部订单状态</option>
                                          <option value="0" >未作任何处理</option>
                                          <option value="1" >用户已经划出款</option>
                                          <option value="2" >服务商已经收到款</option>
                                          <option value="3" >服务商已经发货</option>
                                          <option value="4" >用户已经收到货</option>
                                        </select>
                                        &nbsp; 
                                        <input type="submit" name="Submit" value="查 询">
                                    </td>
                                  </form>
                                </tr>
                              </table>
                            </td>
                          </tr>
</table>-->
<div id="ListBottom">

</div>
<%end sub%>