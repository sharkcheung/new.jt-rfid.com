<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../member/func2.asp"-->
<%dim action,Payid
Payid=request.QueryString("id")
if Payid<>"" then
if not isnumeric(Payid) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
action=request.QueryString("action")
set rs=server.CreateObject("adodb.recordset")
select case action
'////修改支付方式
case "zhifusave"
rs.open "select * from Iheeo_Pay where Payid="&Payid,conn,1,3
rs("PayName")=trim(request("PayName"))
rs("PayList")=request("PayList")
rs("PayKey")=request("PayKey")
rs("PayShopID")=trim(request("PayShopID"))
rs("PayShopKey")=trim(request("PayShopKey"))
rs("PayShopOther1")=trim(request("PayShopOther1"))
rs.update
rs.close
response.write "<script>alert('成功修改支付方式！');</script>"
response.End
'/////添加支付方式
case "zhifuadd"
rs.open "select * from Iheeo_Pay",conn,1,3
rs.addnew
rs("PayName")=trim(request("PayName"))
rs("PayList")=request("PayList")
rs("PayKey")=request("PayKey")
rs("PayShopID")=trim(request("PayShopID"))
rs("PayShopKey")=trim(request("PayShopKey"))
rs("PayShopOther1")=trim(request("PayShopOther1"))
rs.update
rs.close
response.write "<script>alert('成功添加支付方式！请刷新以后查看结果。');</script>"
response.End
'/////删除支付方式
case "zhifudel"
conn.execute "delete from Iheeo_Pay where Payid="&Payid
response.write "<script>alert('成功删除支付方式！');</script>"
end select
set rs=nothing
%>

<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
<td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
                                    
  <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select top 1 * from Iheeo_Pay where PayName='支付宝'",conn,1,1
		  j=rs.recordcount
		  if not rs.eof then%>
                                  <form name="form1" method="post" target="saveiframe" action="zhifu.asp?action=zhifusave&id=<%=rs("Payid")%>">
                              <tr> 
                                    <td align="right" width="100">支付方式：</td><td style="padding-left:10px;"  align="left"><%=trim(rs("PayName"))%>
									<input class="Input"   name="PayName" type="hidden" id="PayName" size="50" value=<%=trim(rs("PayName"))%>>
									</td>
								</tr>
                              <tr> 
                                    <td align="right">商 户ID：</td><td  align="left"><input class="zffs Input" name="PayShopID" type="text" id="PayShopID" size="50" value=<%=rs("PayShopID")%>></td>
								</tr>
                              <tr> 
                                    <td align="right">商户KEY：</td><td  align="left"><input class="zffs Input" name="PayShopKey" type="password" id="PayShopKey" size="50" value=<%=rs("PayShopKey")%>></td>
								</tr>
                              <tr> 
                                    <td align="right">商户账号：</td><td  align="left"><input class="zffs Input" name="PayShopOther1" type="text" id="PayShopOther1" size="50" value=<%=rs("PayShopOther1")%>></td>
								</tr>
                              <tr> 
                                    <td>&nbsp;</td><td align="left"><input style="margin-left:10px;" class="Button" type="submit" name="Submit2" value="修改" onclick="layer.closeAll();$('select').show();"></td>
								</tr>
                                    <!--<td  align="center">
									<input class="Input"   name="PayList" type="text" id="PayList" size="3" value=<%'=rs("PayList")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
                                    </td>
                                    <td  align="center">
									<input class="Input"   name="PayKey" type="text" id="PayKey" size="6" value=<%=rs("PayKey")%> readonly></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopID" type="text" id="PayShopID" size="16" value=<%=rs("PayShopID")%>></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopKey" type="password" id="PayShopKey" size="16" value=<%=rs("PayShopKey")%>></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopOther1" type="text" id="PayShopOther1" size="16" value=<%=rs("PayShopOther1")%>></td>
                                    <td  STYLE='PADDING-LEFT: 20px'>
									<input class="Button" type="submit" name="Submit2" value="修改" onclick="layer.closeAll();$('select').show();">
									&nbsp;<a onclick="layer.closeAll();$('select').show();return confirm('删除以后无法恢复！您确定要删除吗？');" target="saveiframe" href='zhifu.asp?action=zhifudel&amp;id=<%'=rs("Payid")%>'><font color="#FF0000">删除</font></a></td>-->
                                  </form>
                                  <%
                                  end if
		  rs.close
		  set rs=nothing%>						<!--<tr>
								<td colspan="7"  align="left">　添加支付方式</td></tr>
                                <tr> 
                                  <form name="form2" method="post" target="saveiframe" action="zhifu.asp?action=zhifuadd">
                                    <td  align="center">
									<input class="Input"   name="PayName" type="text" id="PayName" size="14">
                                    </td>
                                    <td  align="center">
									<input class="Input"   name="PayList" type="text" id="PayList" value=<%'=j+1%> size="3" onKeyPress= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
</td>
                                    <td  align="center">
									<input class="Input"   name="PayKey" type="text" id="PayKey" size="6"></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopID" type="text" id="PayShopID" size="16"></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopKey" type="password" id="PayShopKey" size="16"></td>
                                    <td  STYLE='PADDING-LEFT: 20px'>
<input type="submit" onclick="layer.closeAll();$('select').show();" class="Button" name="Submit32" value="添加">
</td>
</form>
</tr>-->
</table><iframe name="saveiframe" src="" height="1" width="1"></iframe>
<br><br>
<div align="center">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
        	<td align="right"><strong>支付方式说明</strong></td>
			<td align="left" class="ListTdTop">
			</td></tr>
		<tr>
			<!--<td align="center">货到付款</td>
			<td align="left">使用货到付款支付</td>
		</tr>
		<tr>
			<td align="center" >快钱支付</td>
			<td align="left">推荐使用支付方式 <a target="_blank" href="https://www.99bill.com/">申请地址&gt;&gt;</a></td>
		</tr>-->
		<tr>
			<td align="right" width="100">支付宝支付：</td>
			<td align="left"><a style="width:auto; line-height:21px; margin-left:10px;" title="点击查看申请流程" target="_blank" href="http://act.life.alipay.com/shopping/before/help/index.html">申请流程&gt;&gt;</a></td>
		</tr>
	</table>
</div>
</td>
</tr>
</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
   <input style="margin-left:113px;" type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
<script>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</script>