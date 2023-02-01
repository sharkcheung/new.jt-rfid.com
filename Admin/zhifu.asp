<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../member/func2.asp"-->
<%dim action,Payid,PayKey,PayName,PayList,PayShopID,PayShopKey,PayAccount,formName
'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call paysList() '支付方式列表
	Case 2
		Call payEditDo() '执行支付方式修改
	Case Else
		Response.Write("没有找到此功能项！")
End Select

sub payEditDo()
Payid=request("id")
PayKey=request("PayKey")
PayShopID=request("PayShopID")
PayShopKey=request("PayShopKey")
PayAccount=request("PayAccount")
set rs=server.CreateObject("adodb.recordset")
if PayKey=2 then
	if PayAccount="" then
		response.Write "支付宝账号不能为空!"
		response.end
	end if
end if
rs.open "select * from Iheeo_Pay where Payid="&Payid,conn,1,3
rs("PayShopID")=PayShopID
rs("PayShopKey")=PayShopKey
rs("PayInfo")=PayAccount
rs.update
rs.close
response.write "成功修改支付方式!"
'/////添加支付方式
'case "zhifuadd"
'rs.open "select * from Iheeo_Pay",conn,1,3
'rs.addnew
'rs("PayShopID")=PayShopID
'rs("PayShopKey")=PayShopKey
'rs("PayInfo")=PayAccount
'rs.update
'rs.close
'response.write "<script>alert('成功添加支付方式！请刷新以后查看结果。');<//script>"
'response.End
''/////删除支付方式
'case "zhifudel"
'conn.execute "delete from Iheeo_Pay where Payid="&Payid
'response.write "<script>alert('成功删除支付方式！');<//script>"
'end select
'set rs=nothing
end sub

sub paysList()
Session("NowPage")=FkFun.GetNowUrl()%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    支付设置
</div>
<div id="ListContent">
<table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" >
                                <tr >
                                  <td width="16%" align="center">网上支付方式</td>
                                  <td width="21%" align="center">商户ID</td>
                                  <td width="26%" align="center">商户KEY</td>
                                  <td width="20%" align="center">支付账号(支付宝必填项)</td>
                                  <td width="17%" align="center">操 作</td>
                                </tr>
                                    
  <%set rs=server.CreateObject("adodb.recordset")
		  rs.open "select * from Iheeo_Pay order by PayList",conn,1,1
		  j=rs.recordcount
		  dim i
		  i=0
		  do while not rs.eof
		  i=i+1%>
                              <tr> 
                                  <form name="form" method="post" >
                                    <td  align="center"><%=trim(rs("PayName"))%><input type="hidden" value="<%=trim(rs("PayKey"))%>" name="PayKey" id="PayKey<%=i%>"/></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopID" type="text" id="PayShopID<%=i%>" size="20" value="<%=trim(rs("PayShopID"))%>"></td>
                                    <td  align="center">
									<input class="Input"   name="PayShopKey" type="password" id="PayShopKey<%=i%>" size="22" value="<%=trim(rs("PayShopKey"))%>"></td>
                                    <td  align="center">
									<input class="Input"   name="PayAccount" type="text" id="PayAccount<%=i%>" size="16" value="<%=trim(rs("PayInfo"))%>"></td>
                                    <td  STYLE='PADDING-LEFT: 20px'>
									<input class="Button" type="button" name="Submit2" value="修改" onclick="DelIt('','zhifu.asp?Type=2&id=<%=rs("Payid")%>&PayKey=<%=rs("PayKey")%>&PayShopID='+$('#PayShopID<%=i%>').val()+'&PayShopKey='+$('#PayShopKey<%=i%>').val()+'&PayAccount='+$('#PayAccount<%=i%>').val(),'MainRight','<%=Session("NowPage")%>');">
									&nbsp;<!--<a onclick="$('#Boxs').hide();$('select').show();return confirm('删除以后无法恢复！您确定要删除吗？');" target="saveiframe" href='zhifu.asp?action=zhifudel&amp;id=<%'=rs("Payid")%>'><font color="#FF0000">删除</font></a>--></td>
                                  </form>
                                  <%rs.movenext
		  loop
		  rs.close
		  set rs=nothing%>
                                </tr>								<!--<tr>
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
<input type="submit" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="Submit32" value="添加">
</td>
</form>
</tr>-->
		<tr>
			<td align="left" class="ListTdTop" colspan="5">
			&nbsp; &nbsp;支付方式说明</td>
		</tr>
		<tr>
			<td width="16%" height="30">&nbsp; &nbsp;货到付款</td>
			<td width="84%" height="30" align="left" colspan="4">使用货到付款支付</td>
		</tr>
		<tr>
			<td height="30" >&nbsp; &nbsp;快钱支付</td>
			<td height="30" align="left" colspan="4">推荐使用支付方式 <a target="_blank" href="https://www.99bill.com/">申请地址&gt;&gt;</a></td>
		</tr>
		<tr>
			<td height="30" >&nbsp; &nbsp;支付宝支付</td>
			<td height="30" align="left" colspan="4"><a target="_blank" href="https://www.alipay.com">申请地址&gt;&gt;</a></td>
		</tr>
</table>
</div>
<div id="ListBottom">
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
<%end sub%>