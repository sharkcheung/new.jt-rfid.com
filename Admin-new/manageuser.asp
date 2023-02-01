<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../member/func2.asp"-->
<%
dim zhuangtai,namekey,selectm,userids,currentPage,totalPut,oos

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call UserList() '用户列表
	Case 2
		Call UserEdit() '用户列表
	Case 3
		Call UserEditDo() '用户列表
	Case 4
		Call ListDelDo() '执行批量删除用户
	Case Else
		Call UserList() '用户列表
End Select

sub UserEditDo()
	Id=Trim(Request.Form("Id"))
	set rs=server.CreateObject("adodb.recordset")
	rs.Open "select * from u_members where id="&Id,conn,1,3
	if trim(request("userpassword"))<>"" then rs("m_upass")=Md5(md5(trim(request("userpassword")),32),16)
	rs("m_uname")=trim(request("userzhenshiname"))
	rs("m_uemail")=trim(request("useremail"))
	'rs("m_question")=trim(request("quesion"))
	if trim(request("answer"))<>"" then rs("m_answer")=trim(request("answer"))
	'rs("sfz")=trim(request("sfz"))
	rs("m_usex")=trim(request("shousex"))
	rs("m_uage")=trim(request("nianling"))
	rs("szsheng")=trim(request("hukouprovince"))
	rs("szshi")=trim(request("hukoucapital"))
	rs("szxian")=trim(request("hukoucity"))
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
	response.Write "会员信息修改成功!"
end Sub

sub ListDelDo()
	dim userid
	userid=Request("ListId")
	if userid<>"" Then
		
conn.execute "delete from u_order_product where o_id in (select order_id from u_order where u_id in ("&userid&"))"
conn.execute "delete from u_members where id in ("&userid&") "
conn.execute "delete from u_order where u_id in ("&userid&")"
conn.execute "delete from u_shipaddress where a_uid in ("&userid&")"
conn.execute "delete from u_shopcart where u_id in ("&userid&")"
'conn.execute "delete from BJX_action_jp where userid in ("&userid&")"
'conn.execute "delete from BJX_history where userid in ("&userid&")"
'response.Redirect request.servervariables("http_referer")
response.Write "会员删除成功！"
else
response.Write "请选择要删除会员"
end If
end Sub
Sub UserList()
	Session("NowPage")=FkFun.GetNowUrl()
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
%>

<div id="ListContent">
<div class="gnsztopbtn">
	<h3>会员管理</h3>
    <a class="lxsz" href="javascript:void(0);" onclick="SetRContent('MainRight','manageuser.asp?Type=1');return false">用户列表</a>
</div>
<form name="DelList" method="post" id="DelList" action="manageuser.asp?Type=4" onsubmit="return false;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
<tr  align="center">
<th class="ListTdTop"> 账号</th>
<th class="ListTdTop"> 姓名</th>
<th class="ListTdTop"> 注册时间</th>
<th class="ListTdTop">联系方式</th>
<th class="ListTdTop">用户数</th>
<th class="ListTdTop"> 登陆次数</th>
<th class="ListTdTop"> 选 择</th>
</tr>
<%
		 set rs=server.CreateObject("adodb.recordset")
		   rs.open "select * from u_members order by m_reg_time desc",conn,1,1
		  
      if not rs.EOF Then
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1%>
<tr >
<td style="PADDING-LEFT: 10px;text-align:left;"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("id")%>" id="List<%=Rs("id")%>" />　<a style="width:auto; vertical-align:middle; line-height:21px;" href="javascript:void(0);" onclick="ShowBox('manageuser.asp?Type=2&id=<%=rs("id")%>','会员详细资料','1000px','500px');return false;"><%=trim(rs("m_uid"))%></a></td>
<td style="PADDING-LEFT: 10px;text-align:center;"><%=trim(rs("m_uname"))%></td>
<td style="PADDING-LEFT: 10px;text-align:center;"><%=rs("m_reg_time")%></td>
<td style="PADDING-LEFT: 10px;text-align:center;"><%=rs("m_utel")&"  "&rs("m_umobile")%></td>
<td style="PADDING-LEFT: 10px;text-align:center;"><%dim rs2
			set rs2=conn.execute("select count(id) from u_order where u_id="&rs("id"))
			response.write rs2(0)
			rs2.close
			set rs2=nothing
			%>
</td>
<td align="center"><%=rs("m_login_count")%> 次</td>
<td align="center"><input type="button" onclick="ShowBox('manageuser.asp?Type=2&id=<%=rs("id")%>','会员详细资料','1000px','500px');" class="Button" name="button" id="button" value="详 细" /></td>
</tr>
        <%
			Rs.MoveNext
			i=i+1
		Wend
%>
	<tr > 
		<td colspan="7" style="PADDING-LEFT: 10px;text-align:left;"><input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="vertical-align:middle;">&nbsp;&nbsp;<label for="chkall" style="vertical-align:middle;">全选</label>
    	<input type="submit" class="Button" name="Submit" style="vertical-align: middle; margin-left:10px;" value="删 除" onClick="DelIt('确定要删除选中的用户吗？','manageuser.asp?Type=4&ListId='+GetCheckbox(),'MainRight','<%=Session("NowPage")%>');"></td>
    </tr>
    <tr > 
    	<td height="30" colspan="7" align="center"><%Call FKFun.ShowPageCode("manageuser.asp?Type=1&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
    </tr>
		<%
else%>

        <tr>
            <td height="25" colspan="7" align="center">暂无记录</td>
        </tr>
<%end if
rs.close%>
</table>
</form>
</div>
<div id="ListBottom">
</div>
<%End Sub
	
sub UserEdit()
dim resume_hukouprovinceid,resume_hukoucapitalid,resume_hukoucityid,province_name,capital_name,city_name,workadd,mystring,length,Hwasrs,province,Hwassql
Id=Clng(Request.QueryString("Id"))
if Id<>"" then
if not isnumeric(Id) then 
response.end
end if
end if
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from u_members where id="&Id ,conn,1,1
		resume_hukouprovinceid=rs("szSheng")
		resume_hukoucapitalid=rs("szShi")
		resume_hukoucityid=rs("szXian")
	%>
<form name="UserEdit" id="UserEdit" method="post" action="manageuser.asp?types=3">

<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" >
<tr> 
                                   <td >账 号：</td>
                                   <td>
                                   <font style="font-weight:bold; letter-spacing:1px; color:#FF0000; font-size:14px;"><%=trim(rs("m_uid"))%></font>
                                   <!--<font color=#FF0000>用户类型：</font>
									<select name="reglx">
									<option value="1" <if rs("reglx")=1 then%>selected<end if%>>普通会员</option>
									<option value="2" <if rs("reglx")=2 then%>selected<end if%>>VIP 会员</option>
									</select>
									<font color=#FF0000>VIP期限：</font>
									<input class="Input"  type="text" name="vipdate" size="10" value="<=rs("vipdate")%>">
									格式：2004-02-22</td>-->
   									</td>
                                   <td>
 									密 码：
                                   <td   >
									<input class="Input"  name="userpassword" type="text" id="userpassword">
									<font color=#FF0000>不改密码请留空!
                                  <tr> 
                                    <td >姓 名：</td>
                                    <td ><input class="Input"  name="userzhenshiname" type="text" id="userzhenshiname" size="20" value="<%=trim(rs("m_uname"))%>"></td>
                                    <td >性 别：</td>
                                    <td >
									<input class="Input"  type="radio" name="shousex" value=0 <%if rs("m_usex")=0 then%>checked<%end if%> checked>男
									<input class="Input"  type="radio" name="shousex" value=1 <%if rs("m_usex")=1 then%>checked<%end if%>>女　　　年 龄：<input style=" width:105px;" class="Input"  name=nianling type=text value="<%=trim(rs("m_uage"))%>" size="3" maxlength="2" onKeyPress="event.returnValue=IsDigit();">
                                    </td>
                                    </tr>
                                    <tr> 
                                    <td >密码提问：</td>
                                    <td >
                                    <select class="Input"  name="quesion" id="quesion">
                                    <option value="<%=trim(rs("m_question"))%>" <%if rs("m_question")=0 then response.write "selected"%>>我身份证最后6位数</option>
                                    <option value="<%=trim(rs("m_question"))%>" <%if rs("m_question")=1 then response.write "selected"%>>我父亲的名字</option>
                                    <option value="<%=trim(rs("m_question"))%>" <%if rs("m_question")=2 then response.write "selected"%>>我母亲的名字</option>
                                    <option value="<%=trim(rs("m_question"))%>" <%if rs("m_question")=3 then response.write "selected"%>>我就读的小学校名</option>
                                    <option value="<%=trim(rs("m_question"))%>" <%if rs("m_question")=4 then response.write "selected"%>>我最喜欢的颜色</option>
                                    </select>
									</td>
                                    <td >密码答案：</td>
                                    <td ><input class="Input" size="20" name="answer" type="text" id="answer" value="<%=rs("m_answer")%>"></td>
                                    </tr>
                                
                                  <tr> 
                                    <td >手机号码</td>
                                    <td >
                                      <input class="Input"  name="usermobile" type="text" id="usermobile" size="25" value="<%=trim(rs("m_umobile"))%>"></td>
                                    <td >
									电话号码：</td>
                                    <td >
                                      <input class="Input"  name="usertel" type="text" id="usertel" size="20" value="<%=trim(rs("m_utel"))%>"></td>
                                  </tr><tr> 
                                    <td >Email：</td>
                                    <td >
									<input class="Input" size="25" name="useremail" type="text" id="useremail" value="<%=trim(rs("m_uemail"))%>"></td>
                                    <td >
																		Q Q：</td>
                                    <td >
                                      <input class="Input"  name=QQ type=text value="<%=trim(rs("m_uQQ"))%>" size="20" maxlength="12"></td></tr>
                                  <tr> 
                                    <td >省/市：</td>
                                    <td ><select   style="width:80px" class="Input" name="hukouprovince" size="1" id="select5" onChange="changeProvince(document.UserEdit.hukouprovince.options[document.UserEdit.hukouprovince.selectedIndex].value)">
		<%if resume_hukouprovinceid<>"" then%>
		<option value="<%=resume_hukouprovinceid%>"><%=Hireworkadds(resume_hukouprovinceid)%></option>
		<%else%>
		<option value="">选择省</option>
		<%end if%>
		</select>
						<select class="Input" name="hukoucapital"   style="width:80px" onChange="changeCity(document.UserEdit.hukoucapital.options[document.UserEdit.hukoucapital.selectedIndex].value)">
						  <%if resume_hukoucapitalid<>"" then%>
						  <option value="<%=resume_hukoucapitalid%>"><%=Hireworkadds(resume_hukoucapitalid)%></option>
						  <%else%>
						  <option value="">选择市</option>
						  <%end if%>
	                      </select>
		                  <select class="Input" name="hukoucity"  style="width:80px">
		                    <%if resume_hukoucityid<>"" then%>
		                    <option value="<%=resume_hukoucityid%>"><%=Hireworkadds(resume_hukoucityid)%></option>
		                    <%else%>
		                    <option value="">选择区</option>
		                    <%end if%>
                          </select>
                                    </td>
                                    <td > 
                                      										邮 编：</td>
                                    <td > 
                                      <input class="Input"  name="youbian" type="text" id="youbian" size="20" value="<%=rs("m_uzip")%>" maxlength=6 onKeyPress="event.returnValue=IsDigit();"></td>
                                  </tr>
                                  
                                  <tr> 
                                    <td >地 址：</td>
                                    <td colspan="3" ><input class="Input"  name="shouhuodizhi" type="text" id="shouhuodizhi" size="80" value="<%=trim(rs("m_uaddress"))%>"></td>
                                  </tr>
                                  
                                  <tr style="display:none"> 
                                    <td  style="PADDING-LEFT: 8px; height: 30px;" >身份证号码：</td>
                                    <td  style="PADDING-LEFT: 8px; height: 30px;" colspan="3" >
                                      <input class="Input"  name=sfz type=text value="<%=trim(rs("m_sfz"))%>" maxlength="18" onKeyPress="event.returnValue=IsDigit();"></td>
                                  </tr>
                                  <tr style="display:none"> 
                                    <td >个人主页：</td>
                                    <td colspan="3" > 
                                      <input class="Input"  name=homepage type=text value="<%=trim(rs("m_uWeb"))%>" size="40">
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td >简 介：</td>
                                    <td colspan="3" > 
                                      <textarea style="border: 1px solid #ccc; margin: 5px 0 5px 10px; width: 677px;" name="content" rows="3" cols="60"><%=trim(rs("content"))%></textarea>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td >注 册：</td>
                                    <td style="padding-left:10px"><%=rs("m_reg_time")%></td>
                                    <td >最后登陆：</td>
                                    <td  style="PADDING-LEFT:10px">&nbsp;<%=rs("m_last_logintime")%></td>
                                  </tr>
                                  <tr style="display:none">
                                    <td >购物积分</td>
                                    <td colspan="3" ><%=rs("jifen")%></td>
                                  </tr>
                                  <tr> 
                                    <td >登 陆：</td>
                                    <td  style="padding-left:10px"><%=rs("m_login_count")%> 次</td>
                                    <td >订单数：</td>
                                    <td  style="PADDING-LEFT: 10px">&nbsp; <%dim rs2
			set rs2=conn.execute("select count(id) from u_order where u_id="&Id)
			response.write rs2(0)&"笔订单"
			rs2.close
			set rs2=nothing
			%></td>
                                  </tr>
                                  <tr style="display:none">
								  <td >查找此用户的所有定单：</td>
								  <td height="30" colspan="3" >
								  <select name="zhuangtai" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" ><base target=Right> 
                                        <option value="" selected>--选择查讯状态--</option>
                                        <option value="editdingdan.asp?zhuangtai=0&namekey=<%=trim(rs("m_uname"))%>" >全部订单状态</option>
                                        <option value="editdingdan.asp?zhuangtai=1&namekey=<%=trim(rs("m_uname"))%>" >未作任何处理</option>
                                        <option value="editdingdan.asp?zhuangtai=2&namekey=<%=trim(rs("m_uname"))%>" >用户已经划出款</option>
                                        <option value="editdingdan.asp?zhuangtai=3&namekey=<%=trim(rs("m_uname"))%>" >服务商已经收到款</option>
                                        <option value="editdingdan.asp?zhuangtai=4&namekey=<%=trim(rs("m_uname"))%>" >服务商已经发货</option>
                                        <option value="editdingdan.asp?zhuangtai=5&namekey=<%=trim(rs("m_uname"))%>" >用户已经收到货</option>
                                    </select>
                                    </td>
                                  </tr>
			<%rs.close%>
</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
		<input class="Button" onclick="Sends('UserEdit','manageuser.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" type="Submit" name="Submit" value="修 改">   <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<script language = "JavaScript" charset="gb2312" src="/member/js/GetProvince.js"></script>
					<script language="javascript">
					   
function changeProvince(selvalue)
{
document.UserEdit.hukoucapital.length=0; 
document.UserEdit.hukoucity.length=0;
var selvalue=selvalue;	  
var j,d,mm;
d=0;
for(j=0;j<provincearray.length;j++) 
	{
		if(provincearray[j][1]==selvalue) 
		{
			if (d==0)
			{
			mm=provincearray[j][2];
			}
		var newOption2=new Option(provincearray[j][0],provincearray[j][2]);
		document.all.hukoucapital.add(newOption2);
		d=d+1;	
		}		
		if(provincearray[j][1]==mm) 
		{		
			var newOption3=new Option(provincearray[j][0],provincearray[j][2]);
			document.all.hukoucity.add(newOption3);
		}			
	}
}
function changeCity(selvalue)  
{ 
	document.UserEdit.hukoucity.length=0;  
	var selvalue=selvalue;
	var j;
	for(j=0;j<provincearray.length;j++) 
	{
		if(provincearray[j][1]==selvalue) 
		{
			var newOption4=new Option(provincearray[j][0],provincearray[j][2]);
			document.all.hukoucity.add(newOption4);
		}
	}
}
function selectprovince() 
{ 
	var j;
	for(j=0;j<provincearray.length;j++) 
	{
		if(provincearray[j][1]==0) 
		{
			var newOption4=new Option(provincearray[j][0],provincearray[j][2]);
			document.all.hukouprovince.add(newOption4);
		}
	}
}
selectprovince();
					</script>
					<%
End Sub
	%>
<!--#Include File="../Code.asp"-->