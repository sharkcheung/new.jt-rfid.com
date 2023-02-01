<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../member/func2.asp"-->
<script language=JavaScript>
function chsel(){
		with (document.form1){
			if(szSheng.value) {
				szShi.options.length=0;
				for(var i=0;i<selects[szSheng.value].length;i++){
					szShi.add(selects[szSheng.value][i]);
				}
			}
		}
	}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
<%dim userid,resume_hukouprovinceid,resume_hukoucapitalid,resume_hukoucityid,province_name,capital_name,city_name,workadd,mystring,length,Hwasrs,province,Hwassql
		userid=request.querystring("id")
		if userid<>"" then
if not isnumeric(userid) then 
response.write"<script>alert(""非法访问!"");location.href=""../index.asp"";</script>"
response.end
end if
end if
		set rs=server.createobject("adodb.recordset")
		rs.open "select * from u_members where id="&userid ,conn,1,1
		resume_hukouprovinceid=rs("szSheng")
		resume_hukoucapitalid=rs("szShi")
		resume_hukoucityid=rs("szXian")
		%>
<form target="saveiframe" name="SystemSet" id="SystemSet" method="post" action="saveuser.asp?action=save&id=<%=userid%>">
<div id="BoxTop" style="width:750px;"><span>会员详细资料</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a>
</div>
<div id="BoxContents" style="width:750px;">
<table width="97%" border="0" align="center" cellpadding="5" cellspacing="0" >
<tr> 
                                    <td >
									账 号：</td>
                                   <td><font style="font-weight:bold; letter-spacing:1px; color:#FF0000; font-size:14px;"><%=trim(rs("m_uid"))%></font><!--								<font color=#FF0000>用户类型：</font>
									<select name="reglx">
									<option value="1" <if rs("reglx")=1 then%>selected<end if%>>普通会员</option>
									<option value="2" <if rs("reglx")=2 then%>selected<end if%>>VIP 会员</option>
									</select>
									<font color=#FF0000>VIP期限：</font>
									<input class="Input"  type="text" name="vipdate" size="10" value="<=rs("vipdate")%>">
									格式：2004-02-22</td>-->
    </tr>
                                   <td   >
 									密 码：</tr>
                                   <td   >
									<input class="Input"  name="userpassword" type="text" id="userpassword" size="20">
									<font color=#FF0000>不改密码请留空!</tr>
                                  <tr> 
                                    <td >姓 名：</td>
                                    <td ><input class="Input"  name="userzhenshiname" type="text" id="userzhenshiname" size="20" value="<%=trim(rs("m_uname"))%>"></td>
                                    <td >
									性 别：</td>
                                    <td >
									<input class="Input"  type="radio" name="shousex" value=0 <%if rs("m_usex")=0 then%>checked<%end if%> checked>男
									<input class="Input"  type="radio" name="shousex" value=1 <%if rs("m_usex")=1 then%>checked<%end if%>>女　　　年 龄：<input class="Input"  name=nianling type=text value="<%=trim(rs("m_uage"))%>" size="3" maxlength="2" onKeyPress="event.returnValue=IsDigit();"></td>
    </tr><tr> 
                                    <td >密码提问：</td>
                                    <td ><select class="Input"  name="quesion" id="quesion">
						<option value="<%=trim(rs("m_question"))%>"><%
						select case trim(rs("m_question"))
						case 0
						response.write "我身份证最后6位数"
						case 1
						response.write "我父亲的名字"
						case 2
						response.write "我母亲的名字"
						case 3
						response.write "我就读的小学校名"
						case 4
						response.write "我最喜欢的颜色"
                     end select
						%></option>
						</select>
									</td>
                                    <td >
									密码答案：</td>
                                    <td >
									<input class="Input" size="20" name="answer" type="text" id="answer" value="<%=rs("m_answer")%>">
									</td></tr>
                                
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
                                    <td ><select class="Input" name="hukouprovince" size="1" id="select5" onChange="changeProvince(document.form1.hukouprovince.options[document.form1.hukouprovince.selectedIndex].value)">
		<%if resume_hukouprovinceid<>"" then%>
		<option value="<%=resume_hukouprovinceid%>"><%=Hireworkadds(resume_hukouprovinceid)%></option>
		<%else%>
		<option value="">选择省</option>
		<%end if%>
		</select>
						<select class="Input" name="hukoucapital" onChange="changeCity(document.form1.hukoucapital.options[document.form1.hukoucapital.selectedIndex].value)">
						  <%if resume_hukoucapitalid<>"" then%>
						  <option value="<%=resume_hukoucapitalid%>"><%=Hireworkadds(resume_hukoucapitalid)%></option>
						  <%else%>
						  <option value="">选择市</option>
						  <%end if%>
	                      </select>
		                  <select class="Input" name="hukoucity">
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
                                      <textarea name="content" rows="3" cols="60"><%=trim(rs("content"))%></textarea>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td >注 册：</td>
                                    <td ><%=rs("m_reg_time")%></td>
                                    <td >最后登陆：</td>
                                    <td  style="PADDING-LEFT: 8px">&nbsp;<%=rs("m_last_logintime")%></td>
                                  </tr>
                                  <tr style="display:none">
                                    <td >购物积分</td>
                                    <td colspan="3" ><%=rs("jifen")%></td>
                                  </tr>
                                  <tr> 
                                    <td >登 陆：</td>
                                    <td ><%=rs("m_login_count")%> 次</td>
                                    <td >订单数：</td>
                                    <td  style="PADDING-LEFT: 8px">&nbsp; <%dim rs2
			set rs2=conn.execute("select count(id) from u_order where u_id="&userid)
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
			<%rs.close
			set rs=nothing
			conn.close
			set conn=nothing%>
</table>
<iframe name="saveiframe" src="" height="1" width="1"></iframe>
</div>
<div id="BoxBottom" style="width:730px;">&nbsp;<input class="Button" onclick="$('#Boxs').hide();$('select').show();" type="Submit" name="Submit" value="修 改">   <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<script language = "JavaScript" charset="gb2312" src="/member/js/GetProvince.js"></script>
					<script language="javascript">
					   
function changeProvince(selvalue)
{
document.form1.hukoucapital.length=0; 
document.form1.hukoucity.length=0;
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
	document.form1.hukoucity.length=0;  
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
