

<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Get.asp
'文件用途：页面信息拉取页面
'==========================================

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call GetTopNav() '读取顶部菜单
	Case 2
		Call GetLeftNav() '读取左侧菜单
	Case 3
		Call GetMain() '读取管理首页信息
	Case 4
		Call GetUserInfo() '读取管理用户信息
	Case 5
		Call GetAbout() '读取系统版权
	Case 7
		Call GetMenuNav() '读取分菜单模块
	Case 8
		Call GetLeftNav2() '读取左侧菜单2
End Select

'==========================================
'函 数 名：GetTopNav()
'作    用：读取顶部菜单
'参    数：
'==========================================
Sub GetTopNav()
%>
    	<ul id="TopNav">
        	<li onclick="SetRContent('MainLeft','Get.asp?Type=2');SetRContent('MainRight','Module.asp?Type=1&MenuId=1');$('#TopNav li').removeClass('NavNow');$('#TopNav li').addClass('NavOther');$('#Nav_Main').removeClass('NavOther');$('#Nav_Main').addClass('NavNow');" id="Nav_Main" class="NavNow">管理设置</li>
        	
     <li onclick="SetRContent('MainLeft','Get.asp?Type=8');$('#TopNav li').removeClass('NavNow');$('#TopNav li').addClass('NavOther');$('#Nav_NS').removeClass('NavOther');$('#Nav_NS').addClass('NavNow');" style="display:none;" id="Nav_NS" class="NavOther">内容设置</li>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
        	<li onclick="SetRContent('MainLeft','Get.asp?Type=7&MenuId=<%=Rs("Fk_Menu_Id")%>');$('#TopNav li').removeClass('NavNow');$('#TopNav li').addClass('NavOther');$('#Nav_S<%=Rs("Fk_Menu_Id")%>').removeClass('NavOther');$('#Nav_S<%=Rs("Fk_Menu_Id")%>').addClass('NavNow');SetRContent('MainRight','Module.asp?Type=1&MenuId=<%=Rs("Fk_Menu_Id")%>');" id="Nav_S<%=Rs("Fk_Menu_Id")%>" class="NavOther"><%=Rs("Fk_Menu_Name")%></li>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </ul>
        <div class="Cal"></div>
<%
End Sub

'==========================================
'函 数 名：GetLeftNav()
'作    用：读取左侧菜单
'参    数：
'==========================================
Sub GetLeftNav()
%>
<ul>
	<li><a href="javascript:void(0);"><em class="active">&nbsp;</em>管理设置</a>
	<ul style="display:block">
<%If FkFun.CheckLimit("System1") Then%>
<li><a href="javascript:void(0);" onclick="ShowBox('siteset.asp?Type=1&Snr=1','基础信息','1000px','500px');;return false; "><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>基础信息</a></li>
<li><a href="javascript:void(0);" onclick="ShowBox('siteset.asp?Type=1&Snr=0','广告橱窗','1000px','500px');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>广告橱窗</a></li>
<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Admin.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>账号密码</a></li>
<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Data.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>数据维护</a></li>
 <%If FkFun.CheckLimit("System12") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Vote.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>市场调查</a></li><%End If%>
        <%If FkFun.CheckLimit("System2") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Friends.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>友情链接</a></li><%End If%>
<%End If%>
        	<%If Request.Cookies("FkAdminLimitId")=0 Then%><li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Limit.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>权限管理</a></li>
<%If FkFun.CheckLimit("System3") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Template.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>模板管理</a></li><%End If%>
<%If FkFun.CheckLimit("System3") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','TemplateHelp.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>模板标签生成器</a></li><%End If%>

        	<%End If%>
<%If FkFun.CheckLimit("System9") Then%>        	<li><a href="javascript:void(0);" onclick="ShowBox('HTML.asp?Type=1','生成管理','1000px','500px');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>生成管理</a></li><%End If%>
<%If FkFun.CheckLimit("System10") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Abd.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>JS广告调用管理</a></li><%End If%>
       
<%If FkFun.CheckLimit("System14") Then%><li><a href="javascript:void(0);" onclick="SetRContent('MainRight','File.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>上传文件管理</a></li><%End If%>

  <%If Request.Cookies("FkAdminLimitId")=0 Then%>
<li><a href="javascript:void(0);" onclick="ShowBox('Jpeg.asp?Type=1','水印缩略');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>水印缩略</a></li>    	
        	<li style="display:none"><a href="javascript:void(0);" onclick="ShowBox('QQ.asp?Type=1','客服浮窗代码管理');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>客服浮窗代码管理</a></li><%End If%>
<%If FkFun.CheckLimit("System16") Then%><li><a href="javascript:void(0);" onclick="ShowBox('DelWord.asp?Type=1','过滤字符管理','1000px','500px');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>过滤字符管理</a></li><%End If%>
<%If FkFun.CheckLimit("System4") Then%>        	<li style="display:none"><a href="javascript:void(0);" onclick="SetRContent('MainRight','Job.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>招聘管理</a></li><%End If%>
<%If FkFun.CheckLimit("System8") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Recommend.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>推荐类型管理</a></li><%End If%>
<%If FkFun.CheckLimit("System7") Then%>        	<li style="display:none"><a href="javascript:void(0);" onclick="SetRContent('MainRight','Subject.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>专题管理</a></li><%End If%>
<%If Request.Cookies("FkAdminLimitId")=0 Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Field.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>自定义字段管理</a></li><%End If%>
<%If Request.Cookies("FkAdminLimitId")=0 Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','ExtFunction.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>扩展功能管理</a></li><%End If%>

<%If Request.Cookies("FkAdminLimitId")=0 Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','FormMaker.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>自定义表单</a></li><%End If%>

<!--#Include File="api.asp"-->
	</ul>
</li>
        
<%If FkFun.CheckLimit("System20") Then%> 
<li><a href="javascript:void(0);"><em class="active">&nbsp;</em>电子商务</a>
	<ul  style="display:block">
	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','editdingdan.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>订单管理</a></li> 	
	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','manageuser.asp?Type=1');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>会员管理</a></li> 
	<li><a href="javascript:void(0);" onclick="ShowBox('zhifu.asp?Type=1','支付方式设置');return false;"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>支付设置</a></li>
	<!--<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','address.asp?Type=1');"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>地址管理</a></li>  -->     
	 </ul>
</li>
<%End If%>	  
</ul>

<div style="display:none"><a href="javascript:void(0);" class="LeftMenuTop" onclick="OpenMenu('M3');">其他管理</a></div>
<ul id="M3" style="display:none;">
<%If FkFun.CheckLimit("System17") Then%><li><a href="javascript:void(0);" onclick="ShowBox('KeyWord.asp?Type=1','关键词库管理')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>关键词库管理</a></li><%End If%>
<%If FkFun.CheckLimit("System11") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Word.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>关键词链接管理</a></li><%End If%>

<%If FkFun.CheckLimit("System15") Then%><li><a href="javascript:void(0);" onclick="ShowBox('Map.asp?Type=1,'SEO索引地图生成'');"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>SEO索引地图生成</a></li><%End If%>
</ul>
<%
End Sub

'==========================================
'函 数 名：GetMain()
'作    用：读取管理首页信息
'参    数：
'==========================================
Sub GetMain()
%>
    	<div id="MainRrightTop">欢迎使用<%=FkSystemName%>！</div>
        <div id="NewBox">
        	<div id="NewBoxTop">
            	<ul>
                	<li>官方公告</li>
                </ul>
            </div>
            <div id="NewContent">
            	官方公告
            </div>
        </div>
        <div id="Ad">
            官方推荐
        </div>
        <div class="Cal"></div>
        <div id="AboutSystem">
        	<p>系统名称：<%=FkSystemName%>&nbsp;&nbsp;&nbsp;&nbsp;系统版本：<%=FkSystemVersion%></p>
        	<p>版权所有：深圳市企帮网络技术有限公司&nbsp;&nbsp;&nbsp;&nbsp;技术支持：QQ22925339</p>
        </div>
        <div id="AboutAd">
        	<p>www.qebang.cn</p>
        </div>
        <div class="Cal"></div>
<%
End Sub

'==========================================
'函 数 名：GetUserInfo()
'作    用：读取管理用户信息
'参    数：
'==========================================
Sub GetUserInfo()
%>
您的帐号是：&nbsp;&nbsp;<%=Request.Cookies("FkAdminName")%>&nbsp;&nbsp;[&nbsp;<a href="<%=SiteDir%>" target="_blank" title="前台首页">前台首页</a> <a href="javascript:void(0);" onclick="ShowBox('PassWord.asp?Type=1');" title="修改密码">修改密码</a> <a href="Logout.asp" title="退出登录">退出登录</a>&nbsp;]
<%
End Sub

'==========================================
'函 数 名：GetAbout()
'作    用：读取系统版权
'参    数：
'==========================================
Sub GetAbout()
%>
<div id="BoxTop" style="width:500px;">关于本系统[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:500px;">
	<table width="90%" border="1" bordercolor="#CCCCCC" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">系统名称：</td>
	        <td>&nbsp;<%=FkSystemName%></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">系统版本：</td>
	        <td>&nbsp;<%=FkSystemVersion%></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">开发商：</td>
	        <td>&nbsp;深圳市企帮网络技术有限公司</td>
	        </tr>
	    <tr>
	        <td height="25" align="right">官方主站：</td>
	        <td>&nbsp;<a href="http://www.qebang.cn/" target="_blank" title="访问官方主站">http://www.qebang.cn/</a></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">技术支持：</td>
	        <td>&nbsp;22925339；</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:480px;">
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub

'==========================================
'函 数 名：GetMenuNav()
'作    用：读取分菜单模块
'参    数：
'==========================================
Sub GetMenuNav()
	Dim Rs2
	Id=Clng(Request.QueryString("MenuId"))
%>
		<ul>
    	<%If FkFun.CheckLimit("System13") Then%>
		<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Infos.asp?Type=1');return false;">首页模块</a></li><%End If%>
    	<%
	dim zimenu,modulename
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level=0 Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
		modulename=Rs("Fk_Module_Name")
		Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level="&Rs("Fk_Module_Id")&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
		Rs2.Open Sqlstr,Conn,1,3
		if not rs2.eof then
			zimenu=1
		else
			zimenu=0
		end if
%>
<li>
<a  <%if zimenu=1 then%>title="点击右边箭头展开子菜单 "<%end if%> style="cursor:text" href="javascript:void(0);" class="LeftMenuTop" title="<%=modulename%>">
<%if zimenu=1 then%><em>&nbsp;</em><%end if%> <font style="cursor:pointer;" onclick="<%=GetNavGo(Rs("Fk_Module_Type"),Rs("Fk_Module_Id"))%>"><%=cutStr(modulename,10)%></font>
</a>
<%
		
		If Not Rs2.Eof Then
%>
        <ul>
<%
			While Not Rs2.Eof
		modulename=Rs2("Fk_Module_Name")
%>   	
	<li><a href="javascript:void(0);" onclick="<%=GetNavGo(Rs2("Fk_Module_Type"),Rs2("Fk_Module_Id"))%>" title="<%=modulename%>"><%=cutStr(modulename,9)%></a>
         
		   <%
		   dim Rs3
		   Set Rs3=Server.Createobject("Adodb.RecordSet")
          Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level="&Rs2("Fk_Module_Id")&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs3.Open Sqlstr,Conn,1,3
			if not rs3.eof then
			%>
        <ul>
			<%
			While Not Rs3.Eof
		modulename=Rs3("Fk_Module_Name")
			
          %>
		<li><a href="javascript:void(0);" onclick="<%=GetNavGo(Rs3("Fk_Module_Type"),Rs3("Fk_Module_Id"))%>" title="<%=modulename%>">└ <%=cutStr(modulename,7)%></a>
          <%
		  
		   dim Rs4
		   Set Rs4=Server.Createobject("Adodb.RecordSet")
          Sqlstr="Select * From [Fk_Module] Where Fk_Module_Menu="&Id&" And Fk_Module_Level="&Rs3("Fk_Module_Id")&" Order By Fk_Module_Order Asc,Fk_Module_Id Asc"
			Rs4.Open Sqlstr,Conn,1,3
			if not rs4.eof then
			%>
			<ul>
			<%
			While Not Rs4.Eof
		modulename=Rs4("Fk_Module_Name")
          %>
		<li><a href="javascript:void(0);" onclick="<%=GetNavGo(Rs4("Fk_Module_Type"),Rs4("Fk_Module_Id"))%>" title="<%=modulename%>">└ <%=cutStr(modulename,6)%></a></li>
          <%
          	Rs4.MoveNext
			Wend
          %>
		  </ul><%end if%></li>
          <%
			Rs4.Close
			
          	Rs3.MoveNext
			Wend
          %>
		  </ul><%end if%></li>
<%
			Rs3.Close
			Rs2.MoveNext
			Wend
%>
        </ul>
<%
		End If
		Rs2.Close
		Rs.MoveNext
	Wend
	Rs.Close
%>
</li>
</ul>
<%
End Sub

'==========================================
'函 数 名：GetLeftNav2()
'作    用：读取左侧菜单2
'参    数：
'==========================================

Sub GetLeftNav2()
%>
    	<div id="QuickNav"><a href="javascript:void(0);" onclick="ShowBox('Get.asp?Type=5');" id="QuickNav1"></a><a href="javascript:void(0);" onclick="alert('反馈系统暂时关闭，欢迎通过QQ或者MAIL跟我们联系！');" id="QuickNav2"></a><div class="Cal"></div></div>
<%If FkFun.CheckLimit("System6") Then%>        <div><a href="javascript:void(0);" class="LeftMenuTop" onclick="OpenMenu('M2');">菜单管理</a></div>
        <ul id="M2">
        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Menu.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>菜单管理</a></li>
<%
	Sqlstr="Select * From [Fk_Menu] Order By Fk_Menu_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Module.asp?Type=1&MenuId=<%=Rs("Fk_Menu_Id")%>')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><%=Rs("Fk_Menu_Name")%></a></li>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
        </ul><%End If%>
        <div><a href="javascript:void(0);" class="LeftMenuTop" onclick="OpenMenu('M3');">其他管理</a></div>
        <ul id="M3">
<%If FkFun.CheckLimit("System2") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','FriendsType.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>友情链接类型管理</a></li><%End If%>
<%If FkFun.CheckLimit("System2") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Friends.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>友情链接管理</a></li><%End If%>
<%If FkFun.CheckLimit("System4") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Job.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>招聘管理</a></li><%End If%>
<%If FkFun.CheckLimit("System10") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Abd.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>广告管理</a></li><%End If%>
<%If FkFun.CheckLimit("System11") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Word.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>站内关键字管理</a></li><%End If%>
<%If FkFun.CheckLimit("System8") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Recommend.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>推荐类型管理</a></li><%End If%>
<%If FkFun.CheckLimit("System7") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Subject.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>专题管理</a></li><%End If%>
<%If FkFun.CheckLimit("System12") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Vote.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>在线投票管理</a></li><%End If%>
<%If FkFun.CheckLimit("System13") Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Infos.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>独立信息管理</a></li><%End If%>
<%If Request.Cookies("FkAdminLimitId")=0 Then%>        	<li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Field.asp?Type=1')"><span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>自定义字段管理</a></li><%End If%>
        </ul>

<%
End Sub
private function cutStr(strString,intLen)
	if len(strString)>intLen then
		strString=left(strString,intLen)&"..."
	end if
	cutStr=strString
end function
%><!--#Include File="../Code.asp"-->

