<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：PlusForm.asp
'文件用途：互动管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_PlusForm_Title,Fk_PlusForm_Content,Fk_PlusForm_Name,Fk_PlusForm_Contact,Fk_PlusForm_Ip,Fk_PlusForm_Time,Fk_PlusForm_ReContent,Fk_PlusForm_ReAdmin,Fk_PlusForm_ReIp,Fk_PlusForm_ReTime
Dim Ext_Table_Name,modeid,Ext_Form_Name

modeid=Trim(Request.QueryString("modeid"))
ID = Request("ID")
'获取参数
Types=Clng(Request.QueryString("Type"))

	Sqlstr="Select Ext_Form_Name,Ext_Table_Name From [Ext_FormModel] Where id=" & modeid
	set rs=conn.execute(sqlstr)
	
	If Not Rs.Eof Then
		Ext_Form_Name=Rs("Ext_Form_Name")
		Ext_Table_Name=Rs("Ext_Table_Name")
	Else
		PageErr=1
	End If
	Rs.Close
	
Select Case Types
	Case 1
		Call PlusFormList() '互动列表
	Case 2
		Call PlusFormReForm() '回复互动表单
	Case 3
		Call PlusFormReDo() '执行回复互动
	Case 4
		Call PlusFormDelDo() '执行删除互动
	Case Else
		Response.Write("没有找到此功能项！")
End Select

	sub PlusFormReForm()
 	dim MX_Arr,k,mx
	set rs=conn.execute("SELECT * FROM Ext_Table_fields Where FormID=" & ModeID & " order by OrderID desc,ID asc ")
    If  rs.eof Then response.Write "没有找到这条记录,请返回":response.end
	do while not rs.eof
		mx=mx&rs("FieldName")&","
 	rs.movenext
	loop
	 mx=Left(mx, Len(mx) - 1)
	 rs.close
  	set rs=conn.execute("select "&mx&" ,* from "&Ext_Table_Name&"_Form where id="&id&" order by id desc")
     If  rs.eof Then response.Write "没有找到这条记录,请返回":response.end
 
 conn.execute("update "&Ext_Table_Name&"_Form set viewstatu=1 where id="&id&"")
 
	 %>
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="table">
   <tr>
    <td align="right" >提交时间：&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td align="left">

<%=rs("UpdateTime") %>

</td>
  </tr> 
  
     <tr  onMouseOver=overColor(this) onMouseOut=outColor(this)>
    <td align="right" >用户IP：&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td align="left">

<%=rs("UserIP") %>

</td>
  </tr> 
  
  
  <% =ACT_MXEdit(ModeID,ID) 
    %>
</table>
</div>

<div id="BoxBottom" style="width:93%; margin: 0 auto; text-align:center;" class="tcbtm">
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
<% 



end sub 
	
	Public Function ACT_MXEdit(ModeID,ID)'表现方式.输出模型
	 Dim RS
	  Set RS=conn.execute("Select * from Ext_Table_fields  Where FormID=" & ModeID & " order by OrderID desc,ID asc")
	  	Do While Not RS.Eof
			ACT_MXEdit=ACT_MXEdit &"<tr>"&vbCrLf&"<td width=""13%"" align=""right"">"&RS("Title")&"：&nbsp;&nbsp;&nbsp;&nbsp;</td>"&vbCrLf&"<td>"&EditField(RS,ModeID,ID)&"</td>"&vbCrLf&"</tr>"&vbCrLf
			
		RS.MoveNext
		Loop
	  RS.Close:Set RS=Nothing
	 ACT_MXEdit=vbCrLf&ACT_MXEdit& vbCrLf 
	End function


	Function EditField(RSObj,ModeID,id)
		Dim i,IsNotNull,TitleTypeArr,checked,rs1,FieldName
		Dim arrtitle,arrvalue,titles
	  Set RS1=conn.execute("Select * from "&Ext_Table_Name&"_Form  Where id="&id&"")
	  FieldName= RSObj("FieldName")
	
		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title=""必填"">*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
		 
	    EditField= RS1(FieldName)& vbCrLf 
	  RS1.Close:Set RS1=Nothing
	End Function 

'==========================================
'函 数 名：PlusFormList()
'作    用：互动列表
'参    数：
'==========================================
Sub PlusFormList()
	Session("NowPage")=FkFun.GetNowUrl()
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	
	' dim mx
	' sqlstr="SELECT * FROM Ext_Table_fields Where  FormID=" & modeid & " order by OrderID desc,ID asc"
	' set rs=conn.execute(sqlstr)
	' If rs.eof Then response.Write "没有找到这条记录,请返回":response.end
	' If Not Rs.Eof Then
		' do while not rs.eof
			' mx=mx&rs("FieldName")&","
			' rs.movenext
		' loop
		 ' mx=Left(mx, Len(mx) - 1)
	' end if
	' rs.close
	' set rs=conn.execute("select "&mx&" ,* from "&Ext_Table_Name&"_Form where order by id desc")
     ' If  rs.eof Then response.Write "没有找到这条记录,请返回":response.end
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>【<%=Ext_Form_Name%>】模块</h3>
        <a class="no3" href="javascript:void(0);" onclick="SetRContent('MainRight','PlusForm.asp?Type=1&modeid=<%=modeid%>');return false">刷新内容</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop">编号</th>
            <th align="center" class="ListTdTop">时间</th>
            <th align="center" class="ListTdTop">来源IP</th>
            <th align="center" class="ListTdTop">状态</th>
            <th align="center" class="ListTdTop">操作</th>
        </tr>
<%
	Sqlstr="Select * From ["&Ext_Table_Name&"_Form] Order By UpdateTime Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Dim PlusFormTemplate
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
%>
        <tr>
            <td height="20" align="center"><%=Rs("id")%></td>
            <td align="center"><%=Rs("UpdateTime")%></td>
            <td align="center"><%=Rs("UserIP")%></td>
            <td align="center"><%if Rs("ViewStatu")=1 then response.write "<font style='color:green;font-weight:bolder;'>已读</font>":else:response.write "<font style='color:red;font-weight:bolder;'>未读</font>"%></td>
            <td align="center" class="no6"><a class="no2" href="javascript:void(0);" onclick="ShowBox('PlusForm.asp?Type=2&modeid=<%=modeid%>&Id=<%=Rs("id")%>','详细信息','700px','450px');"></a> <a style="margin-right:0;" class="no4" href="javascript:void(0);" onclick="DelIt('您确认要删除？此操作不可逆！','PlusForm.asp?Type=4&modeid=<%=modeid%>&Id=<%=Rs("id")%>','MainRight','<%=Session("NowPage")%>');"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="5">&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("PlusForm.asp?Type=1&modeid="&modeid&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="5" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==============================
'函 数 名：PlusFormDelDo
'作    用：执行删除互动
'参    数：
'==============================
Sub PlusFormDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	conn.execute("delete * from ["&Ext_Table_Name&"_Form] Where Id=" & Id)
	Response.Write("信息删除成功！")
End Sub
%><!--#Include File="../Code.asp"-->