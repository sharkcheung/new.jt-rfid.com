<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_HTML.asp"-->

<%
'==========================================
'文 件 名：FormMaker.asp
'文件用途：内容管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Form_Name,Fk_Form_TimeLimit,Fk_Form_Statu,Fk_Form_Code,Fk_Form_EndTime,Fk_Form_Describ,Fk_Form_StarTime,Fk_Form_Tablename,Fk_Form_SendMail

dim TitleSize,ISType,RadioPic_Type,check,RadioType_Content,ListBoxType_Content,SupportHtmlType_Width,SupportHtmlType_heigh,MultipleTextType_Width,MultipleTextType_Height,IsEditor,IsNotNull,OrderID,YHtml,savepic,ModeName,ModeID,title,FieldName,Description,fun,regEx,regError,SearchIF,ValueOnly,FieldType,RadioType_Type,ListBoxType_Type,Type_Default,ColumnType,Type_Type,content,TableName
'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call FormMakeList() '内容列表
	Case 2
		Call FormAddForm() '添加内容表单
	Case 3
		Call FormAddDo() '执行添加内容
	Case 4
		Call FormModelEditForm() '修改内容表单
	Case 5
		Call FormModelEditDo() '执行修改内容
	Case 6
		Call FormModelDelDo() '执行删除内容
	Case 7
		Call ListDelDo() '执行批量删除内容
	Case 8
		Call FormFieldList() '字段列表
	Case 9,10
		Call FormFieldAddEdit() '字段添加、修改
	Case 11
		Call FormFieldAddSave() '执行字段添加
	Case 12
		Call FormFieldEditSave() '执行字段修改
	Case 13
		Call FormFieldDel() '执行字段删除
	Case 14
		call FormHtmlShow()
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FormMakeList()
'作    用：内容列表
'参    数：
'==========================================
Sub FormMakeList()
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
%>


<div id="ListContent">
	<div class="gnsztopbtn">
    	<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle; margin-left:20px;"/>&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','FormMaker.asp?Type=1&SearchStr='+escape(document.all.SearchStr.value));" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/>
    	<a class="tjia" href="javascript:void(0);" onclick="ShowBox('FormMaker.asp?Type=2');">添加</a>
        <a class="shuax" href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a>
    </div>
    <form name="DelList" id="DelList" method="post" action="FormMaker.asp?Type=7" onsubmit="return false;">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop">选</th>
            <th align="center" class="ListTdTop">表单名称</th>
            <th align="center" class="ListTdTop">是否开始</th>
            <th align="center" class="ListTdTop">状态</th>
            <th align="center" class="ListTdTop" width="250">操作</th>
        </tr>
<%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Ext_FormModel] "
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And Ext_Form_Name Like '%%"&SearchStr&"%%'"
	End If
	on error resume next
	Rs.Open Sqlstr,Conn,1,1
	if err then
		err.clear
		conn.execute("create table Ext_FormModel(id integer identity(1,1) primary key,Ext_Form_Name varchar(100),Ext_Table_Name varchar(100),UnlockTime int default 0,IsMail int default 0,Ext_Form_Statu int default 0,StartTime date,EndTime date,FormCode int default 0,Ext_Form_Describ varchar(100))")
		Rs.Open Sqlstr,Conn,1,1
	end if
	If Not Rs.Eof Then
		Dim ArticleTemplate
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
            <td height="20" align="center"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("Id")%>" id="List<%=Rs("Id")%>" /></td>
            <td align="left" class="td1">&nbsp;&nbsp;<%=Rs("Ext_Form_Name")%></td>
            <td align="center"><%If now >Rs("StartTime") or Rs("UnlockTime") =1 Then 
		response.write "<font color=green title=结束日期是"& Rs("EndTime")&">已经开始</a>"
	Else
		response.write "<font color=red title=开始日期是"& Rs("StartTime")&">还没有开始</font>"
	End if%></td>
            <td align="center"><%If Rs("Ext_Form_Statu")=0 Then%><span class="gnszxianshi "></span><%Else%><span class="gnszxianshi hidden"></span><%End If%></td>
            <td align="center"><a style="width:auto; line-height:21px;" title="字段列表 " href="javascript:void(0);" onclick="SetRContent('MainRight','FormMaker.asp?Type=8&Formid=<%=Rs("Id")%>&tb=<%=Rs("Ext_Table_Name")%>');return false;"  style="width:auto; line-height:21px;">字段列表</a>&nbsp;| <a  style="width:auto; line-height:21px;" title="HTML调用 " href="javascript:void(0);" onclick="ShowBox('FormMaker.asp?Type=14&Id=<%=Rs("Id")%>&A=HTML','HTML调用代码','1000px','500px');return false;">HTML调用</a>&nbsp;| <a  style="width:auto; line-height:21px;" title="修改 " href="javascript:void(0);" onclick="ShowBox('FormMaker.asp?Type=4&Id=<%=Rs("Id")%>','修改');return false;">修改</a>&nbsp;| <a  style="width:auto; line-height:21px;" title="删除 " href="javascript:void(0);" onclick="DelIt('确认删除？此操作会将表单相关内容全部删除！','FormMaker.asp?Type=6&Id=<%=Rs("Id")%>','MainRight','<%=Session("NowPage")%>');return false;">删除</a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="5">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="vertical-align:middle; margin-left:19px;">&nbsp;<label for="chkall" style="vertical-align:middle;">全选</label>
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除选中的内容吗？')){Sends('DelList','FormMaker.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}" style="vertical-align:middle;">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("FormMaker.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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
    </form>
</div>
<div id="ListBottom">

</div>
<%
end sub

'==========================================
'函 数 名：FormHtmlShow()
'作    用：HTML调用表单
'参    数：
'==========================================
Sub FormHtmlShow()
	dim A
	modeid=Request.QueryString("id")
	A=Request.QueryString("A")
	set rs=conn.execute("select Ext_Form_Name from Ext_FormModel where ID="&ModeID&"")
	if not rs.eof then
		ModeName=rs(0)
	else
		rs.close
		response.end
	end if
	rs.close
%>
<form id="TemplateEdit" name="TemplateEdit" method="post" action="Template.asp?Type=5" onsubmit="return false;">
<!--<div id="BoxTop" style="width:700px;"> 【<%= ModeName %>】模块的HTML调用代码[按ESC关闭窗口]</div>-->
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td style="padding:10px 0;"><%If A="HTML" Then %><textarea name="textarea" class="TextArea" id="textarea" style="width:902px; color:#555; line-height:22px; height:340px; border:1px solid #ccc; padding:5px; font-size:12px;"><% Call ListForm(modeid) %></textarea>
	<%Else %><textarea name="textarea" id="textarea" style="width:902px; height:340px; border:1px solid #ccc; padding:5px; font-size:12px; color:#555; line-height:22px;"><%response.write "<script language=""javascript"" type=""text/javascript"" src=""plus/Form/ACT.F.ASP?ModeID="&ModeID&"""></script>"%></textarea>
	<%End If %></td>
        </tr>
	</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:center;" class="tcbtm">
        <input id="Button1" type="button" value="复 制" class="Button" onClick="A_CP('textarea')" /> &nbsp; <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>

<script>
			function A_CP(ob)
			{
				var obj=MM_findObj(ob); 
				if (obj) 
				{
					obj.select();js=obj.createTextRange();js.execCommand("Copy");}
					alert('复制成功，粘贴到你要调用的html代码里即可!');
				}
				function MM_findObj(n, d) { //v4.0
			  var p,i,x;
			  if(!d) d=document;
			  if((p=n.indexOf("?"))>0&&parent.frames.length)
			   {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
			   }
			  if(!(x=d[n])&&d.all) x=d.all[n];
			  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
  </script>
<%
end sub
	Sub ListForm(modeid)
		dim Act_Form
		set rs=conn.execute("SELECT FormCode FROM Ext_FormModel Where ID=" & ModeID & " order by ID desc")
		If Not rs.eof Then
 		   Act_Form=Act_Form & "<script type='text/javascript' src='/admin/dkidtioenr/kindeditor.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript' src='/admin/dkidtioenr/lang/zh_CN.js'></script>"& vbCrLf
 		   Act_Form=Act_Form & "<script type='text/javascript' src='http://image001.dgcloud01.qebang.cn/website/My97DatePicker/WdatePicker.js'></script>"& vbCrLf
		   Act_Form=Act_Form & "<script type='text/javascript'>KindEditor.ready(function(K) {window.editor = K.create('.kinediter',{resizeType : 1,allowPreviewEmoticons : false,allowImageUpload : false,items : ['fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold', 'italic', 'underline','removeformat', '|', 'justifyleft', 'justifycenter', 'justifyright', 'insertorderedlist','insertunorderedlist', '|', 'emoticons', 'image', 'link'],afterBlur:function(){this.sync();}"
		   Act_Form=Act_Form & "})});</script>"& vbCrLf
 		   Act_Form=Act_Form &"<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"& vbCrLf
  		   Act_Form=Act_Form & "<form name='myform' action='/plus/Form/?A=Save&ModeID=" & ModeID & "' method='post'> "& vbCrLf
 		   Act_Form=Act_Form& ACT_MXList(ModeID)& vbCrLf
				if Rs("FormCode")=0 then 
 					 Act_Form=Act_Form& "<tr><td>验证码：</td><td>"& vbCrLf
  					 Act_Form=Act_Form& "<input type='text' size='10' name='Code'> "& vbCrLf&"<img style='cursor:hand;'  src=""/plus/Code.asp"" id='IMG1' onclick=""this.src='/plus/Code.asp?s='+Math.random();"" alt='看不清楚? 换一张！'>"& vbCrLf
  					 Act_Form=Act_Form& "</td></tr>"& vbCrLf
 				end if 
 		   Act_Form=Act_Form& "<tr> <td  colspan='2' align='center'>"& vbCrLf
  		   Act_Form=Act_Form&"<input type=submit   name=Submit1 value='  提 交  ' />&nbsp;"& vbCrLf
 		   Act_Form=Act_Form& "<input type='reset' name='Submit2'  value='  重 置  ' /></td></tr>"& vbCrLf
		   Act_Form=Act_Form&  "</form>"& vbCrLf
		   Act_Form=Act_Form&  "</table>"& vbCrLf
 		   response.write server.HTMLEncode(Act_Form)
		 End if	
		rs.close
		
	End Sub 

	Public Function ACT_MXList(ModeID)'表现方式.输出模型
	 Dim RSObj
	 on error resume next
	  Set RSObj=conn.execute("Select * from Ext_Table_fields  Where FormID=" & ModeID & "  order by OrderID desc,ID asc")
	  if err then
	 err.clear
	 ACT_MXList=""
	else
		If Not rsobj.eof Then 
			Do While Not RSObj.Eof
 				ACT_MXList=ACT_MXList &"<tr><td  width='15%'  align='right'>"&RSObj("Title")&"：</td>"& vbCrLf&"<td align='left'>"&ListField(RSObj)&"</td></tr>"& vbCrLf
 			RSObj.MoveNext
			Loop
		End If 
	  RSObj.Close:Set RSObj=Nothing
	  end if

	End function


 

 
	Function ListField(RSObj)
		Dim i,TitleTypeArr,checked,IsNotNull
		Dim arrtitle,arrvalue,titles

		If rsobj("IsNotNull")="0" Then 
			IsNotNull="  <font color=red title='必填'>*</font>  "&rsobj("Description")
		Else
			IsNotNull="  "&rsobj("Description")
		End If 
 		 Select Case RSObj("FieldType")
		   Case "TextType"
				ListField= "<input type='text' title='"&RSObj("Description")&"' name='"&RSObj("FieldName")&"' size='"&RSObj("width")&"' value='"&RSObj("Type_Default")&"'>"&IsNotNull
		   Case "MultipleTextType"
				ListField= "<textarea title='"&RSObj("Description")&"' name='"&RSObj("FieldName")&"' style='height:"&RSObj("height")&"px;width:"&RSObj("width")&"px;'>"&RSObj("Type_Default")&"</textarea>"&IsNotNull
		   Case "MultipleHtmlType"
				ListField="<textarea id="&RSObj("FieldName")&" name="&RSObj("FieldName")&" style=width:"&RSObj("width")&"px ;height:"&RSObj("height")&"px class=""kinediter""></textarea>"
		   Case "RadioType"
				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  ListField= ListField&"<select  name='"&RSObj("FieldName")&"'>"
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If RSObj("Type_Default")=arrvalue Then checked="selected" Else checked=""
					ListField = ListField & "<option value='" & arrvalue & "' "&checked&">" & titles & "</option>"
				  Next
					ListField= ListField&" </select>"&IsNotNull
				Else
				  For I = 0 To UBound(TitleTypeArr)
				
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					
					If RSObj("Type_Default")=arrvalue Then checked="checked" Else checked=""
					ListField = ListField &"<label for='"&RSObj("FieldName")&i&"'> <input  id='"&RSObj("FieldName")&i&"' type='radio'  name='"&RSObj("FieldName")&"' value='"&arrvalue&"' "&checked&" />"&titles&"&nbsp;&nbsp;</label>" 
				  Next
				    ListField = ListField&IsNotNull
				End If 
		   Case "ListBoxType"
 				TitleTypeArr=Split(RSObj("Content"), vbCrLf)
				If RSObj("Type_Type")=0 Then 
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If RSObj("Type_Default")=arrvalue Then checked="checked" Else checked=""
					ListField = ListField &"<label for='"&RSObj("FieldName")&i&"'> <input  id='"&RSObj("FieldName")&i&"' type='checkbox'  name='"&RSObj("FieldName")&"' value='"&arrvalue&"' "&checked&" />"&titles&"&nbsp;&nbsp;</label>"
				  Next
				  ListField = ListField&IsNotNull
				Else
				  ListField= ListField&"<select  size='4'   style='width:300px;height:126px'  name='"&RSObj("FieldName")&"' multiple>"
				  For I = 0 To UBound(TitleTypeArr)
					arrtitle=Split(TitleTypeArr(I),"-")
					If  UBound(arrtitle)="1" Then 
						titles=arrtitle(0)
						arrvalue=arrtitle(1)
					ElseIf  TitleTypeArr(I)<>"" Then 
					titles=arrtitle(0)
					arrvalue=arrtitle(0)
					Else
						Exit for
					End If 
					If RSObj("Type_Default")=arrvalue Then checked="checked" Else checked=""
					ListField = ListField & "<option value='"& arrvalue & "' "&checked&">" & titles & "</option>"
				  Next
					ListField= ListField&" </select>"&IsNotNull
				End If 
		   Case "DateType"
				ListField= ListField&"<input name='"&RSObj("FieldName")&"' type='text' id='"&RSObj("FieldName")&"' value='' onclick='WdatePicker()'  class=""Wdate"" >"&IsNotNull
		   Case "PicType"
 			 	If RSObj("Type_Type")="0" Then 
					ListField= "<input  name="""&RSObj("FieldName")&""" type=""text""  value="""" size=""40"" id=""UpID""><input type=""button"" id=""uploadButton"" value=""上传"" class=""Button"" style=""vertical-align:middle;""/>"&IsNotNull
				Else
					ListField="<div id=""sapload"&RSObj("FieldName")&""">"& vbCrLf 
					ListField=ListField&	"</div>"& vbCrLf 
					ListField=ListField& "<script type=""text/javascript"">"& vbCrLf 
					ListField=ListField&"// <![CDATA["& vbCrLf 
					ListField=ListField&"var so = new SWFObject("""&ACTCMS.ACTSYS&"act_inc/sapload.swf"", ""sapload"&RSObj("FieldName")&""", ""450"", ""25"", ""9"", ""#ffffff"");"& vbCrLf 
					ListField=ListField&"so.addVariable('types','"&Replace(ACTCMS.ActCMS_Sys(11),"/",";")&"');"
					ListField=ListField&"so.addVariable('isGet','1');"& vbCrLf 
					ListField=ListField&"so.addVariable('args','myid=Upload;ModeID="&ModeID&";U='+U+"";""+';P='+P+"";""+'Yname="&RSObj("FieldName")&"');"& vbCrLf 
					ListField=ListField&"so.addVariable('upUrl','"&ACTCMS.ACTSYS&"User/Upload.asp');"& vbCrLf 
					ListField=ListField&"so.addVariable('fileName','Filedata');"& vbCrLf 
					ListField=ListField&"so.addVariable('maxNum','110');"& vbCrLf 
					ListField=ListField&"so.addVariable('maxSize','"&ACTCMS.ActCMS_Sys(10)/1024&"');"& vbCrLf 
					ListField=ListField&"so.addVariable('etmsg','1');"& vbCrLf 
					ListField=ListField&"so.addVariable('ltmsg','1');"& vbCrLf 
					ListField=ListField&"so.write(""sapload"&RSObj("FieldName")&""");"& vbCrLf 
					ListField=ListField&"// ]]>"& vbCrLf 
					ListField=ListField&"</script>"			& vbCrLf 	
					ListField=ListField&"<textarea rows=""10"" cols=""80"" name="""&RSObj("FieldName")&""" id="""&RSObj("FieldName")&""" ></textarea>"& vbCrLf
					ListField=ListField&"<script type=""text/javascript"" language=""JavaScript"">"& vbCrLf 
 					ListField=ListField&"CKEDITOR.replace( '"&RSObj("FieldName")&"',"& vbCrLf 
					ListField=ListField&"			{"& vbCrLf 
					ListField=ListField&"				skin : 'v2',height:""250px"", width:""100%"",toolbar:'Simple'"& vbCrLf 
					ListField=ListField&"			});"& vbCrLf 
 					ListField=ListField&"</script>"&IsNotNull
				End If 
		   Case "FileType"
				ListField= "<input  name='"&RSObj("FieldName")&"' type='text'  value='' size='40'><iframe src='../Upload_Admin.asp?ModeID=1&instr=1&instrname="&RSObj("FieldName")&"&YNContent=1&file=yes&amp;instrct=content' name='image' width='75%' height='25' scrolling='No' frameborder='0' id='image'></iframe>"&IsNotNull
		   Case "NumberType"
				ListField= "<input type='text' name='"&RSObj("FieldName")&"' size='"&RSObj("width")&"' value='"&RSObj("Type_Default")&"'>"&IsNotNull
		   Case "RadomType"
				ListField= "<input type='text' name='"&RSObj("FieldName")&"' size='25'  value='"&ACTCMS.MakeRandom(20)&"'>"&IsNotNull
		   Case "DownType"

						 ListField="<table  border='0'   cellpadding='3' cellspacing='1'  >"
						 ListField=ListField&  "<tr ><td width='12%'   ><b>设置下载数量：</b></td>"

						 ListField=ListField& "<td width='85%' colspan='3' ><input type='text' name='no' value='4' size='2'>&nbsp;&nbsp;<input 	"		
						 ListField=ListField& " type='button' name='button' class='act_btn' onclick='setid();' value='添加下载地址数'><font color='red'>"	
						 ListField=ListField& "如果选择了使用下载服务器，请在下面↓输入文件名称。</font>"
						 ListField=ListField& " <font color='blue'>下载服务器路径 + 下载文件名称 = 完整下载地址</font><br>"
						 ListField=ListField& "</td></tr><tr><td   ><b>下载地址：</b></td><td colspan='3' >"
						 ListField=ListField& " <select name='downid' size='1'>"
						 
						 
						 ListField=ListField& "<option value='1' selected>本地软件下载服务器</option><option value='0'>↓不使用下载服务器↓</option></select>"
						 ListField=ListField& " <input name='DownFileName' type='text' size='50' value='5434'>-<input name='DownText' type='text' size='15' value='下载地址2'> "
						 ListField=ListField& "<br> <span id='upid'></span>"



						 ListField=ListField& "</td> </tr>"
						 ListField=ListField& " </table>"


		   Case else
				ListField= "<font color=red>该字段错误</font>"
		   End Select 

 	End Function 

'==========================================
'函 数 名：FormAddForm()
'作    用：添加内容表单
'参    数：
'==========================================
Sub FormAddForm()
%>
<script language="javascript" type="text/javascript" src="http://image001.dgcloud01.qebang.cn/website/My97DatePicker/WdatePicker.js"></script>
<script type="text/javascript">
function time(n){
	if (n == 2){
		times1.style.display='none';
		times2.style.display='none';
	}
	else{
		times1.style.display='';
		times2.style.display='';
	}
}
</script>
<form id="FormAdd" name="FormAdd" method="post" action="FormMaker.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0"> 
    <tr>
        <td height="25" align="right" width="235">表单状态：</td>
        <td><input name="Fk_Form_Statu" type="radio" class="Input" id="Fk_Form_Statu" value="0" style="vertical-align:middle;border:0;" checked/><label for="Fk_Form_Statu" style="vertical-align:middle;">正常</label> <input name="Fk_Form_Statu" value="1" type="radio" class="Input" id="Fk_Form_Statu1" style="vertical-align:middle;border:0;"/><label for="Fk_Form_Statu1" style="vertical-align:middle;">关闭</label></td>
      </tr>
    <tr>
        <td height="25" align="right">表单名称：</td>
        <td><input name="Fk_Form_Name" type="text" class="Input" id="Fk_Form_Name" size="20"  style="vertical-align:middle;"/></td>
      </tr>
    <tr>
        <td height="25" align="right">数据表名称：</td>
        <td><input name="Fk_Form_tablename" type="text" class="Input" id="Fk_Form_tablename" size="20"  style="vertical-align:middle;"/>_Form</td>
      </tr>
	  
    <tr>
        <td height="25" align="right" style="color:red">将提交结果发送到站长信箱：</td>
        <td><input name="Fk_Form_SendMail" type="radio" class="Input" id="Fk_Form_SendMail" value="0"  style="vertical-align:middle;border:0;"/><label for="Fk_Form_SendMail" style="vertical-align:middle;">启用</label> <input name="Fk_Form_SendMail" value="1"  type="radio" class="Input" id="Fk_Form_SendMail1" style="vertical-align:middle;border:0;" checked/><label for="Fk_Form_SendMail1" style="vertical-align:middle;">不启用</label></td>
      </tr>
	  
    <tr>
        <td height="25" align="right">启用时间限制：</td>
        <td><input name="Fk_Form_TimeLimit" type="radio" value="0" class="Input" id="Fk_Form_TimeLimit" style="vertical-align:middle;border:0;" onclick="time(1)"/><label for="Fk_Form_TimeLimit" style="vertical-align:middle;">启用</label> <input name="Fk_Form_TimeLimit" type="radio" class="Input" id="Fk_Form_TimeLimit1" value="1" style="vertical-align:middle;border:0;" onclick="time(2)" checked/><label for="Fk_Form_TimeLimit1" style="vertical-align:middle;">不启用</label></td>
      </tr>
	  
	<tr id="times1" style="display:none">
		<td height="25" align="right">开始时间：  </td>
		<td height="25"><input id="StarTime" class="Input Wdate" type="text" onclick="WdatePicker()" readonly name="StarTime" size="30"></td>
	</tr>
	
	<tr id="times2" style="display:none">
		<td height="25" align="right">结束时间：  </td>
		<td height="25"><input id="EndTime" class="Input Wdate" type="text" onclick="WdatePicker()" readonly name="EndTime" size="30" ></td>
	</tr>
	  
	  
    <tr>
        <td height="25" align="right">表单描述：</td>
        <td><input name="Fk_Form_Describ" type="text" class="Input" id="Fk_Form_Describ" size="50"  style="vertical-align:middle;"/></td>
      </tr>
	  
    <tr>
        <td height="25" align="right">显示验证码：</td>
        <td><input name="Fk_Form_Code" type="radio" value="0" class="Input" id="Fk_Form_Code" style="vertical-align:middle;border:0;" checked/><label for="Fk_Form_Code" style="vertical-align:middle;">是</label> <input name="Fk_Form_Code" type="radio" value="1" class="Input" id="Fk_Form_Code1" style="vertical-align:middle;border:0;" /><label for="Fk_Form_Code1" style="vertical-align:middle;">否</label></td>
      </tr>
	  
   </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
        <input style="margin-left:248px" type="submit" onclick="Sends('FormAdd','FormMaker.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="btnClose" id="btnClose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FormAddDo
'作    用：执行添加内容
'参    数：
'==============================
Sub FormAddDo()
	Fk_Form_Statu=Trim(Request.Form("Fk_Form_Statu"))
	Fk_Form_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Form_Name")))
	Fk_Form_Tablename=Trim(Request.Form("Fk_Form_Tablename"))
	Fk_Form_SendMail=Trim(Request.Form("Fk_Form_SendMail"))
	Fk_Form_TimeLimit=Trim(Request.Form("Fk_Form_TimeLimit"))
	Fk_Form_StarTime=Trim(Request.Form("StarTime"))
	Fk_Form_EndTime=Trim(Request.Form("EndTime"))
	Fk_Form_Describ=FKFun.HTMLEncode(Trim(Request.Form("Fk_Form_Describ")))
	
	Fk_Form_Code=Trim(Request.Form("Fk_Form_Code"))
	
	Call FKFun.ShowString(Fk_Form_Name,1,100,0,"请输入表单名！","表单名不能大于100个字符！")
	Call FKFun.ShowString(Fk_Form_Tablename,1,20,0,"请输入数据表单名！","数据表单名不能大于20个字符！")
	If Fk_Form_TimeLimit="0" Then
		Call FKFun.ShowString(Fk_Form_StarTime,1,30,0,"请选择开始时间！","开始时间不能大于30个字符！")
		Call FKFun.ShowString(Fk_Form_EndTime,1,30,0,"请选择结束时间！","开始结束不能大于30个字符！")
	End If
	
	Sqlstr="Select * From [Ext_FormModel] Where Ext_Form_Name='"&Fk_Form_Name&"' or Ext_Table_Name='"&Fk_Form_Tablename&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Ext_Form_Name")=Fk_Form_Name
		Rs("Ext_Table_Name")=Fk_Form_Tablename
		If Fk_Form_TimeLimit="0" Then
			Rs("StartTime")=Fk_Form_StarTime
			Rs("EndTime")=Fk_Form_EndTime
		end if
		Rs("Ext_Form_Describ")=Fk_Form_Describ
		Rs("Ext_Form_Statu")=Fk_Form_Statu
		Rs("UnlockTime")=Fk_Form_TimeLimit
		Rs("FormCode")=Fk_Form_Code
		Rs("IsMail")=Fk_Form_SendMail
		Rs.Update()
		Application.UnLock()
		conn.execute("create table "&Fk_Form_Tablename&"_Form(id integer identity(1,1) primary key,UpdateTime date default now(),UserIP varchar(20) null,ViewStatu int default 0)")
		Response.Write("新表单添加成功！")
	Else
		Response.Write("该表单名或者数据表单名已经存在，请重新填写！")
	End If
	Rs.Close
End Sub

Function formatDate(Byval t,Byval ftype)
dim y, m, d, h, mi, s
formatDate=""
If IsDate(t)=False Then Exit Function
y=cstr(year(t))
m=cstr(month(t))
If len(m)=1 Then m="0" & m
d=cstr(day(t))
If len(d)=1 Then d="0" & d
h = cstr(hour(t))
If len(h)=1 Then h="0" & h
mi = cstr(minute(t))
If len(mi)=1 Then mi="0" & mi
s = cstr(second(t))
If len(s)=1 Then s="0" & s
select case cint(ftype)
case 1
' yyyy-mm-dd
formatDate=y & "-" & m & "-" & d
case 2
' yy-mm-dd
formatDate=right(y,2) & "-" & m & "-" & d
case 3
' mm-dd
formatDate=m & "-" & d
case 4
' yyyy-mm-dd hh:mm:ss
formatDate=y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
case 5
' hh:mm:ss
formatDate=h & ":" & mi & ":" & s
case 6
' yyyy年mm月dd日
formatDate=y & "年" & m & "月" & d & "日"
case 7
' yyyymmdd
formatDate=y & m & d
case 8
'yyyymmddhhmmss
formatDate=y & m & d & h & mi & s
end select
End Function 

'==========================================
'函 数 名：FormModelEditForm()
'作    用：修改内容表单
'参    数：
'==========================================
Sub FormModelEditForm()
	Id=Request.QueryString("Id")
	'判断权限
	Sqlstr="Select * From [Ext_FormModel] Where Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Form_Name=Rs("Ext_Form_Name")
		Fk_Form_Tablename=Rs("Ext_Table_Name")
		Fk_Form_StarTime=Rs("StartTime")
		Fk_Form_EndTime=Rs("EndTime")
		Fk_Form_Describ=Rs("Ext_Form_Describ")
		Fk_Form_TimeLimit=Rs("UnlockTime")		
		Fk_Form_Statu=Rs("Ext_Form_Statu")
		Fk_Form_Code=Rs("FormCode")		
		Fk_Form_SendMail=Rs("IsMail")
	End If
	Rs.Close
%>
<script language="javascript" type="text/javascript" src="http://image001.dgcloud01.qebang.cn/website/My97DatePicker/WdatePicker.js"></script>
<script type="text/javascript">
$(function(){
	if($("#Fk_Form_TimeLimit:checked")){
		times1.style.display='';
		times2.style.display='';
	}
})
function time(n){
	if (n == 2){
		times1.style.display='none';
		times2.style.display='none';
	}
	else{
		times1.style.display='';
		times2.style.display='';
	}
}
</script>
<form id="ArticleEdit" name="ArticleEdit" method="post" action="FormMaker.asp?Type=5" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">

    <tr>
        <td height="25" align="right" width="235">表单状态：</td>
        <td><input name="Fk_Form_Statu" type="radio" class="Input" id="Fk_Form_Statu" style="vertical-align:middle;border:0;" value="0" <%if Fk_Form_Statu=0 then response.write "checked"%>/><label for="Fk_Form_Statu" style="vertical-align:middle;">正常</label> <input name="Fk_Form_Statu" type="radio" class="Input" value="1" id="Fk_Form_Statu1" style="vertical-align:middle;border:0;" <%if Fk_Form_Statu=1 then response.write "checked"%>/><label for="Fk_Form_Statu1" style="vertical-align:middle;">关闭</label></td>
      </tr>
    <tr>
        <td height="25" align="right">表单名称：</td>
        <td><input name="Fk_Form_Name" type="text" class="Input" id="Fk_Form_Name" size="20" value="<%=Fk_Form_Name%>" style="vertical-align:middle;"/></td>
      </tr>
    <tr>
        <td height="25" align="right">数据表名称：</td>
        <td><input name="Fk_Form_tablename" type="text" class="Input" id="Fk_Form_tablename" value="<%=Fk_Form_Tablename%>_Form" size="20"  style="vertical-align:middle;" disabled="true"/></td>
      </tr>
	  
    <tr>
        <td height="25" align="right" style="color:red">将提交结果发送到站长信箱：</td>
        <td><input name="Fk_Form_SendMail" type="radio" class="Input" id="Fk_Form_SendMail" value="0" style="vertical-align:middle;border:0;" <%if Fk_Form_SendMail=0 then response.write "checked"%>/><label for="Fk_Form_SendMail" style="vertical-align:middle;">启用</label> <input name="Fk_Form_SendMail" type="radio" class="Input" id="Fk_Form_SendMail1" style="vertical-align:middle;border:0;" value="1" <%if Fk_Form_SendMail=1 then response.write "checked"%>/><label for="Fk_Form_SendMail1" style="vertical-align:middle;">不启用</label></td>
      </tr>
	  
    <tr>
        <td height="25" align="right">启用时间限制：</td>
        <td><input name="Fk_Form_TimeLimit" type="radio" value="0" class="Input" id="Fk_Form_TimeLimit" <%if Fk_Form_TimeLimit=0 then response.write "checked"%> style="vertical-align:middle;border:0;" onclick="time(1)"/><label for="Fk_Form_TimeLimit" style="vertical-align:middle;">启用</label> <input name="Fk_Form_TimeLimit" type="radio" class="Input" <%if Fk_Form_TimeLimit=1 then response.write "checked"%> id="Fk_Form_TimeLimit1" value="1" style="vertical-align:middle;border:0;" onclick="time(2)"/><label for="Fk_Form_TimeLimit1" style="vertical-align:middle;">不启用</label></td>
      </tr>
	  
	<tr id="times1" style="display:none">
		<td height="25" align="right">开始时间：</td>
		<td height="25"><input id="StarTime" class="Input Wdate" type="text" onclick="WdatePicker();" value="<%=formatDate(Fk_Form_StarTime,1)%>" name="StarTime" size="30"></td>
	</tr>
	
	<tr id="times2" style="display:none">
		<td height="25" align="right">结束时间：</td>
		<td height="25"><input id="EndTime" class="Input Wdate" type="text" onclick="WdatePicker();" value="<%=formatDate(Fk_Form_EndTime,1)%>" name="EndTime" size="30"></td>
	</tr>
	  
    <tr>
        <td height="25" align="right">表单描述：</td>
        <td><input name="Fk_Form_Describ" type="text" class="Input" id="Fk_Form_Describ" size="50"  value="<%=Fk_Form_Describ%>" style="vertical-align:middle;"/></td>
      </tr>
	  
    <tr>
        <td height="25" align="right">显示验证码：</td>
        <td><input name="Fk_Form_Code" type="radio" value="0" class="Input" id="Fk_Form_Code" style="vertical-align:middle;border:0;"  <%if Fk_Form_Code=0 then response.write "checked"%>/><label for="Fk_Form_Code" style="vertical-align:middle;">是</label> <input name="Fk_Form_Code" type="radio" value="1" class="Input" id="Fk_Form_Code1" style="vertical-align:middle;border:0;" <%if Fk_Form_Code=1 then response.write "checked"%>/><label for="Fk_Form_Code1" style="vertical-align:middle;">否</label></td>
      </tr>
       
</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input style="margin-left:248px" type="submit" onclick="Sends('ArticleEdit','FormMaker.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="btnClose" id="btnClose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FormModelEditDo
'作    用：执行修改内容
'参    数：
'==============================
Sub FormModelEditDo()
	id=Trim(Request.Form("id"))
	Fk_Form_Statu=Trim(Request.Form("Fk_Form_Statu"))
	Fk_Form_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Form_Name")))
	Fk_Form_Tablename=Trim(Request.Form("Fk_Form_Tablename"))
	Fk_Form_SendMail=Trim(Request.Form("Fk_Form_SendMail"))
	Fk_Form_TimeLimit=Trim(Request.Form("Fk_Form_TimeLimit"))
	Fk_Form_StarTime=Trim(Request.Form("StarTime"))
	Fk_Form_EndTime=Trim(Request.Form("EndTime"))
	Fk_Form_Describ=FKFun.HTMLEncode(Trim(Request.Form("Fk_Form_Describ")))
	Fk_Form_Code=Trim(Request.Form("Fk_Form_Code"))
	
	Call FKFun.ShowString(Fk_Form_Name,1,100,0,"请输入表单名！","表单名不能大于100个字符！")
	If Fk_Form_TimeLimit="0" Then
		Call FKFun.ShowString(Fk_Form_StarTime,1,30,0,"请选择开始时间！","开始时间不能大于30个字符！")
		Call FKFun.ShowString(Fk_Form_EndTime,1,30,0,"请选择结束时间！","开始结束不能大于30个字符！")
	End If
	
	
	Sqlstr="Select * From [Ext_FormModel] Where id="&id&""
	Rs.Open Sqlstr,Conn,1,3
	If not Rs.Eof Then
		dim Sqlstr1,rs1
		Sqlstr1="Select * From [Ext_FormModel] Where id<>"&id&" and Ext_Form_Name='"&Fk_Form_Name&"'"
		' response.write sqlstr
		' response.end
		set rs1=conn.execute(Sqlstr1)
		if rs1.eof then
			Application.Lock()
			Rs("Ext_Form_Name")=Fk_Form_Name
			If Fk_Form_TimeLimit="0" Then
				Rs("StartTime")=Fk_Form_StarTime
				Rs("EndTime")=Fk_Form_EndTime
			end if
			Rs("Ext_Form_Describ")=Fk_Form_Describ
			Rs("Ext_Form_Statu")=Fk_Form_Statu
			Rs("UnlockTime")=Fk_Form_TimeLimit
			Rs("FormCode")=Fk_Form_Code
			Rs("IsMail")=Fk_Form_SendMail
			Rs.Update()
			Application.UnLock()
			Response.Write("“"&Fk_Form_Name&"”修改成功！")
		else
			Response.Write("表单名重复，请修改后提交！")
		end if
		rs1.close
	Else
		Response.Write("表单不存在！")
	End If
	rs.close
End Sub

'==============================
'函 数 名：FormModelDelDo
'作    用：执行删除内容
'参    数：
'==============================
Sub FormModelDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Ext_FormModel] Where Id=" & Id
	on error resume next
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		conn.execute("drop table "&rs("Ext_Table_Name")&"_Form")
		conn.execute("delete * from Ext_Table_fields where FormID="&Id)
		Application.Lock()
		Rs.Delete()
		Application.UnLock()		
		Response.Write("表单删除成功！")
	Else
		Response.Write("表单不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ListDelDo
'作    用：执行批量删除内容
'参    数：
'==============================
Sub ListDelDo()
	Id=Replace(Trim(Request.Form("ListId"))," ","")
	If Id="" Then
		Response.Write("请选择要删除的内容！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="delete * From [Ext_FormModel] Where Id In ("&Id&")"
	conn.execute(Sqlstr)
	Response.Write("批量删除表单成功！")
End Sub

'==========================================
'函 数 名：FormFieldList()
'作    用：字段列表
'参    数：
'==========================================
Sub FormFieldList()
	dim tbname,FormId
	Session("NowPage")=FkFun.GetNowUrl()
	FormId=Trim(Request.QueryString("Formid"))
	tbname=Trim(Request.QueryString("tb"))
	on error resume next
	Sqlstr="Select * From [Ext_Table_fields] where FormID="&FormId&" order by orderID desc,id desc"
	Rs.Open Sqlstr,Conn,1,1
	if err then
		conn.execute("create table [Ext_Table_fields](id integer identity(1,1) primary key,FormID int null,[FieldName] varchar(50),[Title] varchar(50),[IsNotNull] int default 0,[OrderID] int default 0,[Description] varchar(250),[FieldType] varchar(50),[Type_Default] text,[width] int default 0,[height] int default 0,[Content] text,[Type_Type] int default 0,[ISType] int default 0,[regEx] varchar(255),[regError] varchar(255),[SearchIF] int default 0,[ValueOnly] int,[check] int default 3)")
		err.clear
		Rs.Open Sqlstr,Conn,1,1
	end if
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<a class="no1" href="javascript:void(0);" onclick="ShowBox('FormMaker.asp?Type=9&id=<%=FormId%>');">添加字段</a><a class="shuax" href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a>
    </div>
    <form name="DelList" id="DelList" method="post" action="FormMaker.asp?Type=7" onsubmit="return false;">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop" width="140">字段名称</th>
            <th align="center" class="ListTdTop">字段别名</th>
            <th align="center" class="ListTdTop">调用标签</th>
            <th align="center" class="ListTdTop">字段类型</th>
            <th align="center" class="ListTdTop">是否必填</th>
            <th align="center" class="ListTdTop">排序</th>
            <th align="center" class="ListTdTop" width="140">操作</th>
        </tr>
<%
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
%>
         <tr>
    <td style="padding-left:20px;"><%= Rs("FieldName") %></td>
    <td align="center"><%= Rs("Title") %></td>
	<td align="center"><font color=red>{$<%= Rs("FieldName") %>} </font></td>
    <td align="center"><%
	
		 Select Case Rs("FieldType")
		   Case "TextType"
				Response.write "单行文本"
		   Case "MultipleTextType"
				Response.write "多行文本(不支持Html)"
		   Case "MultipleHtmlType"
				Response.write "多行文本(支持Html)"
		   Case "RadioType"
				Response.write "单选项"
		   Case "ListBoxType"
				Response.write "多选项"
		   Case "DateType"
				Response.write "日期时间"
		   Case "PicType"
				Response.write "图片"
		   Case "FileType"
				Response.write "文件"
		   Case "NumberType"
				Response.write "数字"
		   Case "RadomType"
				Response.write "随机数"
		   Case else
				Response.write "<font color=red>该字段错误</font>"
		 End Select
%></td>
    <td align="center"><%if Rs("IsNotNull")=0 Then Response.Write "是" Else Response.Write "否" %></td>
 	<td align="center">
          <input name="OrderID" type="text"  class="Input"  id="OrderID" value="<%=rs("OrderID")%>" size="4" maxlength="3" style="border:1px solid #ccc; padding:0 5px; width:30px">
          <input name="A" type="hidden" id="A" value="N">
		  <input name="ID" type="hidden" id="ID" value="<%=rs("ID")%>">	
           
	</td>
	<td align="center"><a style="line-height:21px; width:auto;" href="javascript:void(0);" onclick="ShowBox('FormMaker.asp?Type=10&id=<%=Rs("id")%>&tbname=<%=tbname%>','修改字段','1000px','500px');">修改</a> ┆ 

<a style="line-height:21px; width:auto;" href="javascript:void(0);" onClick="DelIt('您确认要删除该字段吗，此操作不可逆！','FormMaker.asp?Type=13&Id=<%=Rs("id")%>&tbname=<%=tbname%>&FieldName=<%=Rs("FieldName")%>','MainRight','<%=Session("NowPage")%>');return false;">删除</a>  </td>
  </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="7">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="vertical-align:middle; margin-left:19px">&nbsp;<label for="chkall" style="vertical-align:middle;">全选</label>
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除选中的内容吗？')){Sends('DelList','FormMaker.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}" style="vertical-align:middle;">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("FormMaker.asp?Type=1&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="7" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
    </form>
</div>
<div id="ListBottom">
<script language="JavaScript" type="text/javascript">
	function overColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg1"
		Obj.bgColor="";
	}
	
}
function outColor(Obj)
{
	var elements=Obj.childNodes;
	for(var i=0;i<elements.length;i++)
	{
		elements[i].className="tdbg";
		Obj.bgColor="";
	}
}
</SCRIPT>
</div>
<%
End Sub


Sub FormFieldAddEdit()
	dim ModeName
	id=Trim(Request.QueryString("id"))
	ModeName=Request.QueryString("tbname")
	Dim A
	if Types=9 Then
		A="11"
		TitleSize="40"
		ISType="1"
		RadioPic_Type="0":check=3
		RadioType_Content="名称-值"& vbCrLf &"名称-值"& vbCrLf &"名称-值"
		ListBoxType_Content="名称-值"& vbCrLf &"名称-值"& vbCrLf &"名称-值"
		SupportHtmlType_Width=620
		SupportHtmlType_heigh=400
		MultipleTextType_Width=300
		MultipleTextType_Height=100
		IsEditor="Simple"
		IsNotNull=1
		OrderID=10
		YHtml=0
		savepic=0
	Elseif Types=10 Then
		RS.OPen "Select * from Ext_Table_fields Where ID = "&ID&" order by ID desc",Conn,1,1
		FieldName = RS("FieldName")
		FieldType = RS("FieldType")  
		ModeID = RS("FormID") 
		Title = RS("Title")  
		Type_Default = RS("Type_Default") 
		Description = RS("Description") 
		OrderID = RS("OrderID") 
		Select Case FieldType
			Case "TextType"'单行文本
				TitleSize=RS("Width") 
			Case "MultipleTextType"'多行文本(不支持Html
				MultipleTextType_Width=RS("Width") 
				MultipleTextType_Height=RS("Height")
				YHtml=RS("Type_Type")
 			Case "MultipleHtmlType"'多行文本(支持Html)
				IsEditor=RS("Content")
				SupportHtmlType_Width=RS("Width") 
				SupportHtmlType_Heigh=RS("Height")
				savepic=RS("Type_Type")
			Case "RadioType"'单选项
				RadioType_Content=RS("Content")
				RadioType_Type = RS("Type_Type")
			case "PicType"
			    RadioPic_Type = RS("Type_Type")
 			Case "ListBoxType"'多选项
				ListBoxType_Content=RS("Content")
				ListBoxType_Type = RS("Type_Type")
			Case "NumberType"'数字
				TitleSize=RS("Width") 
		End Select 
		ISType=RS("ISType")
		IsNotNull=RS("IsNotNull")
	    check=RS("check")

		if  RS("check")="1" then 
 		fun=RS("regex")
		else
 		regex=RS("regex")
		end if 
		regError=RS("regError")
		SearchIF=RS("SearchIF")
	    ValueOnly=RS("ValueOnly")

		A="12"
 	End  If 
    %>

	<form  name="FormFieldAdd" id="FormFieldAdd" method="post" action="FormMaker.asp?Type=<%=A%>" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table class="table" cellspacing="0" cellpadding="0" width="100%" border="0" align="center">
  <tr >
    <td width="15%" align="right"> 字段别名： </td>
    <td><input name="Title" value="<%=Title%>"  type="text"  class="Input"  maxlength="20" id="Title" />
    <font color="#ff0066">*</font> 如：文章标题</td>
  </tr>
  <tr >
    <td align="right" width="140"> 字段名称： </td>
    <td><input name="FieldName" value="<%=FieldName%>" <%If Types=10 Then Response.write " disabled=""disabled"" "%>  type="text"  class="Input"  maxlength="50" id="FieldName" />
       
	   <%If Types=10 Then Response.write " 该字段在 【"&ModeName&"】 模型内容页调用标签 <font color=red>{$"&FieldName&"}</font>"%> <br> &nbsp;&nbsp;&nbsp;&nbsp;<font color="#ff0066">*</font>为了和系统字段区分，系统创建字段时会自动以“_Cust”结尾,在模板中可以通过“{$字段名称_Cust}”进行调用 </td>
  </tr>
  <tr >
    <td align="right"> 字段描述： </td>
    <td style="padding:10px 0 10px 10px"><textarea name="Description" rows="6" cols="40" id="Description"><%=Description%></textarea>
	</td>
  </tr>
  <tr >


  <td align="right"> 是否必填： </td>
    <td style="padding-left:10px;"><table id="IsNotNull" border="0">
      <tr>
        <td style="border:0"><input id="IsNotNull_0" <% IF IsNotNull = "0" Then Response.Write "Checked" %> type="radio" name="IsNotNull" value="0" style="vertical-align:center;"/><label for="IsNotNull_0" style="vertical-align:center;">是</label></td>
        <td style="padding-left:10px; border:0;"><input id="IsNotNull_1" <% IF IsNotNull = "1" Then Response.Write "Checked" %> type="radio" name="IsNotNull" value="1" style="vertical-align:center;"/><label for="IsNotNull_1" style="vertical-align:center;">否</label>
			  </td>
      </tr>
    </table></td>
  </tr>
<!--   <tr >
    <td align="right"> 会员中心调用：</td>
    <td><table id="ISType"  border="0">
      <tr>
        <td><input id="ISType_0" <% IF ISType = "1" Then Response.Write "Checked" %>   type="radio" name="ISType" value="1" /><label for="ISType_0">是</label></td>
        <td><input id="ISType_1" type="radio" <% IF ISType = "0" Then Response.Write "Checked" %> name="ISType" value="0" /><label for="ISType_1">否</label>
			  
			  
			  </td>
      </tr>
    </table></td>
  </tr> -->
   <tr >
    <td align="right"> 字段排序： </td>
    <td><input name="OrderID" value="<%=OrderID%>"  type="text"  class="Input"  maxlength="20" id="OrderID" />
   数字越大,排的越前</td>
  </tr>
 
 
 
   <tr >
    <td align="right">数据校验规则：</td>
    <td style="padding-left:10px;">
      <input type="radio" name="check"   <% IF check = "3" Then Response.Write "Checked" %>  id="check3"  onClick=chk(3)  value="3" style="vertical-align:center;">
      <label for="check3" style="vertical-align:center;">默认 </label>
    <input type="radio" name="check"  <% IF check = "1" Then Response.Write "Checked" %>  id="check1"  onClick=chk(1)  value="1" style="vertical-align:center;">
      <label for="check1" style="vertical-align:center;">函数 </label>
        <input type="radio" name="check"  <% IF check = "2" Then Response.Write "Checked" %>  id="check2" onClick=chk(2)  value="2" style="vertical-align:center;">
      <label for="check2" style="vertical-align:center;">正则</label></td>
  </tr> 
 
 
   <tr id="checks3">
    <td align="right">函数名称：</td>
    <td><input name="fun" type="text"  class="Input"  id="fun" size="40" value="<%= fun %>">请在Field.asp中自己增加
	
  	</td>
  </tr> 
  
   <tr id="checks1">
    <td align="right">数据校验正则：</td>
    <td><input name="regEx" type="text"  class="Input"  id="regEx" size="40" value="<%= regEx %>">
 	<select name="select"   onchange="document.FormFieldAdd.regEx.value=this.value">
          <option selected>-- 常用正则 --</option>
			<option value="^[A-Za-z]+$">英文</option> 
			<option value="^[\u0391-\uFFE5]+$">中文</option> 
<!--			<option value="^[a-z]\w{2,19}$">中英文</option> 
-->			<option value="^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$">email</option> 
			<option value="^[-\+]?\d+$">整型</option> 
          <option value="^\d+$">数字</option> 
			<option value="^[-\+]?\d+(\.\d+)?$">double</option> 
			<option value="^[1-9]\d{4,9}$">qq</option> 
			<option value="^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$">phone</option> 
			<option value="^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$">mobile</option> 
			<option value="^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\])*$">网址</option> 
			<option value="^[A-Za-z0-9\-]+\.([A-Za-z]{2,4}|[A-Za-z]{2,4}\.[A-Za-z]{2})$">域名</option> 
			<option value="^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$">ip</option> 
               </select>
 	</td>
  </tr>
     <tr   id="checks2">
    <td align="right">检验错误提示信息：</td>
    <td><input name="regError" type="text"  class="Input"  id="regError" size="40"  value="<%= regError %>">
	
	
	
	</td>
  </tr>



  <tr >
    <td align="right">作为搜索条件：</td>
    <td style="padding-left:10px;">
	
<input id="SearchIF_0" <% IF SearchIF = "0" Then Response.Write "Checked" %>   type="radio" name="SearchIF" value="0"  style="vertical-align:center;"/>
<label for="SearchIF_0" style="vertical-align:center;">是</label>	

<input id="SearchIF_1" <% IF SearchIF = "1" Then Response.Write "Checked" %>   type="radio" name="SearchIF" value="1"  style="vertical-align:center;"/>
<label for="SearchIF_1" style="vertical-align:center;">否</label>	
	
	</td>
  </tr>


  <tr >
    <td align="right">数据唯一：</td>
    <td style="padding-left:10px;">
	
<input id="ValueOnly_0" <% IF ValueOnly = "0" Then Response.Write "Checked" %>   type="radio" name="ValueOnly" value="0"  style="vertical-align:center;"/>
<label for="ValueOnly_0" style="vertical-align:center;">是</label>	

<input id="ValueOnly_1" <% IF ValueOnly = "1" Then Response.Write "Checked" %>   type="radio" name="ValueOnly" value="1"  style="vertical-align:center;"/>
<label for="ValueOnly_1" style="vertical-align:center;">否</label>	
	
	</td>
  </tr>

  
  
  <tr >
    <td align="right"> 字段类型： </td>
    <td style="padding-left:10px;">	
	
<table id="FieldType" <%If Types=10 Then Response.write " disabled=""disabled"" "%> onClick="SelectModelType()" border="0">
	<tr>
		<td style="padding:0; border:0;"><input id="Type_0" <% IF FieldType = "TextType" Then Response.Write "Checked" %> type="radio" name="FieldType" value="TextType" checked="checked"  style="vertical-align:center;"/>
		<label for="Type_0" style="vertical-align:center;">单行文本</label></td>
		<td style="padding:0; border:0;"><input id="Type_1" <% IF FieldType = "MultipleTextType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="MultipleTextType"  style="vertical-align:center;"/>
		<label for="Type_1" style="vertical-align:center;">多行文本(不支持Html)</label></td>
        <td style="padding:0; border:0;">
		<input id="Type_2" <% IF FieldType = "MultipleHtmlType" Then Response.Write "Checked" %>   type="radio" name="FieldType" value="MultipleHtmlType"  style="vertical-align:center;"/>
		<label for="Type_2" <% IF FieldType = "SupportHtmlType" Then Response.Write "Checked" %>  style="vertical-align:center;">多行文本(支持Html)</label></td>
        <td  style="padding:0; border:0;">
		<input id="Type_3" <% IF FieldType = "RadioType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="RadioType"  style="vertical-align:center;"/><label for="Type_3" style="vertical-align:center;">单选项</label>
		</td>
		<td style="padding:0; border:0;"><input id="Type_4" <% IF FieldType = "ListBoxType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="ListBoxType"  style="vertical-align:center;"/><label for="Type_4" style="vertical-align:center;">多选项</label></td>
	</tr><tr>
		<td style="padding:0; border:0;"><input id="Type_5" <% IF FieldType = "DateType" Then Response.Write "Checked" %>  type="radio" name="FieldType" value="DateType"  style="vertical-align:center;"/>
		<label for="Type_5" style="vertical-align:center;">日期时间</label></td>
        <td style="padding:0; border:0;">
		<input id="Type_6"  <% IF FieldType = "PicType" Then Response.Write "Checked" %> type="radio" name="FieldType" value="PicType"  style="vertical-align:center;"/>
		<label for="Type_6" style="vertical-align:center;">上传</label></td>
        <td style="padding:0; border:0;"><input <% IF FieldType = "FileType" Then Response.Write "Checked" %>  id="Type_7" type="radio" name="FieldType" value="FileType"  style="vertical-align:center;"/>
		<label for="Type_7" style="vertical-align:center;">下载</label></td>
        <td style="padding:0; border:0;"><input <% IF FieldType = "NumberType" Then Response.Write "Checked" %>  id="Type_8" type="radio" name="FieldType" value="NumberType"  style="vertical-align:center;"/>
		<label for="Type_8" style="vertical-align:center;">数字</label></td>
        <td style="padding:0; border:0;"><input <% IF FieldType = "RadomType" Then Response.Write "Checked" %>  id="Type_9" type="radio" name="FieldType" value="RadomType"  style="vertical-align:center;"/>
		<label for="Type_9" style="vertical-align:center;">随机数</label>
		</td>
	</tr>
</table>	</td>
  </tr>
  <tbody id="DivTextType">
    <tr>
      <td align="right">文本框长度：</td>
      <td><input name="TitleSize" value="<%=TitleSize%>"  type="text"  class="Input"   maxlength="4" size="10" id="TitleSize" /></td>
    </tr>
  </tbody>
  <tbody id="DivMultipleTextType" style="display:none">
    <tr>
      <td align="right">显示的宽度：</td>
      <td><input name="MultipleTextType_Width" type="text"  class="Input"  value="<%=MultipleTextType_Width%>" maxlength="4" size="10" id="MultipleTextType_Width" />
        px</td>
    </tr>
    <tr>
      <td align="right">显示的高度：</td>
      <td><input name="MultipleTextType_Height" type="text"  class="Input"  value="<%=MultipleTextType_Height%>" maxlength="4" size="10" id="MultipleTextType_Height" />
        px</td>
    </tr>
    
    <tr>
      <td align="right">是否允许HTML：</td>
      <td><input id="YHtml_0" <% IF YHtml = "0" Then Response.Write "Checked" %>   type="radio" name="YHtml" value="0"  style="vertical-align:center;"/>
<label for="YHtml_0" style="vertical-align:center;">是</label>	

<input id="YHtml_1" <% IF YHtml = "1" Then Response.Write "Checked" %>   type="radio" name="YHtml" value="1"  style="vertical-align:center;"/>
<label for="YHtml_1" style="vertical-align:center;">否</label>	

</td>
    </tr>    
    
  </tbody>
  <tbody id="DivMultipleHtmlType" style="display:none">
    <tr>
      <td align="right">编辑器菜单名称：</td>
      <td>       
<input name="IsEditor" type="text"  class="Input"  value="<%=IsEditor%>" maxlength="4" size="10" id="IsEditor" />
<select name="select" onChange="FormatTitle(this, FormFieldAdd.IsEditor, '')">
          <option selected>-- 请选择 --</option>
          <option value="1">简洁</option>
        </select></td>
    </tr>
    <tr>
      <td align="right">显示的宽度：</td>
      <td><input name="SupportHtmlType_Width" type="text"  class="Input"  value="<%=SupportHtmlType_Width%>" maxlength="4" size="10" id="SupportHtmlType_Width" />
        px</td>
    </tr>
    <tr>
      <td align="right">显示的高度：</td>
      <td><input name="SupportHtmlType_Heigh" type="text"  class="Input"  value="<%=SupportHtmlType_Heigh%>" maxlength="4" size="10" id="SupportHtmlType_Heigh" />
        px</td>
    </tr>
    
    <tr>
      <td align="right">是否保存远程图片：</td>
      <td><input id="savepic_0" <% IF savepic = "0" Then Response.Write "Checked" %>   type="radio" name="savepic" value="0"  style="vertical-align:center;"/>
<label for="savepic_0" style="vertical-align:center;">是</label>	

<input id="savepic_1" <% IF savepic = "1" Then Response.Write "Checked" %>   type="radio" name="savepic" value="1"  style="vertical-align:center;"/>
<label for="savepic_1" style="vertical-align:center;">否</label>	</td>
    </tr>      
  </tbody>
  <tbody id="DivRadioType" style="display:none">
    <tr>
      <td align="right">分行键入每个选项：</td>
      <td><textarea name="RadioType_Content" rows="6" cols="40" id="RadioType_Content"><%=RadioType_Content%></textarea>
	  <font color=red>注意 要按照格式书写 名称-值, 以 - 隔开,列:合肥-HeFei</font></td></tr>
    <tr>
      <td align="right">显示选项：</td>
      <td><table id="RadioType_Type" border="0">
        <tr>
          <td>
		  <input id="RadioType_Type_0"  <% IF RadioType_Type = "0" Then Response.Write "Checked" %>  type="radio" name="RadioType_Type" value="0" checked="checked"  style="vertical-align:center;"/>
                <label for="RadioType_Type_0" style="vertical-align:center;">单选下拉列表框</label></td>
        </tr>
        <tr>
          <td><input id="RadioType_Type_1" <% IF RadioType_Type = "1" Then Response.Write "Checked" %>  type="radio" name="RadioType_Type" value="1"  style="vertical-align:center;"/>
                <label for="RadioType_Type_1" style="vertical-align:center;">单选按钮</label></td>
        </tr>
      </table></td>
    </tr>
  </tbody>
  <tbody id="DivListBoxType" style="display:none">
    <tr>
      <td align="right">分行键入每个选项：</td>
      <td><textarea name="ListBoxType_Content" rows="6" cols="40" id="ListBoxType_Content"><%=ListBoxType_Content%></textarea></td></tr>
    <tr>
      <td align="right">显示选项：</td>
      <td><table id="ListBoxType_Type" border="0">
        <tr>
          <td><input id="ListBoxType_Type_0"  <% IF ListBoxType_Type = "0" Then Response.Write "Checked" %>  type="radio" name="ListBoxType_Type" value="0" checked="checked"  style="vertical-align:center;"/>
                <label for="ListBoxType_Type_0" style="vertical-align:center;">复选框</label></td>
        </tr>
        <tr>
          <td><input id="ListBoxType_Type_1"  <% IF ListBoxType_Type = "1" Then Response.Write "Checked" %>  type="radio" name="ListBoxType_Type" value="1" style="vertical-align:center;" />
                <label for="ListBoxType_Type_1" style="vertical-align:center;">多选列表框</label></td>
        </tr>
      </table></td>
    </tr>
  </tbody>
  <tbody id="DivDateType" style="display:none">
  </tbody>
  <tbody id="DivPicType" style="display:none">
    <tr>
      <td align="right">显示选项：</td>
      <td><table id="RadioPic_Type" border="0">
        <tr>
          <td>
		  
     <input id="RadioPic_Type_0"    type="radio" name="RadioPic_Type" value="0"  <% IF RadioPic_Type = "0"  Then Response.Write "Checked" %>     style="vertical-align:center;"/>
                <label for="RadioPic_Type_0" style="vertical-align:center;">单个文件上传</label></td>
        </tr>
        <tr>
          <td>
          <input id="RadioPic_Type_1"   type="radio" name="RadioPic_Type" value="1"   <% IF RadioPic_Type = "1" Then Response.Write "Checked" %>   style="vertical-align:center;"/>
                <label for="RadioPic_Type_1" style="vertical-align:center;">多个文件上传</label></td>
        </tr>
      </table></td>
    </tr>  
  
  
  </tbody>
  <tbody id="DivRadomType" style="display: none">
  </tbody>
  <tbody id="DivFileType" style="display:none">
  </tbody>
  <tbody id="DivNumberType" style="display:none">
    <tr>
      <td align="right">文本框长度：</td>
      <td><input name="NumberType_TitleSize" type="text"  class="Input"  value="40" maxlength="4" size="10" id="NumberType_TitleSize" /></td>
    </tr>
  </tbody>
  <tr>
	<Td align="right">默认值：</td>
      <td><input name="type_Default" type="text"  class="Input"  value="<%=type_Default%>" size="10" id="NType_Default" />
	  
	  注：没有数据录入的默认值，与前台显示无关.</td>
  </tr>

</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
		<input type="hidden" name="ModeID" value="<%=ModeID%>" />
		<input type="hidden" name="ID" value="<%=ID%>" />
        <input style="margin-left:150px;" type="submit" onclick="if(CheckForm()){Sends('FormFieldAdd','FormMaker.asp?Type=<%=A%>',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}" class="Button" name="button" id="button" value="<%if Types=9 then response.write "添 加":else:response.write "修 改"%>" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="btnClose" id="btnClose" value="关 闭" />
</div>
	</form>
<script language="JavaScript" type="text/javascript">
function FormatTitle(obj, obj2, def_value)
{
    var FormatFlag = obj.options[obj.selectedIndex].value;
    var tmp_Title = FilterHtmlStr(obj2.value);
    switch(FormatFlag)
    {
        case "1" :
            obj2.value = "UserMode";
            break;
        case "2" :
            obj2.value = "Simple";
            break;
        case "3" :
            obj2.value = "Default";
            break;
    }
    obj.selectedIndex = 0;
}
function FilterHtmlStr(str)
{
    str = str.replace(/<.*?>/ig, "");
    return str;
}
function SelectModelType()
{
    var TypeCount=document.getElementsByName("FieldType"); 
    
    for(var i=1;i<TypeCount.length;i++)
    { 
        var DivType=eval("Div"+TypeCount[i].value);
        
        if(TypeCount[i].checked)
        {
            DivType.style.display="";
        }
        else
        {
            DivType.style.display="none";
        }
    }
}


function chk(n){
	if (n == "1"){
		checks1.style.display='none';
		checks2.style.display='none';
		checks3.style.display='';
 	}
	else  if (n == "2"){
		checks1.style.display='';
		checks2.style.display='';
		checks3.style.display='none';
  	}
	else
	{
		checks1.style.display='none';
		checks2.style.display='none';
		checks3.style.display='none';
 	}
}




	function CheckForm()
	{ 
		var form=document.FormFieldAdd;
		if (form.Title.value=='')
		{ 
			alert("字段别名不能够为空！");
			form.Title.focus();
			return false;
		}
		if (form.FieldName.value=='')
		{ 
			alert("字段名称不能够为空！!");
			form.FieldName.focus();
			return false;
		}
		return true;
	}	
</script>
    <script language="javascript">chk("<%= check %>");</script>


<%  End Sub

	Sub FormFieldAddSave()
		ID=request.form("ID")
		FieldType=request.form("FieldType")
		FieldName=request.form("FieldName")
		Title=request.form("Title")
		IsNotNull=request.form("IsNotNull")
		ISType=request.form("ISType")
		OrderID=request.form("OrderID")
		Type_Default=request.form("Type_Default")
		Description=request.form("Description")
		savepic=request.form("savepic")
		YHtml=request.form("YHtml")
		ModeID=request.form("ModeId")
		
		Call FKFun.ShowString(Title,1,50,0,"请输入字段别名！","字段别名不能大于50个字符！")
		Call FKFun.ShowString(FieldName,1,50,0,"请输入字段名称！","字段名称不能大于50个字符！")
		FieldName=FieldName&"_Cust"
		
		Dim ActMode_Width,ActMode_Height
		'长度.宽度.
		Select Case FieldType
			Case "TextType"'单行文本
				ActMode_Width =  request.form("TitleSize")'文本框长度
				ColumnType="varchar(255)"
				Call FKFun.ShowNum(ActMode_Width,"文本框长度只能是数字！")
			Case "MultipleTextType"'多行文本(不支持Html
				ActMode_Width =  request.form("MultipleTextType_Width")
				ActMode_Height =  request.form("MultipleTextType_Height")
				Type_Type=YHtml
				 ColumnType="text"
				Call FKFun.ShowNum(ActMode_Width,"文本框长度只能是数字！")
				Call FKFun.ShowNum(ActMode_Height,"文本框高度只能是数字！")
			Case "MultipleHtmlType"'多行文本(支持Html)
				Content = request.form("IsEditor")'编辑器属性放入Content字段
				ActMode_Width =  request.form("SupportHtmlType_Width")
				ActMode_Height =  request.form("SupportHtmlType_Heigh")
				Type_Type=savepic
				ColumnType="text"
				Call FKFun.ShowNum(ActMode_Width,"文本框长度只能是数字！")
				Call FKFun.ShowNum(ActMode_Height,"文本框高度只能是数字！")
			Case "RadioType"'单选项
			    Content = request.form("RadioType_Content")
				Type_Type =  request.form("RadioType_Type")'显示方式
				ColumnType="varchar(255)"
			Case "ListBoxType"'多选项
			    Content = request.form("ListBoxType_Content")
				Type_Type =  request.form("ListBoxType_Type")
				ColumnType="text"
			Case "NumberType"'数字
			    ActMode_Width =  request.form("NumberType_TitleSize")'数字的宽度放入总宽度字段名称中
				Call FKFun.ShowNum(ActMode_Width,"文本框长度只能是数字！")
				ColumnType="int"'
		   Case "DateType"
				 ColumnType="datetime"'Response.write "日期时间"
		   Case "NumberType"
				ColumnType="int"'Response.write "数字"
		  case "PicType"
				 Type_Type =request.form("RadioPic_Type")
				 ColumnType="text"

		   Case else
		     ColumnType="varchar(255)"
		End Select 
		set rs=conn.execute("select Ext_Table_Name from Ext_FormModel where id="& ID)
		if not rs.eof then
			TableName=rs("Ext_Table_Name")
		else
			rs.close
			response.write "参数错误！"
			response.end
		end if
		rs.close
		 sqlstr = "Select * From [Ext_Table_fields] Where  FieldName='" & FieldName & "' And FormID=" & ID
		 ' response.write sqlstr
		 ' response.end
		 RS.Open sqlstr, conn, 1, 3
 		 If RS.EOF And RS.BOF Then
			RS.AddNew
			RS("FieldName") = FieldName
			RS("FieldType") = FieldType
			RS("FormID") = ID
			RS("Title") = Title
			RS("IsNotNull") = IsNotNull
			RS("Width") = ActMode_Width
			RS("Height") = ActMode_Height
			RS("Type_Default") = Type_Default
			RS("Description") = Description
			RS("Type_Type") = Type_Type
			RS("Content") = Content
			RS("ISType") = ISType
			RS("OrderID") = OrderID
 			RS("check") = request.form("check")
			if request.form("check")	="1" then 
			RS("regex") = request.form("fun")
			else 
			RS("regex") = request.form("regex")
			RS("regError") = request.form("regError")
			end if 
 			RS("SearchIF") = request.form("SearchIF")
		    RS("ValueOnly") = request.form("ValueOnly")	
		    RS.Update
			 Conn.Execute("Alter Table "&TableName&"_Form Add "&FieldName&" "&ColumnType&"")
			 Response.Write "字段增加成功!"
		 Else
		   response.write("数据库中已存在该字段名称!")
		   Exit Sub
		 End If
	End  Sub 


	Sub FormFieldEditSave()
		ID=request.form("ID")
		Title=request.form("Title")
		IsNotNull=request.form("IsNotNull")
		Type_Default=request.form("Type_Default")
		Description=request.form("Description")
		OrderID=request.form("OrderID")
		savepic=request.form("savepic")
		YHtml=request.form("YHtml")
		Call FKFun.ShowString(Title,1,50,0,"请输入字段别名！","字段别名不能大于50个字符！")
		If TitleSize=0 Then TitleSize=40
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		 sqlstr = "Select * From [Ext_Table_fields] Where ID=" & ID
		 RS.Open sqlstr, conn,1, 3
			RS("Title") = Title
			RS("IsNotNull") = IsNotNull
			Select Case RS("FieldType")
				Case "TextType"'单行文本
					RS("Width") =  request.form("TitleSize")'文本框长度
				Case "MultipleTextType"'多行文本(不支持Html
					RS("Width") =  request.form("MultipleTextType_Width")
					RS("Height") =  request.form("MultipleTextType_Height")
					RS("Type_Type") =YHtml
				Case "MultipleHtmlType"'多行文本(支持Html)
					RS("Content") = request.form("IsEditor")'编辑器属性放入Content字段
					RS("Width") =  request.form("SupportHtmlType_Width")
					RS("Height") =  request.form("SupportHtmlType_Heigh")
					RS("Type_Type") =savepic
				Case "RadioType"'单选项
					RS("Content") = request.form("RadioType_Content")
					RS("Type_Type") =  request.form("RadioType_Type")'显示方式
				Case "ListBoxType"'多选项
					RS("Content") = request.form("ListBoxType_Content")
					RS("Type_Type") =  request.form("ListBoxType_Type")	
				Case "NumberType"'数字
					RS("Width") =  request.form("NumberType_TitleSize")'数字的宽度放入总宽度字段名称中
				case "PicType"
					 RS("Type_Type") = request.form("RadioPic_Type")
 			End Select 
			RS("Description") = Description
			RS("Type_Default") = Type_Default
			RS("ISType")= request.form("ISType")
			RS("OrderID") = OrderID
 			RS("check") = request.form("check")	
			if request.form("check")	="1" then 
			RS("regex") = request.form("fun")
			else 
			RS("regex") = request.form("regex")
			RS("regError") = request.form("regError")
			end if 
			RS("SearchIF") = request.form("SearchIF")
 		    RS("ValueOnly") = request.form("ValueOnly")	

			 RS.Update
			 response.write("字段修改成功")
 	End  Sub

sub FormFieldDel()
	dim tbname
	ID=request.querystring("Id")
	tbname=request.querystring("tbname")
	FieldName=request.querystring("FieldName")
	on error resume next
	conn.execute("alter table "&tbname&"_Form drop column "&FieldName&"")
	conn.execute("delete * from Ext_Table_fields where id= "&ID&"")
	response.write("字段删除成功")
end sub

%><!--#Include File="../Code.asp"-->