<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_HTML.asp"-->
<%
'==========================================
'文 件 名：Product.asp
'文件用途：图文管理拉取页面
'版权所有：深圳企帮
'==========================================

'定义页面变量
Dim Fk_Product_Title,Fk_Product_Content,Fk_Product_Click,Fk_Product_Show,Fk_Product_Time,Fk_Product_Pic,Fk_Product_PicBig,Fk_Product_Template,Fk_Product_FileName,Fk_Product_Recommend,Fk_Product_Subject,Fk_Product_Keyword,Fk_Product_Description,Fk_Product_Color,Fk_Product_Url,Fk_Product_Field,Fk_Product_onTop,Fk_Product_px,Fk_Product_Seotitle
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Dir,Fk_Product_Module
Dim Temp2,KeyWordlist,kwdrs,ki

' On Error Resume next
	dim KeyWorddat,krs
	Sqlstr="select SVkeywords from [keywordSV]"
	set krs=conn.execute(Sqlstr)
	If Not krs.Eof Then
		dim m
		m=0
		KeyWorddat=""
		do while Not krs.Eof
			if m=0 then
				KeyWorddat=FilterText(krs("SVkeywords"))
			else
				KeyWorddat=KeyWorddat&"|"&FilterText(krs("SVkeywords"))
			end if
			m=m+1
			krs.movenext
			if krs.eof then exit do
		loop
	else
		KeyWorddat=""
	end if
	krs.close
if FKFso.IsFile("KeyWordC.dat") then
	KeyWordlist=FKFso.FsoFileRead("KeyWordC.dat")
else
	KeyWordlist=KeyWorddat
	call FKFso.CreateFile("KeyWordC.dat",KeyWorddat)
end if

dim chkContentEx1,chkContentEx2,chkFk_Product_wapimg,chkSlidesImgs,chkSummary,chkSlidesFirst,Fk_Product_wapimg
chkContentEx1=CheckFields("Fk_Product_ContentEx1","fk_product")
chkContentEx2=CheckFields("Fk_Product_ContentEx2","fk_product")
chkSlidesImgs=CheckFields("Fk_Product_SlidesImgs","fk_product")
chkFk_Product_wapimg=CheckFields("Fk_Product_wapimg","fk_product")
chkSummary=CheckFields("Fk_Product_Summary","fk_product")
chkSlidesFirst=CheckFields("Fk_Product_SlidesFirst","fk_product")
'===================================== 
'过滤字符 
'===================================== 
Function FilterText(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
FilterText=t0 
End Function 

Function CheckFields(FieldsName,TableName)
	dim blnFlag,chkStrSql,chkStrRs
	blnFlag=False
	chkStrSql="select * from "&TableName
	Set chkStrRs=Conn.Execute(chkStrSql)
	for i = 0 to chkStrRs.Fields.Count - 1
		if lcase(chkStrRs.Fields(i).Name)=lcase(FieldsName) then
			blnFlag=True
			Exit For
		else
			blnFlag=False
		end if
	Next
	CheckFields=blnFlag
End Function

	Function getDesc(strTableName, strColName) 
		Dim cat 
		Set cat = server.CreateObject( "ADOX.Catalog") 
		cat.ActiveConnection = conn
		getDesc = cat.Tables(strTableName).Columns(strColName).Properties( "Description").Value 
		Set cat = Nothing 
	End Function 

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call ProductList() '图文列表
	Case 2
		Call ProductAddForm() '添加图文表单
	Case 3
		Call ProductAddDo() '执行添加图文
	Case 4
		Call ProductEditForm() '修改图文表单
	Case 5
		Call ProductEditDo() '执行修改图文
	Case 6
		Call ProductDelDo() '执行删除图文
	Case 7
		Call ListDelDo() '执行批量删除图文
	Case 8
		Call ProductMove() '执行批量移动图文
	Case 9
		Call ProductOrderDo() '执行图文内容批量排序
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：ProductList()
'作    用：图文列表
'参    数：
'==========================================
Sub ProductList()
	'新功能，追加SEO title字段
	'2017年5月22日
	'middy241@163.com
	if CheckFields("Fk_Product_Seotitle","Fk_Product")=false then
		conn.execute("alter table Fk_Product add column Fk_Product_Seotitle varchar(255) null")
	end if
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
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
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
	End If
	Rs.Close
%>

<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3><%=Fk_Module_Name%>栏目</h3><input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle"/><input type="button" class="Button" onclick="SetRContent('MainRight','Product.asp?Type=1&ModuleId=<%=Fk_Module_Id%>&SearchStr='+escape(document.all.SearchStr.value));" name="S" Id="S" value="  查询  " style="vertical-align:middle" />
        <h3>请选择栏目</h3>
        <select name="D1" id="D1" onChange="window.execScript(this.options[this.selectedIndex].value);" style="vertical-align:middle">
            <option value="alert('请选择栏目');">请选择栏目</option>
            <%
            Call ModuleSelectUrl(Fk_Module_Menu,0,Fk_Module_Id)
            %>
        </select>
        <a class="tjia" href="javascript:void(0);" onclick="ShowBox('Product.asp?Type=2&ModuleId=<%=Fk_Module_Id%>','添加','1000px','500px');">添加</a>
        <a class="shuax" href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a>
    </div>
    <form name="DelList" id="DelList" method="post" action="Article.asp?Type=7" onsubmit="return false;">
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th width="80" align="center" class="ListTdTop">选</th>
            <th align="left" class="ListTdTop" width="300">标题</th>
            <th align="center" style="display:none" class="ListTdTop">文件名</th>
            <th align="left" class="ListTdTop">显示</th>
            <th align="center" class="ListTdTop">点击量</th>
            <th align="center" class="ListTdTop">排序</th>
            <th align="center" class="ListTdTop">转播微博</th>
            <th align="center" class="ListTdTop">添加时间</th>
            <th width="120" align="center" class="ListTdTop">操作</th>
        </tr>
<%
	Dim Rs2,ProductUrl,zfurl
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Fk_Product] Where Fk_Product_Module="&Fk_Module_Id&""
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And Fk_Product_Title Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By Fk_Product_Ip desc,Px desc, Fk_Product_Time Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Dim ProductTemplate
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
			If Rs("Fk_Product_Template")>0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & Rs("Fk_Product_Template")
				Rs2.Open Sqlstr,Conn,1,1
				If Not Rs2.Eof Then
					Fk_Product_Template=Rs2("Fk_Template_Name")
				Else
					Fk_Product_Template="未知模板"
				End If
				Rs2.Close
			Else
				Fk_Product_Template="默认模板"
			End If
			Fk_Product_Recommend=""
			If Rs("Fk_Product_Recommend")<>"" Then
				TempArr=Split(Rs("Fk_Product_Recommend"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Product_Recommend=Fk_Product_Recommend&","&Rs2("Fk_Recommend_Name")
						End If
						Rs2.Close
					End If
				Next
			End If
			Fk_Product_Subject=""
			If Rs("Fk_Product_Subject")<>"" Then
				TempArr=Split(Rs("Fk_Product_Subject"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Product_Subject=Fk_Product_Subject&","&Rs2("Fk_Subject_Name")
						End If
						Rs2.Close
					End If
				Next
			End If
			
			zfurl="http://"&Request.ServerVariables("Server_name")
			If Rs("Fk_Product_Url")<>"" Then
				ProductUrl=Rs("Fk_Product_Url")
				if instr(ProductUrl,"http://") then
					zfurl=ProductUrl
				else
					zfurl=zfurl&ProductUrl
				end if
			Else
				If Fk_Module_Dir<>"" Then
					ProductUrl=Fk_Module_Dir&"/"
				Else
					ProductUrl="Product"&Fk_Module_Id&"/"
				End If
				If Rs("Fk_Product_FileName")<>"" Then
					ProductUrl=ProductUrl&Rs("Fk_Product_FileName")&".html"
				Else
					ProductUrl=ProductUrl&Rs("Fk_Product_Id")&".html"
				End If
				If SiteHtml=1 and sitetemplate<>"wap" Then
					ProductUrl="/html"&SiteDir&ProductUrl
				Else
					ProductUrl=SiteDir&sTemp&"?"&ProductUrl
				End If
				zfurl=zfurl&ProductUrl
			End If
%>
        <tr>
            <td height="20" align="left" style="padding-left:33px;"><input type="hidden" value="<%=Rs("Fk_Product_Id")%>" name="id[]"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("Fk_Product_Id")%>" id="List<%=Rs("Fk_Product_Id")%>" /></td>
            <td class="td1"><span style="line-height:22px; padding-right:15px"><%=Rs("Fk_Product_Title")%><%If Rs("Fk_Product_Color")<>"" Then%><span style="color:<%=Rs("Fk_Product_Color")%>">■</span><%End If%><%If Rs("Fk_Product_Url")<>"" Then%>[转向链接]<%End If%></span></td>
            <td align="center" style="display:none"><%=Rs("Fk_Product_FileName")%></td>
            <td align="left"><%If Rs("Fk_Product_Show")=1 Then%><span class="gnszxianshi"></span><%Else%><span class="gnszxianshi hidden"></span><%End If%><%If Rs("Fk_Product_Pic")<>"" Then%><span style="color:#F00">[图]</span><%End If%><a style="display:none; line-height:21px; width:auto;" href="javascript:void(0);" title="<%=Fk_Product_Template%> ">[模]</a><%If InStr(Fk_Product_Recommend,"推荐")>0 Then%><a style="line-height:21px; width:auto;" href="javascript:void(0);" title="<%=Replace(Fk_Product_Recommend,",","")%> ">[推]</a><%End If%><%If trim(Rs("Fk_Product_Ip")&" ")="1" Then%><a style=" line-height:21px; width:auto;" href="javascript:void(0);" title="置顶 "><font color=blue style="font-weight:bolder;">[顶]</font></a><%End If%><%If Fk_Product_Subject<>"" Then%><a style=" line-height:21px; width:auto;" href="javascript:void(0);" title="<%=Replace(Fk_Product_Subject,",","")%> ">[专]</a><%End If%></td>
            <td align="center"><%=Rs("Fk_Product_Click")%></td>
            <td height="20" align="center"><input type="text" value="<%=Rs("Px")%>" name="Fk_Product_px[]" style="width:25px;text-align:center"/></td>
            <td align="center">
            <a href="http://v.t.qq.com/share/share.php?url=<%=zfurl%>&appkey=3016f499d2714531819b4c5e5ec10cc1&site=&title=<%=server.URLEncode(Rs("Fk_Product_Title")&"("&Rs("Fk_Product_Keyword")&")")%>&pic=<%
if left(Rs("Fk_Product_Pic"),4)<>"http" then
	response.write "http://"&Request.ServerVariables("Server_name")&Rs("Fk_Product_Pic")
else
	response.write Rs("Fk_Product_Pic")
end if
			%>" target="_blank"><img style="cursor:pointer;vertical-align:middle" alt="转发到腾讯微博 " src="Images/weiboicon16.png" ></a>&nbsp;
            <a href="http://service.weibo.com/share/share.php?url=<%=zfurl%>&appkey=1525536596&title=<%=server.URLEncode(Rs("Fk_Product_Title")&"("&Rs("Fk_Product_Keyword")&")")%>&pic=<%if left(Rs("Fk_Product_Pic"),4)<>"http" then
	response.write "http://"&Request.ServerVariables("Server_name")&Rs("Fk_Product_Pic")
else
	response.write Rs("Fk_Product_Pic")
end if
			%>&ralateUid=" target="_blank"><img style="cursor:pointer;vertical-align:middle" alt="转发到新浪微博 " src="Images/weiboicon16-sina.png" ></a></td>
            <td align="center"><%=Rs("Fk_Product_Time")%></td>
            <td align="left"  class="no6" width="120"><a class="no2" title="修改 " href="javascript:void(0);" onclick="ShowBox('Product.asp?Type=4&ModuleId=<%=Fk_Module_Id%>&Id=<%=Rs("Fk_Product_Id")%>','修改','1000px','500px');"></a> <a class="no4" title="删除 " href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Product_Title")%>”，此操作不可逆！','Product.asp?Type=6&Id=<%=Rs("Fk_Product_Id")%>','MainRight','<%=Session("NowPage")%>');"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="7" style="padding-left:33px;" class="NowPage">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="vertical-align:middle">&nbsp;&nbsp;<label for="chkall">全选</label>&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="submit" value="删 除" class="Button" onClick="DelIt('确定要删除选中的图文吗？','Product.asp?Type=7&ListId='+GetCheckbox(),'MainRight','<%=Session("NowPage")%>');" style="vertical-align:middle">
			<input type="submit" value="更新排序" class="Button" onClick="Sends('DelList','Product.asp?Type=9',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style="vertical-align:middle;">
<select name="ProductMove" id="ProductMove" onchange="DelIt('确实要移动这部分图文？','Product.asp?Type=8&Id='+this.options[this.options.selectedIndex].value+'&ListId='+GetCheckbox(),'MainRight','<%=Session("NowPage")%>');" style="vertical-align:middle">
      <option value="">转移到</option>
<%
Call ModuleSelectId(Fk_Module_Menu,0,0)
%>
</select>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("Product.asp?Type=1&ModuleId="&Fk_Module_Id&"&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="8" align="center">暂无记录</td>
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
End Sub

'==========================================
'函 数 名：ProductAddForm()
'作    用：添加图文表单
'参    数：
'==========================================
Sub ProductAddForm()
	dim rnd_num
	RANDOMIZE
	rnd_num=INT(100*RND)+1
	on error resume next
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	End If
	Rs.Close
%>
    <script type="text/javascript" charset="utf-8" id="colorpickerjs" src="/js/colorpicker.js"></script>
	<link  rel="stylesheet" type="text/css" id="mutiImgcss"  href="ext/css/mutiImg.css"/>
    <script type="text/javascript" charset="utf-8" id="extMutiImgjs" src="ext/js/extMutiImg.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$("#colorpickerjs").attr("src","/js/colorpicker.js?r="+Math.random());
	$("#mutiImgcss").attr("href","ext/css/mutiImg.css?r="+Math.random());
	$("#extMutiImgjs").attr("src","ext/js/extMutiImg.js?r="+Math.random());

	$("#Fk_Product_Video_Div").hide();
	
	if(window.KindEditor){
	$("#Fk_Product_Video_File").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于mp4格式,文件最大允许5M");
		var editor = window.KindEditor.editor({
				fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
				uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp',
				allowFileManager : true
			});
			$('#uploadButton').click(function() {
				editor.loadPlugin('insertfile', function() {
					editor.plugin.fileDialog({
						fileUrl : $('#Fk_Product_Video_File').val(),
						clickFn : function(url) {
							$('#Fk_Product_Video_File').val(url);
							$('#Fk_Product_Video_Div').show();
							document.getElementById("Fk_Product_Video_Tag").src=url;
							// document.getElementById("Fk_Product_Video_Tag").play();
							editor.hideDialog();
						}
					});
				});
			});

	}
	else
	{
		$("#Fk_Product_Video_File").after(" <iframe frameborder=\"0\" width=\"200\" height=\"25\" scrolling=\"No\" id=\"Fk_Product_Video_Files\" name=\"Fk_Product_Video_Files\" src=\"PicUpLoad.asp?Form=ProductAdd&Input=Fk_Product_Video_File\" style=\"vertical-align:middle\"></iframe> *上传限于mp4格式,文件最大允许3M");
	}
})
</script>


<form id="ProductAdd" name="ProductAdd" method="post" action="Product.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">

	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">标题：</td>
        <td><input name="Fk_Product_Title"<%If SiteToPinyin=1 Then%> onchange="GetPinyin('Fk_Product_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> type="text" class="Input" id="Fk_Product_Title" size="50"  style="vertical-align:middle"/><input type="hidden" id="Fk_Product_Color" name="Fk_Product_Color" value="" />
<span class="colorpicker" onclick="colorpicker('colorpanel_title','set_title_color');" title="标题颜色" style="vertical-align:middle"></span>
                           <span class="colorpanel" id="colorpanel_title" style="position:absolute;z-index:99999999;"></span>
                           <script type="text/javascript">
						   set_title_color("");
                           </script> </td>
		<td rowspan="7" style="vertical-align:top;">
		
	  <%if chkSlidesImgs then%>
		
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tbody><tr>
                          <td style="width:260px;border-bottom:0" valign="top">
                            <div class="uploadimgArea" style="margin-left:8px;">
                               <div class="top">
                                 <span class="fr sysBtn" style="margin:5px 0px 0px 0px;"><a id="MutiImg" class="Button" style="cursor:pointer;display: inline-block;width: 90px;"><b class="icon icon_add"></b>添加图片</a></span>
                                 多图片(最多允许10张)</div>
                               <div id="bigimg" class="bigimg"><table border="0" cellpadding="0" cellspacing="0"><tr><td style="border-bottom:0;padding:0"></td></tr></table></div>
                               <ul class="imgslist" id="imgslist">
                               </ul>
                            </div>                          </td></tr>
                        </tbody>
                        </table>
	  <%end if%></td>
    </tr>
	<tr><td align="right">文件名：</td>
	<td><input name="Fk_Product_FileName" type="text" class="Input" id="Fk_Product_FileName" size="20"  style="vertical-align:middle"/>&nbsp;*不可修改</td>
		</tr>
	<tr><td align="right">排序：</td>
	<td><input class="Input" type="text" value="0" name="Fk_Product_px" id="Fk_Product_px" size="4" maxlength="6" style="vertical-align:middle"/> （仅限数字，越大越排前）</td>
		</tr>
	<tr>
		<td height="28" align="right">SEO标题：</td>
		<td><input name="Fk_Product_Seotitle" value="" type="text" class="Input" id="Fk_Product_Seotitle" size="60"  style="vertical-align:middle;"/></td>
	</tr>
    <tr>
        <td height="28" align="right">SEO关键词：</td>
        <td><input name="Fk_Product_Keyword" type="text" class="Input" id="Fk_Product_Keyword" size="40"  style="vertical-align:middle"/>
		<input type="button" onclick="tiqu(0,'Fk_Product_Content','Fk_Product_Keyword');" class="Button" name="btntqkwd" id="btntqkwd" value="提 取"  style="vertical-align:middle"/>
		<input  value="<%=KeyWordlist%>" type="hidden" id="Fk_Keywordlist" /></td>
	  </tr>
    <tr>
        <td height="28" align="right">SEO描述：</td>
        <td><input name="Fk_Product_Description" type="text" class="Input" id="Fk_Product_Description" size="40"  style="vertical-align:middle"/>
        <input type="button" onclick="tiqu(1,'Fk_Product_Content','Fk_Product_Description');" class="Button" name="btntqdesc" id="btntqdesc" value="提 取"  style="vertical-align:middle"/></td>
	  </tr>
     <tr>
        <td height="28" align="right" valign="top">缩略图：</td>
        <td>
		<div class="image_grid" style="padding:10px;">
                              <table border="0" cellpadding="0" cellspacing="0"><tbody><tr>
                              <td valign="top" style="padding:0px;border-bottom:0px;">
                                  <div class="imageBox" id="ThumbnailBox">
                                  <img class="img" id="SlImg" src="http://image001.dgcloud01.qebang.cn/website/ext/images/image.jpg">                                  </div>                              </td>
                              <td valign="top" style="padding:0px 0px 0px 6px;border-bottom:0px;">
                                  <input type="text" class="Input" id="Fk_Product_Pic" name="Fk_Product_Pic" style="width:314px; " value="">
                                  <div style="padding:5px 0px 0px 10px;">
                                  <span class="sysBtn" id="UploadImage"><a id="chooseImg" style="cursor:pointer;display: inline-block;width: 100px;" class="Button"><b class="icon icon_add"></b>添加图片...</a></span>                                  </div>
                                  <label class="Normal" id="d_Thumbnail"></label>                              </td>
                              </tr></tbody></table>
                          </div>						  </td>
      </tr>
	  
	<tr>
		<td height="28" align="right">主图视频：</td>
		<td colspan="2" style="padding-left:10px;"><div id="Fk_Product_Video_Div" style="position:relative;margin-top:10px;width:320px;"><video controls="controls" id="Fk_Product_Video_Tag" width="320" src=""></video><span id="removeVideo" title="移除视频" style="position:absolute;right:10px;top:10px;width: 20px;height: 20px;background: #000;display: flex;justify-content: center;align-items: center;border-radius: 10px;color: #fff;cursor: pointer;">X</span></div><input name="Fk_Product_Video_File" type="hidden" class="Input" id="Fk_Product_Video_File" size="50"  style="vertical-align:middle;"/></td>
	</tr>
	  <%if chkSummary then%>
    <tr>
        <td height="28" align="right">内容摘要：</td>
        <td style="padding:10px;" colspan="2"><textarea style="border:1px solid #ccc; padding:5px; font-size:12px; color:#555;" name="Fk_Product_Summary"  style="border: 1px solid #D3E3F0;width:300px;padding-top:5px;padding-left:5px;" id="Fk_Product_Summary" rows="5" cols="50" disabled></textarea>
		<div style="width:413px;"><input type="checkbox" id="GetAbstract" name="GetAbstract" value="1" checked="checked" style="vertical-align:middle;"> <label for="GetAbstract" style="vertical-align:middle;">采用产品详细的前<%=SiteMini%>个字作为摘要</label></div></td>
	  </tr>
	  <%end if%>
     <tr>
        <td height="28" align="right" >内容：</td>
        <td colspan="2" style="padding:10px;">
		<div class="dt-detail">
                      <div class="tabbar" id="J_TabBar">
                          <ul><li class="current"><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',1);return false;"><span>产品详细</span></a></li>                              
                              <li <%if not chkContentEx1 then response.write "style='display:none;'"%>><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',2);return false;"><span><%=getDesc("fk_product","Fk_Product_ContentEx1")%></span></a></li>
                              <li <%if not chkContentEx2 then response.write "style='display:none;'"%>><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',3);return false;"><span><%=getDesc("fk_product","Fk_Product_ContentEx2")%></span></a></li>
                              <li <%if not chkFk_Product_wapimg then response.write "style='display:none;'"%>><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',4);return false;"><span>小程序商品详情</span></a></li>
                              <li><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',5);return false;"><span>资料下载</span></a></li>
                              <li><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',6);return false;"><span>测试视频</span></a></li>
                          </ul>
                      </div>
                      <!--详细内容-->
                      <div class="product-infoitem" id="TabBar_1">
                          <textarea name="Fk_Product_Content" class="<%=Bianjiqi%>" id="Fk_Product_Content" rows="15" style="width:100%;"></textarea>                          
                      </div>
                      <!--详细内容:end-->
                      
                      <div class="product-infoitem" id="TabBar_2" style="display:none;">
						  <textarea name="Fk_Product_ContentEx1" class="<%=Bianjiqi%>" id="Fk_Product_ContentEx1" rows="15" style="width:100%;"></textarea>
                      </div>
                      
                      <div class="product-infoitem" id="TabBar_3" style="display:none;">
						  <textarea name="Fk_Product_ContentEx2" class="<%=Bianjiqi%>" id="Fk_Product_ContentEx2" rows="15" style="width:100%;"></textarea>
                      </div>
					  
                      <div class="product-infoitem" id="TabBar_4" style="display:none;">
						  <textarea name="Fk_Product_wapimg" class="<%=Bianjiqi%>" id="Fk_Product_wapimg" rows="15" style="width:100%;"></textarea>
                      </div>
	  
                      <div class="product-infoitem" id="TabBar_5" style="display:none;">
						  <div class="related_down_div">
						  <a style="color: #f06100;text-decoration:underline;" href="javascript:void(0);" onclick="OpenBoxNew('Down.asp?Type=10','关联资料下载','800px','500px', '关联选中资料', $('input[name=downId]').map(function(){return $(this).val();}).get(), 'related_down', 'downId');">关联资料下载</a>
							<ul id="related_down">
							</ul>
						  
						  </div>
                      </div>
	  
                      <div class="product-infoitem" id="TabBar_6" style="display:none;">
						  <div class="related_down_div">
						  <a style="color: #f06100;text-decoration:underline;" href="javascript:void(0);" onclick="OpenBoxNew('Article.asp?Type=10','关联测试视频','800px','500px', '关联选中视频', $('input[name=videoId]').map(function(){return $(this).val();}).get(), 'related_video', 'videoId');">关联测试视频</a>
							<ul id="related_video">
							</ul>
						  
						  </div>
                      </div>
                      
                  </div>
		</td>
    </tr>
  <%
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="2"><input name="Fk_Product__<%=Rs("Fk_Field_Tag")%>" type="text" class="Input" id="Fk_Product__<%=Rs("Fk_Field_Tag")%>"  style="vertical-align:middle"/></td>
    </tr>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
    <tr>
        <td align="right">转向链接：</td>
        <td colspan="2"><input name="Fk_Product_Url" type="text" class="Input" id="Fk_Product_Url" size="50"  style="vertical-align:middle"/>&nbsp;*正常请留空</td>
    </tr>
        <tr>
        <td height="28" align="right">推荐：</td>
        <td colspan="2"><select name="Fk_Product_Recommend" class="Input" id="Fk_Product_Recommend" style="vertical-align:middle">
            <option value="0">无推荐</option>
            <option value="2">推荐</option>
            </select>
			<input type="checkbox" name="Fk_Product_onTop" id="Fk_Product_onTop" class="textarea" value="1"  style="vertical-align:middle"/><label for="Fk_Product_onTop" style="vertical-align:middle">&nbsp;置顶</label>
          　<select name="Fk_Product_Subject" class="TextArea" id="Fk_Product_Subject" style="vertical-align:middle;display:none">
            <option value="0">无专题</option>
            </select>是否显示：<input name="Fk_Product_Show" type="radio" id="Fk_Product_Show" class="Input" value="1" checked="true"  style="vertical-align:middle"/><label for="Fk_Product_Show" style="vertical-align:middle">显示</label>
        <input type="radio" name="Fk_Product_Show" class="Input" id="Fk_Product_Show1" value="0"  style="vertical-align:middle"/><label for="Fk_Product_Show1" style="vertical-align:middle">不显示</label>　模板：<select name="Fk_Product_Template" class="Input" id="Fk_Product_Template" style="vertical-align:middle">
            <option value="0"<%=FKFun.BeSelect(Fk_Product_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"><%=Rs("Fk_Template_Name")%></option>
    <%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>            &nbsp; &nbsp; 点击量：<input style="width:50px" name="Fk_Product_click" type="text" id="Fk_Product_click" class="Input" size="6" maxlength="6" value="<%=rnd_num%>" style="vertical-align:middle;"></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto;" class="tcbtm">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('ProductAdd','Product.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="btnclose" id="btnclose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ProductAddDo
'作    用：执行添加图文
'参    数：
'==============================
Sub ProductAddDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	dim GetAbstract,Fk_Product_Summary,FK_Product_SlidesImgs,FK_Product_SlidesFirst,Fk_Product_ContentEx2,Fk_Product_ContentEx1 '2013-11-28 shark 添加
	
	GetAbstract=trim(Request.Form("GetAbstract")&" ")
	FK_Product_SlidesImgs=Trim(Request.Form("SlidesImg[]")&" ")
	FK_Product_SlidesFirst=Trim(Request.Form("SlidesImg[]firstImg")&" ")

	Fk_Product_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Title")))
	Fk_Product_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Color")))
	Fk_Product_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Seotitle")))
	Fk_Product_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Keyword")))
	Fk_Product_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Description")))
	Fk_Product_Content=Request.Form("Fk_Product_Content")
	Fk_Product_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Url")))
	Fk_Product_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Pic")))
	Fk_Product_PicBig=Fk_Product_Pic
	Fk_Product_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Recommend"))," ",""))&","
	Fk_Product_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Subject"))," ",""))&","
	Fk_Product_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_FileName")))
	Fk_Product_Template=Trim(Request.Form("Fk_Product_Template"))
	Fk_Product_Show=Trim(Request.Form("Fk_Product_Show"))
	Fk_Product_click=Trim(Request.Form("Fk_Product_click"))
	Fk_Product_onTop=Trim(Request.Form("Fk_Product_onTop"))
	Fk_Product_px=Trim(Request.Form("Fk_Product_px"))
	Call FKFun.ShowString(Fk_Product_Title,1,255,0,"请输入图文标题！","图文标题不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Product_Show,"排序只能是数字！请重新填写")
	Call FKFun.ShowString(Fk_Product_Seotitle,0,255,2,"请输入图文SEO标题！","图文SEO标题不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Keyword,0,255,2,"请输入图文SEO关键词！","图文SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Description,0,255,2,"请输入图文SEO描述！","图文SEO描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Url,0,255,2,"请输入图文转向链接！","图文转向链接不能大于255个字符！")
	
	if GetAbstract<>"1" then
		Fk_Product_Summary=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Summary")))
		Call FKFun.ShowString(Fk_Product_Summary,0,500,2,"请输入图文摘要！","图文摘要不能大于500个字符！")
	end if 
	Fk_Product_ContentEx1=Trim(Request.Form("Fk_Product_ContentEx1")&" ")
	Fk_Product_ContentEx2=Trim(Request.Form("Fk_Product_ContentEx2")&" ")
	Fk_Product_wapimg=Trim(Request.Form("Fk_Product_wapimg")&" ")
	
	If Fk_Product_Url="" Then
		Call FKFun.ShowString(Fk_Product_Content,20,1,1,"请输入图文内容，不少于20个字符！","图文内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Product_Pic,0,255,2,"请输入图文缩略图路径！","图文缩略图小图路径不能大于255个字符！")
	'Call FKFun.ShowString(Fk_Product_PicBig,0,255,2,"请输入图文缩略图路径！","图文缩略图大图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_FileName,0,50,2,"请输入图文文件名！","图文文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Product_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Product_Show,"请选择图文是否显示！")
	Call FKFun.ShowNum(Fk_Product_click,"请输入正确的点击量！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Fk_Product_Field="" Then
			Fk_Product_Field=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Product__"&Rs("Fk_Field_Tag"))))
		Else
			Fk_Product_Field=Fk_Product_Field&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Product__"&Rs("Fk_Field_Tag"))))
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If IsNumeric(Fk_Product_FileName) Then
		Response.Write("文件名不可用纯数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Left(Fk_Product_FileName,4)="Info" Or Left(Fk_Product_FileName,4)="Page" Or Left(Fk_Product_FileName,5)="GBook" Or Left(Fk_Product_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		Response.Write("栏目不存在！")
		Rs.Close
		Call FKDB.DB_Close()
		Response.End()
	End If
	Rs.Close
	If SiteDelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Product_Content=Replace(Fk_Product_Content,Temp,"**")
				Fk_Product_Title=Replace(Fk_Product_Title,Temp,"**")
				Fk_Product_Keyword=Replace(Fk_Product_Keyword,Temp,"**")
				Fk_Product_Description=Replace(Fk_Product_Description,Temp,"**")
			End If
		Next
	End If
	Sqlstr="Select * From [Fk_Product] Where Fk_Product_Module="&Fk_Module_Id&" And (Fk_Product_Title='"&Fk_Product_Title&"'"
	If Fk_Product_FileName<>"" Then
		Sqlstr=Sqlstr&" Or Fk_Product_FileName='"&Fk_Product_FileName&"'"
	End If
	Sqlstr=Sqlstr&")"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Product_Title")=Fk_Product_Title
		Rs("Fk_Product_Color")=Fk_Product_Color
		Rs("Fk_Product_Seotitle")=Fk_Product_Seotitle
		Rs("Fk_Product_Keyword")=Fk_Product_Keyword
		Rs("Fk_Product_Description")=Fk_Product_Description
		Rs("Fk_Product_Show")=Fk_Product_Show
		Rs("Fk_Product_click")=Fk_Product_click
		Rs("Fk_Product_Field")=trim(Fk_Product_Field&" ")
		Rs("Fk_Product_Pic")=Fk_Product_Pic
		Rs("Fk_Product_PicBig")=Fk_Product_PicBig
		'on error resume next
		if chkSummary then
			if GetAbstract<>"1" then
				if Fk_Product_Summary<>"" then
					Rs("Fk_Product_Summary")=trim(Fk_Product_Summary&" ")
				end if
			end if
		end if
		if chkSlidesFirst then
			if FK_Product_SlidesFirst<>"" then
				FK_Product_SlidesFirst=replace(FK_Product_SlidesFirst,"http://"&request.ServerVariables("HTTP_HOST"),"")
				Rs("FK_Product_SlidesFirst")=trim(FK_Product_SlidesFirst&" ")
			end if
		end if
		
		if chkSlidesImgs then
			if FK_Product_SlidesImgs<>"" then
				Rs("FK_Product_SlidesImgs")=trim(FK_Product_SlidesImgs&" ")
			end if
		end if
		
		if chkContentEx1 then
			if Fk_Product_ContentEx1<>"" then
				Rs("Fk_Product_ContentEx1")=Fk_Product_ContentEx1
			end if
		end if
		
		if chkContentEx2 then
			if Fk_Product_ContentEx2<>"" then
				Rs("Fk_Product_ContentEx2")=Fk_Product_ContentEx2
			end if
		end if
		
		if chkFk_Product_wapimg then
			if Fk_Product_wapimg<>"" then
				Rs("Fk_Product_wapimg")=Fk_Product_wapimg
			end if
		end if
	
		Rs("Fk_Product_Content")=Fk_Product_Content
		Rs("Fk_Product_Url")=Fk_Product_Url
		Rs("Fk_Product_Recommend")=Fk_Product_Recommend
		Rs("Fk_Product_Subject")=Fk_Product_Subject
		Rs("Fk_Product_Module")=Fk_Module_Id
		Rs("Fk_Product_Menu")=Fk_Module_Menu
		Rs("Fk_Product_FileName")=Fk_Product_FileName
		Rs("Fk_Product_Template")=Fk_Product_Template
		Rs("Fk_Product_Ip")=Fk_Product_onTop
		Rs("Px")=Fk_Product_px
		Rs.Update()
		Application.UnLock()
		Response.Write("新图文添加成功！")


		dim selectedDownIds, selectedVideoIds
		selectedDownIds = FKFun.HTMLEncode(Trim(Request.Form("downId")))
		selectedVideoIds = FKFun.HTMLEncode(Trim(Request.Form("videoId")))
		' 如果有关联资料，则修改Fk_Down表数据 
		if selectedDownIds <> "" then 
			conn.execute("update [Fk_Down] set Fk_Relation_Product_Id = " & Rs("Fk_Product_Id") & " Where Fk_Down_Id in ("& selectedDownIds &")")
		end if
		' 如果有关联测试视频，则修改Fk_Article表数据 
		if selectedVideoIds <> "" then 
			conn.execute("update [Fk_Article] set Fk_Relation_Product_Id = " & Rs("Fk_Product_Id") & " Where Fk_Article_Id in ("& selectedVideoIds &")")
		end if

		if err then 
			err.clear
		end if
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="添加图文：【"&Fk_Product_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
		
	Else
		Response.Write("该图文标题已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：ProductEditForm()
'作    用：修改图文表单
'参    数：
'==========================================
Sub ProductEditForm()
	'on error resume next
	Id=Clng(Request.QueryString("Id"))
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	dim SlidesImgs,FK_Product_SlidesFirst,Fk_Product_Summary,Fk_Product_ContentEx1,Fk_Product_ContentEx2
	dim Fk_Product_Video_File
	
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Id=Rs("Fk_Module_Id")
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Product] Where Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Dim Rs2, Arr_Rows, Related_Down_List, Sqlstr1, Sqlstr2, Rs3, Related_Down_Count, Arr_Rows_Video, Related_Video_Count
		
		Sqlstr1="Select * From [Fk_Down] Where Fk_Relation_Product_Id=" & Id
		Set Rs2 = Conn.execute(Sqlstr1)
		If Not Rs2.Eof Then
			Arr_Rows=Rs2.getrows()
			Related_Down_Count = Ubound(Arr_Rows, 2)+1
		Else
			Related_Down_Count = 0
		End If
		Rs2.close
		
		Sqlstr2="Select * From [Fk_Article] Where Fk_Relation_Product_Id=" & Id
		Set Rs3 = Conn.execute(Sqlstr2)
		If Not Rs3.Eof Then
			Arr_Rows_Video=Rs3.getrows()
			Related_Video_Count = Ubound(Arr_Rows_Video, 2)+1
		Else
			Related_Video_Count = 0
		End If
		Rs3.close
		Fk_Product_Title=Rs("Fk_Product_Title")
		Fk_Product_Color=Rs("Fk_Product_Color")
		Fk_Product_Seotitle=Rs("Fk_Product_Seotitle")
		Fk_Product_Keyword=Rs("Fk_Product_Keyword")
		Fk_Product_Description=Rs("Fk_Product_Description")
		Fk_Product_Content=Rs("Fk_Product_Content")
		Fk_Product_Url=Rs("Fk_Product_Url")
		Fk_Product_Pic=Rs("Fk_Product_Pic")
		Fk_Product_PicBig=Rs("Fk_Product_PicBig")
		Fk_Product_Show=Rs("Fk_Product_Show")
		Fk_Product_click=Rs("Fk_Product_click")
		Fk_Product_Template=Rs("Fk_Product_Template")
		Fk_Product_FileName=Rs("Fk_Product_FileName")
		Fk_Product_Recommend=Rs("Fk_Product_Recommend")
		Fk_Product_Subject=Rs("Fk_Product_Subject")
		Fk_Product_onTop=trim(Rs("Fk_Product_Ip")&" ")
		Fk_Product_px=Rs("Px")
		Fk_Product_Video_File=Rs("Fk_Product_Video_File")
		
		if chkSummary then
			Fk_Product_Summary=trim(Rs("Fk_Product_Summary")&" ")
		end if
		if chkSlidesImgs then
			SlidesImgs=trim(Rs("FK_Product_SlidesImgs")&" ")
		end if
		if chkSlidesFirst then
			FK_Product_SlidesFirst=trim(Rs("FK_Product_SlidesFirst")&" ")
		end if
		if chkContentEx1 then
			Fk_Product_ContentEx1=trim(Rs("Fk_Product_ContentEx1")&" ")
		end if
		if chkContentEx2 then
			Fk_Product_ContentEx2=trim(Rs("Fk_Product_ContentEx2")&" ")
		end if
		if chkFk_Product_wapimg then
			Fk_Product_wapimg=trim(Rs("Fk_Product_wapimg")&" ")
		end if
		
		If IsNull(Rs("Fk_Product_Field")) Or Rs("Fk_Product_Field")="" Then
			Fk_Product_Field=Split("-_-|-Fangka_Field-|1")
		Else
			Fk_Product_Field=Split(Rs("Fk_Product_Field"),"[-Fangka_Field-]")
		End If
	End If
	Rs.Close
%>
    
    <script type="text/javascript" charset="utf-8" id="colorpickerjs" src="/js/colorpicker.js"></script>
	<link  rel="stylesheet" type="text/css" id="mutiImgcss"  href="ext/css/mutiImg.css"/>
    <script type="text/javascript" charset="utf-8" id="extMutiImgjs" src="ext/js/extMutiImg.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$("#colorpickerjs").attr("src","/js/colorpicker.js?r="+Math.random());
	$("#mutiImgcss").attr("href","ext/css/mutiImg.css?r="+Math.random());
	$("#extMutiImgjs").attr("src","ext/js/extMutiImg.js?r="+Math.random());
	$("#removeVideo").on("click", function(){
		console.log(123);
		$('#Fk_Product_Video_File').val("");
		document.getElementById("Fk_Product_Video_Tag").src="";
		$('#Fk_Product_Video_Div').hide();
	})
	<%if Fk_Product_Video_File="" then %>
	$("#Fk_Product_Video_Div").hide();
	<%end if%>
	if(window.KindEditor){
		$("#Fk_Product_Video_File").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于mp4格式,文件最大允许5M");
			var editor = window.KindEditor.editor({
					fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
					uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp',
					allowFileManager : true
				});
				$('#uploadButton').click(function() {
					editor.loadPlugin('insertfile', function() {
						editor.plugin.fileDialog({
							fileUrl : $('#Fk_Product_Video_File').val(),
							clickFn : function(url) {
								$('#Fk_Product_Video_File').val(url);
								$('#Fk_Product_Video_Div').show();
								document.getElementById("Fk_Product_Video_Tag").src=url;
								// document.getElementById("Fk_Product_Video_Tag").play();
								editor.hideDialog();
							}
						});
					});
				});

		}
		else
		{
			$("#Fk_Product_Video_File").after(" <iframe frameborder=\"0\" width=\"200\" height=\"25\" scrolling=\"No\" id=\"Fk_Product_Video_Files\" name=\"Fk_Product_Video_Files\" src=\"PicUpLoad.asp?Form=ProductAdd&Input=Fk_Product_Video_File\" style=\"vertical-align:middle\"></iframe> *上传限于mp4格式,文件最大允许3M");
		}
})
</script>
	
<form id="ProductEdit" name="ProductEdit" method="post" action="Product.asp?Type=5" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="85" height="28" align="right" >标题：</td>
        <td><input name="Fk_Product_Title"<%If SiteToPinyin=1 Then%> onmouseout="GetPinyin('Fk_Product_FileName','ToPinyin.asp?Str='+this.value);" onchange="GetPinyin('Fk_Product_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> value="<%=Fk_Product_Title%>" type="text" class="Input" id="Fk_Product_Title" size="50"  style="vertical-align:middle"/><input type="hidden" id="Fk_Product_Color" name="Fk_Product_Color" value="" />
<span class="colorpicker" onclick="colorpicker('colorpanel_title','set_title_color');" title="标题颜色" style="vertical-align:middle"></span>
                           <span class="colorpanel" id="colorpanel_title" style="position:absolute;z-index:99999999;"></span>
                           <script type="text/javascript">
						   set_title_color("<%=Fk_Product_Color%>");
                           </script> <input type="checkbox" name="Fk_Product_Time" id="Fk_Product_Time" value="1" style="vertical-align:middle;"/><label style="vertical-align:middle;" for="Fk_Product_Time">更新时间</label></td>
						   <td rowspan="7" style="vertical-align:top;">
		
	  <%if  chkSlidesFirst and chkSlidesImgs then%>
		
          <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tbody><tr>
                          <td style="width:260px;border-bottom:0" valign="top">
                            <div class="uploadimgArea" style="margin-left:8px;">
                               <div class="top">
                                 <span class="fr sysBtn" style="margin:5px 0px 0px 0px;"><a id="MutiImg" style="cursor:pointer;display: inline-block;width:90px;" class="Button"><b class="icon icon_add"></b>添加图片</a></span>
                                 多图片(最多允许10张)</div>
                               <div id="bigimg" class="bigimg"><table border="0" cellpadding="0" cellspacing="0"><tr><td style="border-bottom:0;padding:0"><%if FK_Product_SlidesFirst<>"" and SlidesImgs<>"" then response.write "<input type=""hidden"" name=""SlidesImg[]FirstImg"" value="""&FK_Product_SlidesFirst&""" /><img src="""&FK_Product_SlidesFirst&"""/>"%></td></tr></table></div>
                               <ul class="imgslist" id="imgslist">
							    <% If not IsNull(SlidesImgs) Then
									dim TempImg,liclass
                                     SlidesImgs=Split(SlidesImgs,",")
                                     For i=0 to Ubound(SlidesImgs)
                                     TempImg=Trim(SlidesImgs(i))
									 'response.write TempImg&"-"&FK_Product_SlidesFirst&"<br>"
									 if TempImg=FK_Product_SlidesFirst then liclass="class='current'" : else : liclass="":end if
                                     response.write  "<li "&liclass&"><input type=""hidden"" name=""SlidesImg[]"" value="""&TempImg&""" /><table border=""0"" cellpadding=""0"" cellspacing=""0""><tbody><tr><td class=""imgdiv""><a href=""javascript:;"" target=""_self""><img src="""&TempImg&"""></a></td></tr></tbody></table><p><a onclick=""deleteCurrentPic(this)"" href=""javascript:;"" target=""_self""><b class=""icon icon_del"" title=""删除该图片""></b></a><a onclick=""DiaWindowOpen_GetImgURL(this)"" title="""&TempImg&""" href=""javascript:;"" target=""_self""><b class=""icon icon_imglink"" title=""图片路径""></b></a></p></li>"
                                     Next
                                  End If
                               %>
                               </ul>
                            </div>                          </td></tr>
                        </tbody>
                        </table>
						<%end if%>
						</td>
    </tr>
	<tr><td align="right">文件名：</td>
	<td><input name="Fk_Product_FileName" type="text" class="Input" id="Fk_Product_FileName" size="20" value="<%=Fk_Product_FileName%>" style="vertical-align:middle"/>&nbsp;*不可修改</td>
		</tr>
	<tr><td align="right">排序：</td>
	<td><input class="Input" type="text" value="<%if isnull(Fk_Product_px) then response.write "0" else response.write Fk_Product_px%>" name="Fk_Product_px" id="Fk_Product_px" size="4" maxlength="6" style="vertical-align:middle"/> （仅限数字，越大越排前）</td>
		</tr>
	<tr>
		<td height="28" align="right">SEO标题：</td>
		<td><input name="Fk_Product_Seotitle" value="<%=Fk_Product_Seotitle%>" type="text" class="Input" id="Fk_Product_Seotitle" size="60"  style="vertical-align:middle;"/></td>
	</tr>
    <tr>
        <td height="28" align="right">SEO关键词：</td>
        <td><input name="Fk_Product_Keyword" value="<%=Fk_Product_Keyword%>" type="text" class="Input" id="Fk_Product_Keyword" size="60" /> <input type="button" onclick="tiqu(0,'Fk_Product_Content','Fk_Product_Keyword');" class="Button" name="btntqkwd" id="btntqkwd" value="提 取" />
        <input  value="<%=KeyWordlist%>" type="hidden" id="Fk_Keywordlist" /> </td>
    </tr>
    <tr>
        <td height="28" align="right">SEO描述：</td>
        <td><input name="Fk_Product_Description" value="<%=Fk_Product_Description%>" type="text" class="Input" id="Fk_Product_Description" size="60" /> 
        <input type="button" onclick="tiqu(1,'Fk_Product_Content','Fk_Product_Description');" class="Button" name="btntqdesc" id="btntqdesc" value="提 取" /></td>
    </tr>
	<tr>
        <td height="28" align="right" valign="top">缩略图：</td>
        <td>
		<div class="image_grid" style="padding:10px;">
                              <table border="0" cellpadding="0" cellspacing="0"><tbody><tr>
                              <td valign="top" style="padding:0px;border-bottom:0px;">
                                  <div class="imageBox" id="ThumbnailBox">
                                  <img class="img" width="80" height="80" id="SlImg" src="<%if Fk_Product_Pic<>"" then response.write Fk_Product_Pic :else : response.write "http://image001.dgcloud01.qebang.cn/website/ext/images/image.jpg"%>">                                  </div>                              </td>
                              <td valign="top" style="padding:0px 0px 0px 6px;border-bottom:0px;">
                                  <input type="text" class="Input" id="Fk_Product_Pic" name="Fk_Product_Pic" style="width:314px; " value="<%=Fk_Product_Pic%>">
                                  <div style="padding:5px 0px 0px 10px;">
                                  <span class="sysBtn" id="UploadImage"><a id="chooseImg" style="cursor:pointer;display: inline-block;width: 100px;" class="Button"><b class="icon icon_add"></b>添加图片...</a></span>                                  </div>
                                  <label class="Normal" id="d_Thumbnail"></label>                              </td>
                              </tr></tbody></table>
                          </div>						  </td>
      </tr>
	  
	<tr>
		<td height="28" align="right">主图视频：</td>
		<td colspan="2" style="padding-left:10px;"><div id="Fk_Product_Video_Div" style="position:relative;margin-top:10px;width:320px;"><video controls="controls" id="Fk_Product_Video_Tag" width="320" src="<%=Fk_Product_Video_File%>"></video><span id="removeVideo" title="移除视频" style="position:absolute;right:10px;top:10px;width: 20px;height: 20px;background: #000;display: flex;justify-content: center;align-items: center;border-radius: 10px;color: #fff;cursor: pointer;">X</span></div><input name="Fk_Product_Video_File" type="hidden" class="Input" id="Fk_Product_Video_File" size="50"  style="vertical-align:middle;"/></td>
	</tr>
	  <%if chkSummary then%>
    <tr>
        <td height="28" align="right">内容摘要：</td>
        <td style="padding:10px;" colspan="2"><textarea name="Fk_Product_Summary"  style="border: 1px solid #D3E3F0;width:300px;padding:5px; font-size:14px;"  id="Fk_Product_Summary" rows="5" cols="50" <%if Fk_Product_Summary="" then response.write "disabled"%>><%=Fk_Product_Summary%></textarea>
		<div style="width:300px;"><input type="checkbox" id="GetAbstract" name="GetAbstract" value="1" <%if Fk_Product_Summary="" then response.write "checked"%> style="vertical-align:middle;"><label for="GetAbstract" style="vertical-align:middle;">采用产品详细的前<%=SiteMini%>个字作为摘要</label></div></td>
	  </tr>
    <%end if%>
        <tr>
        <td height="28" align="right">内容：</td>
        <td colspan="2">
		<div class="dt-detail">
                      <div class="tabbar" id="J_TabBar">
                          <ul><li class="current"><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',1);return false;"><span>产品详细</span></a></li>
	  
                          
                              
                              <li style="<%if not chkContentEx1 then response.write "display:none;"%>"><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',2);return false;"><span><%if chkContentEx1 then response.write getDesc("fk_product","Fk_Product_ContentEx1")%></span></a></li>
                              <li <%if not chkContentEx2 then response.write "style='display:none;'"%>><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',3);return false;"><span><%if chkContentEx2 then response.write getDesc("fk_product","Fk_Product_ContentEx2")%></span></a></li>
                              <li <%if not chkFk_Product_wapimg then response.write "style='display:none;'"%>><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',4);return false;"><span>小程序商品详情</span></a></li>
                              <li><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',5);return false;"><span>资料下载</span></a></li>
                              <li><a href="javascript:void(0);" onclick="TabSwitch('J_TabBar','TabBar_',6);return false;"><span>测试视频</span></a></li>
                          </ul>
                      </div>
                      <!--详细内容-->
                      <div class="product-infoitem" id="TabBar_1">
                          <textarea name="Fk_Product_Content" class="<%=Bianjiqi%>" id="Fk_Product_Content" rows="15" style="width:100%;"><%=Fk_Product_Content%></textarea>                          
                      </div>
                      <!--详细内容:end-->
                      
	  
                      <div class="product-infoitem" id="TabBar_2" style="display:none;">
						  <%if chkContentEx1 then%><textarea name="Fk_Product_ContentEx1" class="<%=Bianjiqi%>" id="Fk_Product_ContentEx1" rows="15" style="width:100%;"><%=Fk_Product_ContentEx1%></textarea><%end if%>
                      </div>
	  
                      <div class="product-infoitem" id="TabBar_3" style="display:none;">
						  <%if chkContentEx2 then%><textarea name="Fk_Product_ContentEx2" class="<%=Bianjiqi%>" id="Fk_Product_ContentEx2" rows="15" style="width:100%;"><%=Fk_Product_ContentEx2%></textarea><%end if%>
                      </div>
	  
                      <div class="product-infoitem" id="TabBar_4" style="display:none;">
						  <%if chkFk_Product_wapimg then%><textarea name="Fk_Product_wapimg" class="<%=Bianjiqi%>" id="Fk_Product_wapimg" rows="15" style="width:100%;"><%=Fk_Product_wapimg%></textarea><%end if%>
                      </div>
	  
                      <div class="product-infoitem" id="TabBar_5" style="display:none;">
						  <div class="related_down_div">
						  <a style="color: #f06100;text-decoration:underline;" href="javascript:void(0);" onclick="OpenBoxNew('Down.asp?Type=10','关联资料下载','800px','500px', '关联选中资料', $('input[name=downId]').map(function(){return $(this).val();}).get(), 'related_down', 'downId');">关联资料下载</a>
							<ul id="related_down">
						  <%if Related_Down_Count > 0 then%>
							<%
							dim arr_i
							for arr_i=0 to Related_Down_Count-1%>
							<li><input type="hidden" value="<%=Arr_Rows(0, arr_i)%>" name="downId"/><span><%=Arr_Rows(1, arr_i)%></span><span class="remove-related"><a href='javascript:void(0);' onclick="$(this).parent().parent('li').remove();">移除</a></span></li>
							<%next%>
						  <%end if%>
							</ul>
						  
						  </div>
                      </div>
	  
                      <div class="product-infoitem" id="TabBar_6" style="display:none;">
						  <div class="related_down_div">
						  <a style="color: #f06100;text-decoration:underline;" href="javascript:void(0);" onclick="OpenBoxNew('Article.asp?Type=10','关联测试视频','800px','500px', '关联选中视频', $('input[name=videoId]').map(function(){return $(this).val();}).get(), 'related_video', 'videoId');">关联测试视频</a>
							<ul id="related_video">
						  <%if Related_Video_Count > 0 then%>
							<%
							dim arr_v
							for arr_v=0 to Related_Video_Count-1%>
							<li><input type="hidden" value="<%=Arr_Rows_Video(0, arr_v)%>" name="downId"/><span><%=Arr_Rows_Video(1, arr_v)%></span><span class="remove-related"><a href='javascript:void(0);' onclick="$(this).parent().parent('li').remove();">移除</a></span></li>
							<%next%>
						  <%end if%>
							</ul>
						  
						  </div>
                      </div>
                  </div>
				  </td>
    </tr>

    <%
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		Temp2=""
		For Each Temp In Fk_Product_Field
			If Split(Temp,"|-Fangka_Field-|")(0)=Rs("Fk_Field_Tag") Then
				Temp2=FKFun.HTMLDncode(Split(Temp,"|-Fangka_Field-|")(1))
				Exit For
			End If
		Next
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="2"><input name="Fk_Product__<%=Rs("Fk_Field_Tag")%>" value="<%=Temp2%>" type="text" class="Input" id="Fk_Product__<%=Rs("Fk_Field_Tag")%>" /></td>
    </tr>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
    <tr>
        <td height="28" align="right">转向链接：</td>
        <td colspan="2"><input name="Fk_Product_Url" type="text" class="Input" id="Fk_Product_Url" value="<%=Fk_Product_Url%>" size="50" style="vertical-align:middle"/>&nbsp;*正常请留空</td>
    </tr>
    <tr>
        <td height="28" align="right">推荐：</td>
        <td colspan="2"><select name="Fk_Product_Recommend" class="Input"  id="Fk_Product_Recommend" style="vertical-align:middle;">
            <option value="0">无推荐</option>
            <option value="2"<%If Instr(Fk_Product_Recommend,",2,")>0 Then%> selected="selected"<%End If%>>推荐</option>
            </select>
			<input name="Fk_Product_onTop" type="checkbox"  id="Fk_Product_onTop" value="1" style="vertical-align:middle;" <%if Fk_Product_onTop="1" then response.write "checked"%>/><label style="vertical-align:middle;" for="Fk_Product_onTop">&nbsp;置顶</label>
			　<select name="Fk_Product_Subject"  id="Fk_Product_Subject" style="vertical-align:middle;display:none;">
            <option value="0">无专题</option>
            </select>是否显示：<input name="Fk_Product_Show" type="radio" class="Input" id="Fk_Product_Show" value="1"<%=FKFun.BeCheck(Fk_Product_Show,1)%> checked="true"  style="vertical-align:middle;"/><label for="Fk_Product_Show" style="vertical-align:middle">显示</label>
        <input type="radio" name="Fk_Product_Show" class="Input" id="Fk_Product_Show1" value="0"<%=FKFun.BeCheck(Fk_Product_Show,0)%>  style="vertical-align:middle;"/><label for="Fk_Product_Show1" style="vertical-align:middle">不显示</label>　模板：<select name="Fk_Product_Template" class="Input" id="Fk_Product_Template" style="vertical-align:middle;">
            <option value="0"<%=FKFun.BeSelect(Fk_Product_Template,0)%>>默认模板</option>
    <%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject')"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Product_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
    <%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select> &nbsp; &nbsp; 点击量：<input style="width:50px" name="Fk_Product_click" type="text" id="Fk_Product_click" class="Input" size="6" maxlength="6" value="<%=Fk_Product_click%>" style="vertical-align:middle;"></td>
    </tr>
</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto;" class="tcbtm">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('ProductEdit','Product.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="btnclose" id="btnclose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ProductEditDo
'作    用：执行修改图文
'参    数：
'==============================
Sub ProductEditDo()
	'on error resume next
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
	'	Response.Write("无权限！")
	'	Call FKDB.DB_Close()
	'	Session.CodePage=936
	'	Response.End()
	'End If
	
	dim GetAbstract,Fk_Product_Summary,FK_Product_SlidesImgs,FK_Product_SlidesFirst,Fk_Product_ContentEx1,Fk_Product_ContentEx2 '2013-11-28 shark 添加
	dim Fk_Product_Video_File
	GetAbstract=Request.Form("GetAbstract")
	FK_Product_SlidesImgs=Trim(Request.Form("SlidesImg[]"))
	FK_Product_SlidesFirst=Trim(Request.Form("SlidesImg[]firstImg"))
	
	Fk_Product_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Title")))
	Fk_Product_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Color")))
	Fk_Product_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Seotitle")))
	Fk_Product_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Keyword")))
	Fk_Product_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Description")))
	Fk_Product_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Pic")))
	Fk_Product_PicBig=Fk_Product_Pic
	Fk_Product_Content=Request.Form("Fk_Product_Content")
	Fk_Product_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Url")))
	Fk_Product_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Recommend"))," ",""))&","
	Fk_Product_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Product_Subject"))," ",""))&","
	Fk_Product_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_FileName")))
	Fk_Product_Show=Trim(Request.Form("Fk_Product_Show"))
	Fk_Product_click=Trim(Request.Form("Fk_Product_click"))
	Fk_Product_Template=Trim(Request.Form("Fk_Product_Template"))
	Fk_Product_Time=Trim(Request.Form("Fk_Product_Time"))
	Fk_Product_onTop=Trim(Request.Form("Fk_Product_onTop"))
	Fk_Product_px=Trim(Request.Form("Fk_Product_px"))
	Fk_Product_Video_File=Trim(Request.Form("Fk_Product_Video_File"))
	If Fk_Product_Time="" Then
		Fk_Product_Time=0
	End If
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Product_Title,1,255,0,"请输入图文标题！","图文标题不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Product_px,"排序只能是数字，请重新填写！")
	Call FKFun.ShowString(Fk_Product_Keyword,0,255,2,"请输入图文SEO关键词！","图文SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Description,0,255,2,"请输入图文SEO描述！","图文SEO描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_Url,0,255,2,"请输入图文转向链接！","图文转向链接不能大于255个字符！")
	
	if GetAbstract<>"1" then
		Fk_Product_Summary=FKFun.HTMLEncode(Trim(Request.Form("Fk_Product_Summary")))
		Call FKFun.ShowString(Fk_Product_Summary,0,500,2,"请输入图文摘要！","图文摘要不能大于500个字符！")
	end if 
	Fk_Product_ContentEx1=Trim(Request.Form("Fk_Product_ContentEx1")&" ")
	Fk_Product_ContentEx2=Trim(Request.Form("Fk_Product_ContentEx2")&" ")
	Fk_Product_wapimg=Trim(Request.Form("Fk_Product_wapimg")&" ")
	
	If Fk_Product_Url="" Then
		Call FKFun.ShowString(Fk_Product_Content,20,1,1,"请输入图文内容，不少于20个字符！","图文内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Product_Pic,0,255,2,"请输入图文缩略图路径！","图文缩略图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Product_PicBig,0,255,2,"请输入图文缩略图路径！","图文缩略图大图路径不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Call FKFun.ShowString(Fk_Product_FileName,0,100,2,"生成文件名不符合要求","生成文件名不能大于100个字符！")
	Call FKFun.ShowNum(Fk_Product_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Product_Show,"请选择是否显示！")
	Call FKFun.ShowNum(Fk_Product_click,"请输入正确的点击量！")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=1 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Fk_Product_Field="" Then
			Fk_Product_Field=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Product__"&Rs("Fk_Field_Tag"))))
		Else
			Fk_Product_Field=Fk_Product_Field&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Product__"&Rs("Fk_Field_Tag"))))
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If IsNumeric(Fk_Product_FileName) Then
		Response.Write("文件名不可用纯数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Left(Fk_Product_FileName,4)="Info" Or Left(Fk_Product_FileName,4)="Page" Or Left(Fk_Product_FileName,5)="GBook" Or Left(Fk_Product_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If SiteDelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Product_Content=Replace(Fk_Product_Content,Temp,"**")
				Fk_Product_Title=Replace(Fk_Product_Title,Temp,"**")
				Fk_Product_Keyword=Replace(Fk_Product_Keyword,Temp,"**")
				Fk_Product_Description=Replace(Fk_Product_Description,Temp,"**")
			End If
		Next
	End If
	Sqlstr="Select * From [Fk_Product] Where Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Product_Title")=Fk_Product_Title
		Rs("Fk_Product_Color")=Fk_Product_Color
		Rs("Fk_Product_Seotitle")=Fk_Product_Seotitle
		Rs("Fk_Product_Keyword")=Fk_Product_Keyword
		Rs("Fk_Product_Field")=trim(Fk_Product_Field&" ")
		Rs("Fk_Product_Url")=Fk_Product_Url
		Rs("Fk_Product_Description")=Fk_Product_Description
		Rs("Fk_Product_Pic")=Fk_Product_Pic
		Rs("Fk_Product_PicBig")=Fk_Product_PicBig
		Rs("Fk_Product_Recommend")=Fk_Product_Recommend
		Rs("Fk_Product_Subject")=Fk_Product_Subject
		Rs("Fk_Product_Show")=Fk_Product_Show
		Rs("Fk_Product_click")=Fk_Product_click
		Rs("Fk_Product_Content")=Fk_Product_Content
		Rs("Fk_Product_FileName")=Fk_Product_FileName
		Rs("Fk_Product_Template")=Fk_Product_Template
		Rs("Fk_Product_Ip")=Fk_Product_onTop
		Rs("Px")=Fk_Product_px
		'添加主图视频字段 修改于2023年1月31日
		Rs("Fk_Product_Video_File")=Fk_Product_Video_File
		
		if chkSummary then
			if GetAbstract<>"1" then
				Rs("Fk_Product_Summary")=Fk_Product_Summary
			else
				Rs("Fk_Product_Summary")=""
			end if 
		end if
		if chkSlidesFirst then
			if FK_Product_SlidesFirst<>"" then
				FK_Product_SlidesFirst=replace(FK_Product_SlidesFirst,"http://"&request.ServerVariables("HTTP_HOST"),"")
				Rs("FK_Product_SlidesFirst")=FK_Product_SlidesFirst
			end if
		end if
		if chkSlidesImgs then
			Rs("FK_Product_SlidesImgs")=FK_Product_SlidesImgs
		end if
		if chkContentEx1 then
			Rs("Fk_Product_ContentEx1")=Fk_Product_ContentEx1
		end if
		if chkContentEx2 then
			Rs("Fk_Product_ContentEx2")=Fk_Product_ContentEx2
		end if
		if chkFk_Product_wapimg then
			Rs("Fk_Product_wapimg")=Fk_Product_wapimg
		end if
		
		If Fk_Product_Time=1 Then
			Rs("Fk_Product_Time")=Now()
		End If
		Rs.Update()
		Application.UnLock()

		dim selectedDownIds, selectedVideoIds
		selectedDownIds = FKFun.HTMLEncode(Trim(Request.Form("downId")))
		selectedVideoIds = FKFun.HTMLEncode(Trim(Request.Form("videoId")))
		conn.execute("update [Fk_Down] set Fk_Relation_Product_Id = null Where Fk_Relation_Product_Id = "& Id)
		conn.execute("update [Fk_Article] set Fk_Relation_Product_Id = null Where Fk_Relation_Product_Id = "& Id)
		' 如果有关联资料下载，则修改Fk_Down表数据 
		if selectedDownIds <> "" then 
			conn.execute("update [Fk_Down] set Fk_Relation_Product_Id = " & Id & " Where Fk_Down_Id in ("& selectedDownIds &")")
		end if
		' 如果有关联测试视频，则修改Fk_Article表数据 
		if selectedVideoIds <> "" then 
			conn.execute("update [Fk_Article] set Fk_Relation_Product_Id = " & Id & " Where Fk_Article_Id in ("& selectedVideoIds &")")
		end if

		Response.Write("“"&Fk_Product_Title&"”修改成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="修改图文：【"&Fk_Product_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("图文不存在！")
	End If
	Rs.Close
	If SiteHtml=1 And Fk_Product_Show=1 And Fk_Product_Url="" Then
		Dim FKHTML
		Set FKHTML=New Cls_HTML
		Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Product_Module=Rs("Fk_Product_Module")
		Rs.Close
		Call FKHTML.CreatProduct(Fk_Product_Template,Fk_Product_Module,Fk_Module_Dir,Fk_Product_FileName,Fk_Product_Title,1)
	Else
		Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Product_Module=Rs("Fk_Product_Module")
		Rs.Close
		If Fk_Module_Dir<>"" Then
			Temp="../"&Fk_Module_Dir&"/"
		Else
			Temp="../Product"&Fk_Product_Module&"/"
		End If
		If Fk_Product_FileName<>"" Then
			Temp=Temp&Fk_Product_FileName&".html"
		Else
			Temp=Temp&Id&".html"
		End If
		Call FKFso.DelFile(Temp)
	End If
End Sub

'==============================
'函 数 名：ProductDelDo
'作    用：执行删除图文
'参    数：
'==============================
Sub ProductDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_ProductList] Where Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		'判断权限
		'If Not FkFun.CheckLimit("Module"&Rs("Fk_Product_Module")) Then
		'	Response.Write("无权限！")
		'	Call FKDB.DB_Close()
		'	Session.CodePage=936
		'	Response.End()
		'End If
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Product_Module=Rs("Fk_Product_Module")
		Fk_Product_FileName=Rs("Fk_Product_FileName")
		Fk_Product_Title=Rs("Fk_Product_Title")
		Rs.Close
		If Fk_Module_Dir<>"" Then
			Temp="../"&Fk_Module_Dir&"/"
		Else
			Temp="../Product"&Fk_Product_Module&"/"
		End If
		If Fk_Product_FileName<>"" Then
			Temp=Temp&Fk_Product_FileName&".html"
		Else
			Temp=Temp&Id&".html"
		End If
		Call FKFso.DelFile(Temp)
	Else
		Rs.Close
		Response.Write("图文不存在！")
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Product] Where Fk_Product_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("图文删除成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="删除图文：【"&Fk_Product_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("图文不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ListDelDo
'作    用：执行批量删除图文
'参    数：
'==============================
Sub ListDelDo()
	Id=Replace(Trim(Request("ListId"))," ","")
	If Id="" Then
		Response.Write("请选择要删除的图文！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="select Fk_Product_Title From [Fk_Product] Where Fk_Product_Id In ("&Id&")"
	Rs.Open Sqlstr,Conn,1,3
	if not rs.eof then
		i=0
		do while not rs.eof
		if i=0 then
			Fk_Product_Title="【"&rs("Fk_Product_Title")&"】"
		else
			Fk_Product_Title=Fk_Product_Title&","&"【"&rs("Fk_Product_Title")&"】"
		end if
		i=i+1
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		rs.movenext
		loop
	end if
	Response.Write("图文批量删除成功！")
	'插入日志
	on error resume next
	dim log_content,log_ip,log_user
	log_content="批量删除图文："&Fk_Product_Title&""
	log_user=Request.Cookies("FkAdminName")
	
	log_ip=FKFun.getIP()
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
End Sub

'==============================
'函 数 名：ProductMove
'作    用：执行批量移动图文
'参    数：
'==============================
Sub ProductMove()
	Dim Fk_Module_Type
	Id=Replace(Trim(Request.QueryString("ListId"))," ","")
	Fk_Module_Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Fk_Module_Id,"请选择转移到的栏目！")
	If Id="" Then
		Response.Write("请选择要移动的图文！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Fk_Module_Type=1000
	Sqlstr="Select Fk_Module_Type From [Fk_Module] Where Fk_Module_Id="&Fk_Module_Id&""
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Type=Rs("Fk_Module_Type")
	End If
	Rs.Close
	If Fk_Module_Type=1000 Then
		Response.Write("要移到的栏目不存在！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Fk_Module_Type<>2 Then
		Response.Write("只能移动到图文栏目！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Update [Fk_Product] Set Fk_Product_Module="&Fk_Module_Id&" Where Fk_Product_Id In ("&Id&")"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("图文批量移动成功！")
End Sub

'==============================
'函 数 名：ProductOrderDo
'作    用：执行图文内容批量排序
'参    数：
'==============================
Sub ProductOrderDo()
	Id=Replace(Trim(Request.Form("id[]"))," ","")
	Fk_Product_px=Replace(Trim(Request.Form("Fk_Product_px[]"))," ","")
	dim idarr
	If Id="" Then
		Response.Write("请选择要排序的内容！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Fk_Product_px="" Then
		Response.Write("请输入排序数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	dim arrId,arrPx,comma
	comma=""
	if instr(Id,",")>0 and instr(Fk_Product_px,",")>0 then
		arrId = split(Id,",")
		arrPx = split(Fk_Product_px,",")
		if ubound(arrId)=ubound(arrPx) then
			for i = 0 to ubound(arrId)
				Call FKFun.ShowNum(arrPx(i),"排序序号必须是有效数字！")
				Sqlstr="Select Px From [Fk_Product] Where Fk_Product_Id="&arrId(i)&""
				Rs.Open Sqlstr,Conn,1,3
				if not Rs.eof then
					Application.Lock()
						Fk_Product_Title=Fk_Product_Title & comma&"【"&arrId(i)&"】"
						comma = ","
						Rs("Px")=arrPx(i)
						Rs.Update()
					Application.UnLock()
					Rs.Close
				else
					Rs.Close
				end if
			next
		else
			Response.Write("参数错误，请刷新页面！")
			Call FKDB.DB_Close()
			Response.End()
		end if
	elseif instr(Id,",")=0 and instr(Fk_Product_px,",")=0 and Id>0 then
		Call FKFun.ShowNum(Fk_Product_px,"排序序号必须是有效数字！")
		Sqlstr="Select Px From [Fk_Product] Where Fk_Product_Id="&Id&""
		Rs.Open Sqlstr,Conn,1,3
		if not Rs.eof then
			Application.Lock()
				Fk_Product_Title="【"&Id&"】"
				Rs("Px")=Fk_Product_px
				Rs.Update()
			Application.UnLock()
			Rs.Close
		else
			Rs.Close
		end if
	else
		Response.Write("参数错误，请刷新页面！")
		Call FKDB.DB_Close()
		Response.End()
	end if
	
	Response.Write("内容排序批量更新成功！")
	'插入日志
	on error resume next
	dim log_content,log_ip,log_user
	log_content="批量更新图文栏目下ID为"&Fk_Product_Title&"的内容排序"
	log_user=Request.Cookies("FkAdminName")
		
	log_ip=FKFun.getIP()
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
End Sub
%><!--#Include File="../Code.asp"-->