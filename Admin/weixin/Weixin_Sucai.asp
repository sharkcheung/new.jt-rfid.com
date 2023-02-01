<!--#Include File="../AdminCheck.asp"--><!--#Include File="../../inc/Md5.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<%
'==========================================
'文 件 名：weixin_Sucai.asp
'文件用途：微信素材管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System2") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'定义页面变量
Dim Fk_Sucai_Title,Fk_Sucai_source,Fk_Sucai_Summary,Fk_Sucai_status,Fk_Sucai_px,Fk_Sucai_url,Fk_Sucai_Content,Fk_Sucai_Id_List

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WeixinSucaiList() '微信素材列表
	Case 2
		Call WeixinSucaiPicAdd() '添加微信图片素材
	Case 3
		Call WeixinSucaiPicAddDo() '执行添加微信图片素材
	Case 4
		Call WeixinSucaiPicEditForm() '修改微信图片素材
	Case 5
		Call WeixinSucaiPicEditDo() '执行修改微信图片素材
	Case 6
		Call WeixinSucaiDelDo() '执行删除微信素材
	Case 7
		Call WeixinSucaiPx() '执行批量排序
	Case 8
		Call WeixinSucaiOpen() '执行批量启用
	Case 9
		Call WeixinSucaiClose() '执行批量禁用
	Case 10
		Call WeixinSucaiYulan()  '微信素材预览
	Case 11
		Call WeixinSucaiMediaAdd() '添加微信语音素材
	Case 12
		Call WeixinSucaiMediaAddDo() '执行添加微信语音素材
	Case 13
		Call WeixinSucaiMediaEditForm() '修改微信语音素材
	Case 14
		Call WeixinSucaiMediaEditDo() '执行修改微信语音素材
	Case Else
		Response.Write("没有找到此功能项！")
End Select

sub WeixinSucaiOpen()	
	Id=Trim(Request("Id"))
	if id<>"" then
		if instr(id,",")>0 then
			dim arr,arrpx
			arr=split(id,",")
			for i=0 to ubound(arr)			
				conn.execute("update [weixin_Sucai] set Sucai_status=0 where id="&arr(i))
			next
		else
			conn.execute("update [weixin_Sucai] set Sucai_status=0 where id="&Id)
		end if	
		Response.Write("批量启用成功！")
	end if
end sub


sub WeixinSucaiClose()	
	Id=Trim(Request("Id"))
	if id<>"" then
		if instr(id,",")>0 then
			dim arr,arrpx
			arr=split(id,",")
			for i=0 to ubound(arr)			
				conn.execute("update [weixin_Sucai] set Sucai_status=1 where id="&arr(i))
			next
		else
			conn.execute("update [weixin_Sucai] set Sucai_status=1 where id="&Id)
		end if	
		Response.Write("批量禁用成功！")
	end if
end sub

sub WeixinSucaiPx()	
	dim px
	Id=Trim(Request("Id"))
	px=Trim(Request("px"))
	if id<>"" then
		if instr(id,",")>0 then
			dim arr,arrpx
			arr=split(id,",")
			arrpx=split(px,",")
			for i=0 to ubound(arr)			
				conn.execute("update [weixin_Sucai] set Sucai_px="&arrpx(i)&" where id="&arr(i))
			next
		else
			conn.execute("update [weixin_Sucai] set Sucai_px="&px&" where id="&Id)
		end if	
		Response.Write("批量排序成功！")
	end if
end sub

'==========================================
'函 数 名：WeixinSucaiList()
'作    用：微信素材列表
'参    数：
'==========================================
Sub WeixinSucaiList()
Session("NowPage")=FkFun.GetNowUrl()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false;">素材管理</a></li>
        <li><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/weixin_Sucai.asp?Type=11');return false;">添加语音</a></li>
        <li><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/weixin_Sucai.asp?Type=2');return false;">添加图片</a></li>
    </ul>
</div>
<div id="ListContent">
    <form name="DelList" id="DelList" method="post" action="Down.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop">名称</td>
            <td align="center" class="ListTdTop">类型</td>
            <td align="center" class="ListTdTop">来源</td>
            <td align="center" class="ListTdTop">素材</td>
            <td align="center" class="ListTdTop">大小</td>
            <td align="center" class="ListTdTop">排序</td>
            <td align="center" class="ListTdTop">时间</td>
            <td align="center" class="ListTdTop">状态</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Dim Rs,yurl,t
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_Sucai] Order By Sucai_Px Desc,id desc"
	Rs.Open Sqlstr,Conn,1,1
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
            <td height="20" align="center"><input type="checkbox" name="Id" class="Checks" value="<%=Rs("id")%>" id="List<%=Rs("id")%>" /></td>
            <td ><%=Rs("Sucai_Title")%></td>
            <td align="center"><%if Rs("Sucai_type")=0 then 
			response.write "图片"
			t=4
			else
			response.write "语音"
			t=13
			end if%></td>
			<td align="center"><%if Rs("Sucai_source")=0 then 
				response.write "外链"
			else
				response.write "上传"
			end if%></td>
            <td align="center"><%if Rs("Sucai_type")=0 then%>			
			<img class="preview" width="45" bimg="<%=rs("Sucai_file")%>" src="<%=rs("Sucai_file")%>" title="<%=Rs("Sucai_Title")%>">			
			<%else%>
			<embed flashvars="mp3=<%=Rs("Sucai_file")%>&autoplay=0" height="20" src="http://image001.dgcloud01.qebang.cn/website/weixin/music_player.swf" type="application/x-shockwave-flash" width="160" wmode="transparent">	
			<%end if%></td>
			<td align="right"><%=Rs("Sucai_filesize")%></td>
            <td align="center"><input type="text" value="<%=Rs("Sucai_px")%>" class="Input" name="px" size=2 style="text-align:center"/></td>
            <td height="20" align="center"><%=Rs("Sucai_time")%></td>
            <td align="center"><%if Rs("Sucai_status")=0 then:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_1.gif' title='启用'>":else:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif' title='禁用'>":end if%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/weixin_Sucai.asp?Type=<%=t%>&Id=<%=Rs("id")%>');return false;"><img src="/admin/images/edit.png" title="编辑"></a> </td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>        <tr>
            <td height="30" colspan="10">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style='text-indent:10px;vertical-align:middle'> 全选
            <input type="submit" value="排序" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_Sucai.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'>
			<input type="submit" value="启用" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_Sucai.asp?Type=8',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'>
			<input type="submit" value="禁用" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_Sucai.asp?Type=9',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'>
			<input type="submit" value="删除" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_Sucai.asp?Type=6',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'></td>
        </tr>
		 <tr>
            <td height="30" colspan="10"><%Call FKFun.ShowPageCode("Down.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="10" align="center">暂无素材</td>
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
'函 数 名：WeixinSucaiPicAdd()
'作    用：添加微信图片素材
'参    数：
'==========================================
Sub WeixinSucaiPicAdd()
%>
<link href="/admin/dkidtioenr/themes/default/default.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	.uploadclass{display:inline;}
</style>
<script type="text/javascript">	

	function t(){
		if(window.KindEditor){
			if(("#Fk_Sucai_url").length>0){
				var h,t,ty,typ;
				
					if($("input.Fk_Sucai_source:checked").val()==0){ //外链
						h="<span> 推荐格式：png, jpg, gif, bmp</span>";
					}
					else{
						h="<span> 文件小于500k；推荐格式：png, jpg, gif, bmp</span>";
					}
					t="image";
					ty=0;
				if($(".Fk_Sucai_source:checked").val()==0){
					$(".uploadclass").html(h);
				}
				else{
					$(".uploadclass").html(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> "+h+"");
				}
						$('#uploadButton').unbind("click");
						
						$('#uploadButton').bind("click",function() {
								var editor = window.KindEditor.editor({
									fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
									uploadJson		: '/admin/weixin/upload_json.asp',
									allowFileManager : true
								});
								editor.loadPlugin('image', function() {
									editor.plugin.imageDialog({
										imageUrl: $('#Fk_Sucai_url').val(),
										clickFn : function(url) {
											$('#Fk_Sucai_url').val(url);
											editor.hideDialog();
										}
									});
								});
							}
						);
		
					
			}
		}
	}
    $(document).ready(function() {
		
		
		// 素材来源切换
		$('#Fk_Sucai_source, #Fk_Sucai_source1').click(function() {
			t();
			
		})
		
		

    });
	
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/weixin_Sucai.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>添加图片</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_Sucai_Title" type="text" class="Input" id="Fk_Sucai_Title" size="60" /></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">来源：</td>
	        <td><input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source" style="vertical-align:middle;" value="0"  checked="checked"/> <label for="Fk_Sucai_source" style="vertical-align:middle;">外链</label> &nbsp; <input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source1"  style="vertical-align:middle;" value="1"/> <label for="Fk_Sucai_source1" style="vertical-align:middle;">上传</label></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">素材：</td>
	        <td><input name="Fk_Sucai_url" id="Fk_Sucai_url" class="Input" size=60/> <div class="uploadclass"><span>推荐格式：png, jpg, gif, bmp</span></div></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">描述：</td>
	        <td><textarea name="Fk_Sucai_desc" id="Fk_Sucai_desc" class="Input" style="background:#E8F6FE;height:70px;width:320px;"></textarea></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">排序：</td>
	        <td><input name="Fk_Sucai_px" class="Input" type="text" id="Fk_Sucai_px" value="0"></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">状态：</td>
	        <td><input name="Fk_Sucai_status" class="Input" type="radio" id="Fk_Sucai_status" value="0" checked="checked"  style="vertical-align:middle;"/> <label for="Fk_Sucai_status" style="vertical-align:middle;">启用</label>
            &nbsp; <input type="radio" name="Fk_Sucai_status" class="Input" id="Fk_Sucai_status1" value="1"  style="vertical-align:middle;"/> <label for="Fk_Sucai_status1" style="vertical-align:middle;">禁用</label></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/weixin_Sucai.asp?Type=3',0,'',0,1,'MainRight','/admin/weixin/weixin_Sucai.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinSucaiPicAddDo
'作    用：执行添加微信图片素材
'参    数：
'==============================
Sub WeixinSucaiPicAddDo()
	Fk_Sucai_Title	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Title")))
	Fk_Sucai_source = Trim(Request.Form("Fk_Sucai_source"))
	Fk_Sucai_Summary= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Summary")))
	Fk_Sucai_url	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_url")))
	Fk_Sucai_px		= Trim(Request.Form("Fk_Sucai_px"))
	Fk_Sucai_status	= Trim(Request.Form("Fk_Sucai_status"))
	Call FKFun.ShowString(Fk_Sucai_Title,1,100,0,"请输入标题名称！","标题名称不能大于100个字节！")
	Call FKFun.ShowString(Fk_Sucai_Title,1,255,0,"素材不能为空！","素材不能大于255个字节！")
	Sqlstr="Select * From [weixin_Sucai]"
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs.AddNew()
		Rs("Sucai_Title")	=Fk_Sucai_Title
		Rs("Sucai_desc")	=Fk_Sucai_Summary
		Rs("Sucai_file")	=Fk_Sucai_url
		Rs("Sucai_source")	=Fk_Sucai_source
		Rs("Sucai_px")		=Fk_Sucai_px
		Rs("Sucai_status")	=Fk_Sucai_status
		Rs("Sucai_type")	=0
		Rs.Update()
		Application.UnLock()
		Response.Write("素材添加成功！")
	Rs.Close
End Sub

'==========================================
'函 数 名：WeixinSucaiPicEditForm
'作    用：修改微信图片素材
'参    数：
'==========================================
Sub WeixinSucaiPicEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [weixin_Sucai] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Sucai_Title	= Rs("Sucai_Title")
		Fk_Sucai_Summary= Rs("Sucai_desc")
		Fk_Sucai_url	= Rs("Sucai_file")
		Fk_Sucai_px		= Rs("Sucai_px")
		Fk_Sucai_status	= Rs("Sucai_status")
		Fk_Sucai_source	= Rs("Sucai_source")
	End If
	Rs.Close
%>
<link href="/admin/dkidtioenr/themes/default/default.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	.uploadclass{display:inline;}
</style>
<script type="text/javascript">	

	function t(){
		if(window.KindEditor){
			if(("#Fk_Sucai_url").length>0){
				var h,t,ty,typ;
				
					if($("input.Fk_Sucai_source:checked").val()==0){ //外链
						h="<span> 推荐格式：png, jpg, gif, bmp</span>";
					}
					else{
						h="<span> 文件小于500k；推荐格式：png, jpg, gif, bmp</span>";
					}
					t="image";
					ty=0;
				if($(".Fk_Sucai_source:checked").val()==0){
					$(".uploadclass").html(h);
				}
				else{
					$(".uploadclass").html(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> "+h+"");
				}
						$('#uploadButton').unbind("click");
						
						$('#uploadButton').bind("click",function() {
								var editor = window.KindEditor.editor({
									fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
									uploadJson		: '/admin/weixin/upload_json.asp',
									allowFileManager : true
								});
								editor.loadPlugin('image', function() {
									editor.plugin.imageDialog({
										imageUrl: $('#Fk_Sucai_url').val(),
										clickFn : function(url) {
											$('#Fk_Sucai_url').val(url);
											editor.hideDialog();
										}
									});
								});
							}
						);
		
					
			}
		}
	}
    $(document).ready(function() {
		
		
		// 素材来源切换
		$('#Fk_Sucai_source, #Fk_Sucai_source1').click(function() {
			t();
			
		})
		
		

    });
	
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/weixin_Sucai.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>修改图片素材</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_Sucai_Title" type="text" class="Input" id="Fk_Sucai_Title" size="40" value="<%=Fk_Sucai_Title%>"/><input type="hidden" value="<%=id%>" name="id"/></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">来源：</td>
	        <td><input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source" style="vertical-align:middle;" value="0" <%if Fk_Sucai_source=0 then response.write "checked=""checked"""%> /> <label for="Fk_Sucai_source" style="vertical-align:middle;">外链</label> &nbsp; <input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source1"  style="vertical-align:middle;" value="1" <%if Fk_Sucai_source=1 then response.write "checked=""checked"""%>/> <label for="Fk_Sucai_source1" style="vertical-align:middle;">上传</label></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">素材：</td>
	        <td><input name="Fk_Sucai_url" id="Fk_Sucai_url" class="Input" size="60"  value="<%=Fk_Sucai_url%>"/> <div class="uploadclass"><%if Fk_Sucai_source=1 then 
			response.write "<input type=""button"" id=""uploadButton"" value=""上传"" class=""Button"" style=""vertical-align:middle;""/> <span> 文件小于500k；推荐格式：png, jpg, gif, bmp</span>"
			else
			response.write "<span> 推荐格式：png, jpg, gif, bmp</span>"
			end if
			%></div></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">摘要：</td>
	        <td><textarea name="Fk_Sucai_Summary" id="Fk_Sucai_Summary" class="Input" style="background:#E8F6FE;height:70px;width:320px;"><%=Fk_Sucai_Summary%></textarea></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">排序：</td>
	        <td><input name="Fk_Sucai_px" class="Input" type="text" id="Fk_Sucai_px"  value="<%=Fk_Sucai_px%>"></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">状态：</td>
	        <td><input name="Fk_Sucai_status" class="Input" type="radio" id="Fk_Sucai_status" value="0" checked="checked" <%if Fk_Sucai_status=0 then response.write "checked"%>  style="vertical-align:middle;"/> <label for="Fk_Sucai_status" style="vertical-align:middle;">启用</label>
            <input type="radio" name="Fk_Sucai_status" class="Input" id="Fk_Sucai_status1" value="1" <%if Fk_Sucai_status=1 then response.write "checked"%>  style="vertical-align:middle;"/> <label for="Fk_Sucai_status1" style="vertical-align:middle;">禁用</label></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/weixin_Sucai.asp?Type=5',0,'',0,1,'MainRight','/admin/weixin/weixin_Sucai.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinSucaiPicEditDo
'作    用：执行修改图片素材
'参    数：
'==============================
Sub WeixinSucaiPicEditDo()
	Id				= Trim(Request.Form("Id"))
	Fk_Sucai_Title	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Title")))
	Fk_Sucai_source = Trim(Request.Form("Fk_Sucai_source"))
	Fk_Sucai_Summary= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Summary")))
	Fk_Sucai_url	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_url")))
	Fk_Sucai_px		= Trim(Request.Form("Fk_Sucai_px"))
	Fk_Sucai_status	= Trim(Request.Form("Fk_Sucai_status"))
	Call FKFun.ShowString(Fk_Sucai_Title,1,100,0,"请输入标题名称！","标题名称不能大于100个字节！")
	Call FKFun.ShowString(Fk_Sucai_Title,1,255,0,"素材不能为空！","素材不能大于255个字节！")
	Sqlstr="Select * From [weixin_Sucai] where id="&id
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs("Sucai_Title")	=Fk_Sucai_Title
		Rs("Sucai_desc")	=Fk_Sucai_Summary
		Rs("Sucai_file")	=Fk_Sucai_url
		Rs("Sucai_source")	=Fk_Sucai_source
		Rs("Sucai_px")		=Fk_Sucai_px
		Rs("Sucai_status")	=Fk_Sucai_status
		Rs.Update()
		Application.UnLock()
		Response.Write("图片素材修改成功！")
	Rs.Close
End Sub


'==========================================
'函 数 名：WeixinSucaiMediaAdd()
'作    用：添加微信语音素材
'参    数：
'==========================================
Sub WeixinSucaiMediaAdd()
%>
<link href="/admin/dkidtioenr/themes/default/default.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	.uploadclass{display:inline;}
</style>
<script type="text/javascript">	

	function t(){
		if(window.KindEditor){
			if(("#Fk_Sucai_url").length>0){
				var h,t,ty,typ;
				
					if($("input.Fk_Sucai_source:checked").val()==0){ //外链
						h="<span> 推荐格式：mp3, wma, wav, amr</span>";
					}
					else{
						h="<span> 文件小于2M；推荐格式：mp3, wma, wav, amr</span>";
					}
					t="image";
					ty=0;
				if($(".Fk_Sucai_source:checked").val()==0){
					$(".uploadclass").html(h);
				}
				else{
					$(".uploadclass").html(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> "+h+"");
				}
						$('#uploadButton').unbind("click");
						
						$('#uploadButton').bind("click",function() {
								var editor = window.KindEditor.editor({
									fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
									uploadJson		: '/admin/weixin/upload_json.asp',
									allowFileManager : true
								});
								editor.loadPlugin('insertfile', function() {
									editor.plugin.fileDialog({
										fileUrl: $('#Fk_Sucai_url').val(),
										clickFn : function(url) {
											$('#Fk_Sucai_url').val(url);
											editor.hideDialog();
										}
									});
								});
							}
						);
		
					
			}
		}
	}
    $(document).ready(function() {
		
		
		// 素材来源切换
		$('#Fk_Sucai_source, #Fk_Sucai_source1').click(function() {
			t();
			
		})
		
		
		$(".MusicCj").bind("click",function(){
			
			ymPrompt.win({message:'/admin/weixin/weixin_MusicCj.asp?type=1&id=0',
				width : 800,
				height :450,
				title:'采集语音',
				btn: [['确定','ok'],['关闭','close']],
				maxBtn : true,
				minBtn : true,
				closeBtn : true,
				iframe : true,handler:function(msg){
					if (msg == 'error') {
					
					}else if(msg == 'ok'){ 
						if($("iframe").contents().find("input.Checks:checked").length>0){	
							var id;
							id=$("iframe").contents().find("input.Checks:checked").val();
							$('#wx_Subscribe').val("[wx_news="+id+"]");
						}
					}
				}
			});return false;
		
		})

    });
	
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/weixin_Sucai.asp?Type=12" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>添加语音</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_Sucai_Title" type="text" class="Input" id="Fk_Sucai_Title" size="60" /></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">来源：</td>
	        <td><input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source" style="vertical-align:middle;" value="0"  checked="checked"/> <label for="Fk_Sucai_source" style="vertical-align:middle;">外链</label> &nbsp; <input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source1"  style="vertical-align:middle;" value="1"/> <label for="Fk_Sucai_source1" style="vertical-align:middle;">上传</label></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">素材：</td>
	        <td><input name="Fk_Sucai_url" id="Fk_Sucai_url" class="Input" size=60/> <div class="uploadclass"><span>推荐格式：mp3, wma, wav, amr</span></div></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">描述：</td>
	        <td><textarea name="Fk_Sucai_desc" id="Fk_Sucai_desc" class="Input" style="background:#E8F6FE;height:70px;width:320px;"></textarea></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">排序：</td>
	        <td><input name="Fk_Sucai_px" class="Input" type="text" id="Fk_Sucai_px" value="0"></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">状态：</td>
	        <td><input name="Fk_Sucai_status" class="Input" type="radio" id="Fk_Sucai_status" value="0" checked="checked"  style="vertical-align:middle;"/> <label for="Fk_Sucai_status" style="vertical-align:middle;">启用</label>
            &nbsp; <input type="radio" name="Fk_Sucai_status" class="Input" id="Fk_Sucai_status1" value="1"  style="vertical-align:middle;"/> <label for="Fk_Sucai_status1" style="vertical-align:middle;">禁用</label></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/weixin_Sucai.asp?Type=12',0,'',0,1,'MainRight','/admin/weixin/weixin_Sucai.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinSucaiMediaAddDo
'作    用：执行添加微信语音素材
'参    数：
'==============================
Sub WeixinSucaiMediaAddDo()
	Fk_Sucai_Title	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Title")))
	Fk_Sucai_source = Trim(Request.Form("Fk_Sucai_source"))
	Fk_Sucai_Summary= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Summary")))
	Fk_Sucai_url	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_url")))
	Fk_Sucai_px		= Trim(Request.Form("Fk_Sucai_px"))
	Fk_Sucai_status	= Trim(Request.Form("Fk_Sucai_status"))
	Call FKFun.ShowString(Fk_Sucai_Title,1,100,0,"请输入标题名称！","标题名称不能大于100个字节！")
	Call FKFun.ShowString(Fk_Sucai_Title,1,255,0,"素材不能为空！","素材不能大于255个字节！")
	Sqlstr="Select * From [weixin_Sucai]"
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs.AddNew()
		Rs("Sucai_Title")	=Fk_Sucai_Title
		Rs("Sucai_desc")	=Fk_Sucai_Summary
		Rs("Sucai_file")	=Fk_Sucai_url
		Rs("Sucai_source")	=Fk_Sucai_source
		Rs("Sucai_px")		=Fk_Sucai_px
		Rs("Sucai_status")	=Fk_Sucai_status
		Rs("Sucai_type")	=1
		Rs.Update()
		Application.UnLock()
		Response.Write("语音素材添加成功！")
	Rs.Close
End Sub

'==========================================
'函 数 名：WeixinSucaiMediaEditForm
'作    用：修改微信语音素材
'参    数：
'==========================================
Sub WeixinSucaiMediaEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [weixin_Sucai] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Sucai_Title	= Rs("Sucai_Title")
		Fk_Sucai_Summary= Rs("Sucai_desc")
		Fk_Sucai_url	= Rs("Sucai_file")
		Fk_Sucai_px		= Rs("Sucai_px")
		Fk_Sucai_status	= Rs("Sucai_status")
		Fk_Sucai_source	= Rs("Sucai_source")
	End If
	Rs.Close
%>
<link href="/admin/dkidtioenr/themes/default/default.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	.uploadclass{display:inline;}
</style>
<script type="text/javascript">	

	function t(){
		if(window.KindEditor){
			if(("#Fk_Sucai_url").length>0){
				var h,t,ty,typ;
				
					if($("input.Fk_Sucai_source:checked").val()==0){ //外链
						h="<span> 推荐格式：mp3, wma, wav, amr</span>";
					}
					else{
						h="<span> 文件小于2M；推荐格式：mp3, wma, wav, amr</span>";
					}
					t="image";
					ty=0;
				if($(".Fk_Sucai_source:checked").val()==0){
					$(".uploadclass").html(h);
				}
				else{
					$(".uploadclass").html(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> "+h+"");
				}
						$('#uploadButton').unbind("click");
						
						$('#uploadButton').bind("click",function() {
								var editor = window.KindEditor.editor({
									fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
									uploadJson		: '/admin/weixin/upload_json.asp',
									allowFileManager : true
								});
								editor.loadPlugin('insertfile', function() {
									editor.plugin.fileDialog({
										fileUrl: $('#Fk_Sucai_url').val(),
										clickFn : function(url) {
											$('#Fk_Sucai_url').val(url);
											editor.hideDialog();
										}
									});
								});
							}
						);
		
					
			}
		}
	}
    $(document).ready(function() {
		
		if ($('#Fk_Sucai_source1:checked').length>0){
			t();
		}
		// 素材来源切换
		$('#Fk_Sucai_source, #Fk_Sucai_source1').click(function() {
			t();
			
		})
		
		

    });
	
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/weixin_Sucai.asp?Type=14" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>修改语音素材</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_Sucai_Title" type="text" class="Input" id="Fk_Sucai_Title" size="40" value="<%=Fk_Sucai_Title%>"/><input type="hidden" value="<%=id%>" name="id"/></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">来源：</td>
	        <td><input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source" style="vertical-align:middle;" value="0" <%if Fk_Sucai_source=0 then response.write "checked=""checked"""%> /> <label for="Fk_Sucai_source" style="vertical-align:middle;">外链</label> &nbsp; <input name="Fk_Sucai_source" type="radio" class="Input Fk_Sucai_source" id="Fk_Sucai_source1"  style="vertical-align:middle;" value="1" <%if Fk_Sucai_source=1 then response.write "checked=""checked"""%>/> <label for="Fk_Sucai_source1" style="vertical-align:middle;">上传</label></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">素材：</td>
	        <td><input name="Fk_Sucai_url" id="Fk_Sucai_url" class="Input" size="60"  value="<%=Fk_Sucai_url%>"/> <div class="uploadclass"><%if Fk_Sucai_source=1 then 
			response.write "<input type=""button"" id=""uploadButton"" value=""上传"" class=""Button"" style=""vertical-align:middle;""/> <span> 文件小于500k；推荐格式：png, jpg, gif, bmp</span>"
			else
			response.write "<span> 推荐格式：png, jpg, gif, bmp</span>"
			end if
			%></div></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">摘要：</td>
	        <td><textarea name="Fk_Sucai_Summary" id="Fk_Sucai_Summary" class="Input" style="background:#E8F6FE;height:70px;width:320px;"><%=Fk_Sucai_Summary%></textarea></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">排序：</td>
	        <td><input name="Fk_Sucai_px" class="Input" type="text" id="Fk_Sucai_px"  value="<%=Fk_Sucai_px%>"></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">状态：</td>
	        <td><input name="Fk_Sucai_status" class="Input" type="radio" id="Fk_Sucai_status" value="0" checked="checked" <%if Fk_Sucai_status=0 then response.write "checked"%>  style="vertical-align:middle;"/> <label for="Fk_Sucai_status" style="vertical-align:middle;">启用</label>
            <input type="radio" name="Fk_Sucai_status" class="Input" id="Fk_Sucai_status1" value="1" <%if Fk_Sucai_status=1 then response.write "checked"%>  style="vertical-align:middle;"/> <label for="Fk_Sucai_status1" style="vertical-align:middle;">禁用</label></td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/weixin_Sucai.asp?Type=14',0,'',0,1,'MainRight','/admin/weixin/weixin_Sucai.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinSucaiMediaEditDo
'作    用：执行修改语音素材
'参    数：
'==============================
Sub WeixinSucaiMediaEditDo()
	Id				= Trim(Request.Form("Id"))
	Fk_Sucai_Title	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Title")))
	Fk_Sucai_source = Trim(Request.Form("Fk_Sucai_source"))
	Fk_Sucai_Summary= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_Summary")))
	Fk_Sucai_url	= FKFun.HTMLEncode(Trim(Request.Form("Fk_Sucai_url")))
	Fk_Sucai_px		= Trim(Request.Form("Fk_Sucai_px"))
	Fk_Sucai_status	= Trim(Request.Form("Fk_Sucai_status"))
	Call FKFun.ShowString(Fk_Sucai_Title,1,100,0,"请输入标题名称！","标题名称不能大于100个字节！")
	Call FKFun.ShowString(Fk_Sucai_Title,1,255,0,"素材不能为空！","素材不能大于255个字节！")
	Sqlstr="Select * From [weixin_Sucai] where id="&id
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs("Sucai_Title")	=Fk_Sucai_Title
		Rs("Sucai_desc")	=Fk_Sucai_Summary
		Rs("Sucai_file")	=Fk_Sucai_url
		Rs("Sucai_source")	=Fk_Sucai_source
		Rs("Sucai_px")		=Fk_Sucai_px
		Rs("Sucai_status")	=Fk_Sucai_status
		Rs.Update()
		Application.UnLock()
		Response.Write("语音素材修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：WeixinSucaiDelDo
'作    用：执行删除微信素材
'参    数：
'==============================
Sub WeixinSucaiDelDo()
	Id=Trim(Request("Id"))
	Sqlstr="Select * From [weixin_Sucai] Where id in("& Id &")"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("微信素材删除成功！")
	Else
		Response.Write("微信素材不存在！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：WeixinSucaiYulan
'作    用：微信素材预览
'参    数：
'==========================================
Sub WeixinSucaiYulan()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [weixin_Sucai] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Sucai_Title	= Rs("Sucai_Title")
		Fk_Sucai_Summary	= Rs("Sucai_Summary")
		Fk_Sucai_url		= Rs("Sucai_url")
		Fk_Sucai_Pic		= Rs("Sucai_Pic")
		Fk_Sucai_px		= Rs("Sucai_px")
		Fk_Sucai_status	= Rs("Sucai_status")
		Fk_Sucai_Id_List	= Rs("Sucai_Id_List")
		Fk_Sucai_Content	= Rs("Sucai_Content")
	End If
	Rs.Close
%>
<style type="text/css">
/*
178
*/
/**
 * General
 */

.left { float:left}
.right { float:right}
.clear { clear:both}
.hide {display:none;}
.container {}
.clr { clear:both; height:1px; overflow:hidden;display:block; }
.clrLeft { clear:left; height:1px; overflow:hidden; }
.clrRight { clear:right; height:1px; overflow:hidden; }

.btn_green {
	border: 1px solid #8CAD4F;
	border-right: 1px solid #6D883A;
	border-bottom: 3px solid #6D883A;
	background-color: #B2D56C;
	color: #5E7634;
	font-weight: bold;
	text-shadow: 1px 0 0 rgba(255, 255, 255, 0.4);
	cursor: pointer;
}

/**
 * Avatar_face
 */
.icon_face { display: block; width: 26px; height: 26px; background-image: url("../images/icon_face.png"); }
.face {background-position: -0px -0px;}

.icon_face_large { display: block; width: 42px; height: 42px; background-image: url("../images/icon_face_large.png"); }
.face_large {background-position: -0px -0px;}

.icon_face_robot {	
	width: 30px;
	height: 30px;
	overflow: hidden;
}
.face_robot_def {
	background: url('../images/v5_chat/v5_call_icons.png') no-repeat scroll 0 -46px transparent;
}

/**
 * chat_box
 */
 .chat_box {
	min-height: 300px;
	/*
    height: 380px;
	min-height: 56px;
	min-width: 299px;
	*/
	padding: 10px 10px 5px 10px;
	z-index: 1005;
}

.chat_box .chat_wrap {
	float: right;
    width: 508px;
}
	.chat_wrap .rndBtn {
		float: left;
		margin: 10px 10px 10px 0;
		display: block;
	}

.tabBarCtn {
	position: relative;
	border-top: 1px solid #D6D6D6;
	border-left: 1px solid #FFFFFF;
	border-right: 1px solid #E2E2E2;
}
	.chatTabBar.tabBarCtn, .chatTabBar.tabBarCtn:hover {
		height: 60px;
		margin-bottom: 10px;
		background: none repeat scroll 0 0 #FCFCFC;
	}

	.tabBarCtn .tabBtn {
		position: relative;
		float: left;
		width: 90px;
		height: 60px;
		border-left: 1px solid #D6D6D6;
		border-bottom: 1px solid #D6D6D6;
		-moz-transition: width 300ms ease-out 0s;
		display: block;
	}
	.tabBarCtn .tabBtn.crt {
		border-bottom: 1px solid #8EB800;
		width: 145px;
	}
	.tabBtn .tabBtnImg {
		background: url("../images/icons_chat_bar.png") no-repeat scroll center center transparent;
		height: 60px;
		left: 50%;
		margin-left: -20px;
		position: absolute;
		width: 40px;
	}
	.tabBtn.messageCt .tabBtnImg {
		background-position: 2px 0;
	}
	.tabBtn.voteCt .tabBtnImg {
		background-position: -284px 0;
	}

	.tabBarCtn .tabMenuCtn {
		border-left: 1px solid #D6D6D6;
		border-bottom: 1px solid #D6D6D6;
	}
	
	.chatEvent .tpsMenuCtn {
		position: absolute;
		right: 12px;
		top: 8px;
	}
		.tpsMenuCtn .rndBtn {
			margin-right: 4px;
		}
	
.scrollable{
	border-top:1px solid #e2e2e2;
	border-right:1px solid #e2e2e2;
	background:#fcfcfc;
}
.jspScrollable{
	outline: none 0;
}
	.chatCtn .jspScrollable {
		width: 505px;
		height: 282px;
		padding: 0px; 
		overflow: hidden; 
	}

.jspContainer{
	position: relative;
	overflow: hidden;
}
	.chatCtn .jspContainer {
		width: 505px; 
		/*
		height: 132px;
		*/
		height: 282px;
		overflow-y: auto;
	}

.jspPane{
	position: absolute;
}
	.chatCtn .jspPane{
		top: 0px;
		width: 485px; 
		padding: 0px; 
	}

.rctCtn {
	position: relative;
	float: left;
	height: 86px;
	width: 466px;
	border-left: 1px solid #FFFFFF;
	border-right: 1px solid #E2E2E2;
	border-top: 1px solid #D6D6D6;
	background: none repeat scroll 0 0 #FCFCFC;
}
	.chatCtn .rctCtn {
		float: none;
		height: 65px;
		width: 504px;
	}
	.jspPane .rctCtn {
		/*
		float: none;
		height: 65px;
		*/
		width: 484px;
	}
	.chatCtn .chtCtn, .jspPane .chtCtn {
		float: left;
		height: auto;
		border-right: medium none;
		border-top: medium none;
	}

.chatContent  { padding:10px; width:100%;}
.you { float:left; width:100%; /*ie6 hack*/_background:none; _border:none;}
.me { float:right; width:100%; }
.chatItem { padding:4px 0px 10px 0px;_background:none; _border:none; }
.chatItemContent{
	cursor:pointer;
}
.cloudPannel{
	position: relative;
	_position:static;
}

.you { float:left; width:100%; /*ie6 hack*/_background:none; _border:none;}
.me { float:right; width:100%; }
.chatItem { 
	position: relative;
	float: left;
	padding: 4px 0px 10px 0px;
	_background: none; 
	_border:none; 
}


.media {
	margin: 10px auto;
	width: 365px;
	border: 1px solid #AEB4B9;
	box-shadow:0px 1px 1px #D7D7D7; 
	-webkit-border-radius:5px;
	-moz-border-radius:5px;
	border-radius:5px;
	background-color:#FAFAFA;
	background:-webkit-gradient(linear,
					left top,left bottom,
					from(#FEFEFE),to(#FAFAFA));
	background-image:-moz-linear-gradient(top, #FEFEFE 0%,#FAFAFA 100%);
}
.media a {
	display: block;
}
.media .mediaContent {
  margin: 0;
  padding: 0;
}

.media .mediaPanel{
		padding:0px;
		margin:0px;
	}
	.media .mediaImg{
		margin:6px 0px -22px;
	}
	.media .mediaImg .mediaImgPanel{
		position:relative;
		padding:0px;
		margin:0px;
		height:164px;
		width:100%;
		overflow:hidden;
	}
	.media .mediaImg img{
		/* width:100%;
		height:164px;*/
		position: absolute;
		left: 0px;
		max-width: 365px;
		/*
		max-height: 295px;
		*/
	}
	.media .mediaImg .mediaImgFooter{
		position:relative;
		top: -29px;
		height:29px;
		background-color:#000;
		background-color:rgba(0,0,0,0.4);
		text-shadow:none;
		color:#FFF;
		text-align:left;
		padding:0px 11px;
		line-height:29px;
	}
	.media .mediaImg .mediaImgFooter a:hover p{
		color:#B8B3B3;
	}
	.media .mediaImg .mediaImgFooter .mesgTitleTitle{
		line-height:28px;
		color:#FFF;
		max-width:318px;
		height:26px;
		white-space:nowrap;
		text-overflow:ellipsis;
		-o-text-overflow:ellipsis;
		overflow:hidden;
	}
	.media .mesgIcon{
		display:inline-block;
		height:13px;
		width: 25px;
		margin:8px 0px -2px 4px;
	}
	.media .mediaContent{
		margin:0px;
		padding:0px;
	}
	.media .mediaContent .mediaMesg{
		border-top:1px solid #D7D7D7;
		padding:0px 10px;
	}
	.media .mediaContent .mediaMesg .mediaMesgDot{
		display: block;
		position:relative;
		top: -3px;
		left:20px;
		height:6px;
		width:6px;
		-moz-border-radius: 3px;
		-webkit-border-radius: 3px;
		border-radius: 3px;
	}
	.media .mediaContent .mediaMesg .mediaMesgTitle:hover p{
		color:#1A1717;
	}
	.media .mediaContent .mediaMesg .mediaMesgTitle a{
		color:#707577;
	}
	.media .mediaContent .mediaMesg .mediaMesgTitle a:hover p{
		color:#444440;
	} 
	.media .mediaContent .mediaMesg .mediaMesgIcon{
	}
	.media .mediaContent .mediaMesg .mediaMesgTitle p{
		line-height:1.5em;
		max-height: 45px;
		max-width: 286px;
		min-width:176px;
		margin-top:2px;
		color:#5D6265;
		text-overflow:ellipsis;
		-o-text-overflow:ellipsis;
		overflow:hidden;
		text-align: left;
		text-overflow:ellipsis;
	}
	.media .mediaContent .mediaMesg .mediaMesgIcon img{
		height:45px;
		width:45px;
	}
	/*media mesg detail*/
	.media .mediaHead{
		/*height:48px;*/
		padding:0px 8px 4px;
		border-bottom:1px solid #D3D8DC;
		color:#A51000;
		font-size:16px;
	}
	.media .mediaHead .title{
		line-height:1.2em;
		margin-top: 22px;
		display:block;
		max-width:312px;
		text-align: left;
		/*height:25px;
		white-space:nowrap;
		text-overflow:ellipsis;
		-o-text-overflow:ellipsis;
		overflow:hidden;*/
	}
	.mediaFullText .mediaImg{
		height:164px;
		width:100%;
		padding:0px 0px 5px;
		margin:0px;
		margin-top:17px;
		overflow:hidden;
		position:relative;
	}
	.mediaFullText .mediaImg img{
		margin-top:17px;
		position:absolute;
	}
	.mediaFullText .mediaContent{
		padding:6px 8px 10px;
		font-size:12px;
		line-height: 1.5em;
		text-align:left;
		color:#666B6E;
	}
	.mediaFullText .mediaContentP{
		padding:12px 8px 10px;
	}
	.media .mediaHead .time{
		margin:0px;
		margin-top: 21px;
		color:#82888C;
		background:none;
		width:auto;
	}
	.media .mediaFooter{
		background-color:#F0F4F8;
		-webkit-border-radius:0px 0px 5px 5px;
		-moz-border-radius:0px 0px 5px 5px;
		border-radius:0px 0px 5px 5px;
	}
	.media .mediaFooter a{
		color:#792F2E;
		font-size:14px;
		padding:0px 7px;
		width: 100%;
	}
	.media .mediaFooter .mesgIcon{
		margin:12px 3px 0px 0px;
	}
	.media a:hover{
		cursor: pointer;
	}
	.media a:hover .mesgIcon {
		width: 25px;
		/* background:url("../images/button_chat13dfb3.png") no-repeat -188px -987px;	*/
	}
</style>
<div id="BoxTop" style="width:56%;"><span>预览</span></div>
<div id="BoxContents" style="width:56%;">
<div class="chat_box">

	<div class="chat_wrap">

<div class="chatFrom tabBarCtn">
	<!-- chat of start -->
	<div class="tabMenuCtn chatCtn">


		<!-- scrollable of start -->
		<div class="scrollable jspScrollable" id="chtScroll" tabindex="0">
			<div class="jspContainer">
				<div class="jspPane">
<div class="chatItem you">
	<div class="media">
		<div class="mediaPanel">
		<a href="#">
				<div class="mediaImg">
										<div class="mediaImgPanel">
						<img onerror="this.parentNode.removeChild(this)" src="<%=Fk_Sucai_pic%>" />
					</div>
										<div class="mediaImgFooter">
						<p class="mesgTitleTitle left"><%=Fk_Sucai_Title%></p>
						<div class="clr"></div>
					</div>
				</div>
				</a>
			
					<%if Fk_Sucai_Id_List<>"" then
					set rs=conn.execute("select * from weixin_Sucai where id in("&Fk_Sucai_Id_List&") order by Sucai_px desc")
					if not rs.eof then%>
					<div class="mediaContent">
					<%do while not rs.eof%>
					<a href="#">
					<div class="mediaMesg">
						<span class="mediaMesgDot"></span>
						<div class="mediaMesgTitle left">
							<p class="left"><%=rs("Sucai_Title")%></p>
							<div class="clr"></div>
						</div>
						<div class="mediaMesgIcon right">
							<img onerror="this.parentNode.removeChild(this)" src="<%=rs("Sucai_pic")%>" />
						</div>
						<div class="clr"></div>
					</div>
					</a>
					<%
					rs.movenext
					loop%>
					</div>
					<%end if
					rs.close
					end if
					%>
		</div>
	</div>
</div>

				</div>
			</div>
		</div>
		<!-- scrollable of over -->
	</div>
	<!-- chat of over -->

</div>

	</div>
</div>
</div>
<div id="BoxBottom" style="width:54%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/weixin_Sucai.asp?Type=5',0,'',0,1,'MainRight','/admin/weixin/weixin_Sucai.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
<%
End Sub
%><!--#Include File="../../Code.asp"-->