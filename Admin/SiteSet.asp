<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名SiteSet.asp
'文件用途站点设置拉取页面
'版权所有企帮网络www.qebang.cn
'==========================================

'判断权限
If Not FkFun.CheckLimit("System1") Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

'获取参数
Types=Clng(Request.QueryString("Type"))
dim Snr,curhost
Snr=Clng(Request.QueryString("Snr"))
curhost=request.ServerVariables("HTTP_HOST")

Select Case Types
	Case 1
		Call SiteSetBox() '读取系统信息
	Case 2
		Call SiteSetDo() '系统设置操作
	Case 3
		Call TestFetion() '测试飞信
End Select

'==========================================
'函 数 名SiteSetBox()
'作    用读取系统信息
'参    数
'==========================================
Sub SiteSetBox()
	dim mHosts,toUrls,arrSite301
	if instr(Site301,"|")>0 Then 
		arrSite301=split(Site301,"|")
		mHosts=arrSite301(0)
		toUrls=arrSite301(1)
	else
		mHosts=""
		toUrls=""
	end if
%>
<link href="/admin/dkidtioenr/themes/default/default.css" rel="stylesheet" type="text/css" />
<script type="text/javascript">
$(document).ready(function(){

	if(window.KindEditor){
		if(("#SiteLogo").length>0){
			$("#SiteLogo").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
				var editor = window.KindEditor.editor({
						fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
						uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
						allowFileManager : true
					});
					$('#uploadButton').click(function() {
						editor.loadPlugin('image', function() {
							editor.plugin.imageDialog({
								imageUrl : $('#SiteLogo').val(),
								clickFn : function(url) {
									$('#SiteLogo').val(url);
									editor.hideDialog();
								}
							});
						});
					});
	
		}
		if(("#Sitepic1").length>0){
		
				$("#Sitepic1").after(" <input type=\"button\" id=\"uploadButton1\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
		
					var editor = window.KindEditor.editor({
							fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
							uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
							allowFileManager : true
						});
						$('#uploadButton1').click(function() {
							editor.loadPlugin('image', function() {
								editor.plugin.imageDialog({
									imageUrl : $("#Sitepic1").val(),
									clickFn : function(url) {
										$("#Sitepic1").val(url);
										editor.hideDialog();
									}
								});
							});
						});
						
						
				$("#Sitepic2").after(" <input type=\"button\" id=\"uploadButton2\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
		
					var editor = window.KindEditor.editor({
							fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
							uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
							allowFileManager : true
						});
						$('#uploadButton2').click(function() {
							editor.loadPlugin('image', function() {
								editor.plugin.imageDialog({
									imageUrl : $("#Sitepic2").val(),
									clickFn : function(url) {
										$("#Sitepic2").val(url);
										editor.hideDialog();
									}
								});
							});
						});
				
				
				$("#Sitepic3").after(" <input type=\"button\" id=\"uploadButton3\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
		
					var editor = window.KindEditor.editor({
							fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
							uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
							allowFileManager : true
						});
						$('#uploadButton3').click(function() {
							editor.loadPlugin('image', function() {
								editor.plugin.imageDialog({
									imageUrl : $("#Sitepic3").val(),
									clickFn : function(url) {
										$("#Sitepic3").val(url);
										editor.hideDialog();
									}
								});
							});
						});
						
						
					$("#Sitepic4").after(" <input type=\"button\" id=\"uploadButton4\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
		
					var editor = window.KindEditor.editor({
							fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
							uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
							allowFileManager : true
						});
						$('#uploadButton4').click(function() {
							editor.loadPlugin('image', function() {
								editor.plugin.imageDialog({
									imageUrl : $("#Sitepic4").val(),
									clickFn : function(url) {
										$("#Sitepic4").val(url);
										editor.hideDialog();
									}
								});
							});
						});
						
						
					$("#Sitepic5").after(" <input type=\"button\" id=\"uploadButton5\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
		
					var editor = window.KindEditor.editor({
							fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
							uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
							allowFileManager : true
						});
						$('#uploadButton5').click(function() {
							editor.loadPlugin('image', function() {
								editor.plugin.imageDialog({
									imageUrl : $("#Sitepic5").val(),
									clickFn : function(url) {
										$("#Sitepic5").val(url);
										editor.hideDialog();
									}
								});
							});
						});
			
			
		}
	}
	else{
		if(("#SiteLogo").length>0){
			$("#SiteLogo").after(" <iframe frameborder=\"0\" width=\"290\" height=\"25\" scrolling=\"No\" id=\"I2\" name=\"I2\" src=\"PicUpLoad.asp?Form=SystemSet&Input=SiteLogo\" style=\"vertical-align:middle\"></iframe> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许200K");
		}
		if(("#Sitepic1").length>0){
			for(var i=1;i<=5;i++){
				$("#Sitepic"+i).after(" <iframe frameborder=\"0\" width=\"290\" height=\"25\" scrolling=\"No\" id=\"I2\" name=\"I2\" src=\"PicUpLoad.asp?Form=SystemSet&Input=Sitepic"+i+"\" style=\"vertical-align:middle\"></iframe> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许200K");
			}		
		}
	}
})
	
</script>

<form id="SystemSet" name="SystemSet" method="post" action="SiteSet.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span><%
if Snr=1 then
response.write "基础信息"
else
response.write "广告橱窗"
end if
%></span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
    <div style=" text-align:center; display:none;"><a href="javascript:void(0);" onclick="document.getElementById('table001').style.display='block';document.getElementById('table002').style.display='none'">基础信息</a>　|　<a href="javascript:void(0);" onclick="document.getElementById('table002').style.display='block';document.getElementById('table001').style.display='none'">幻灯图片</a></div>
	<table width="90%" border="0" align="center" id="table001" style="display:<%
	if Snr=1 then
	response.write "block"
	else
	response.write "none"
	end if
	%>" cellpadding="0" cellspacing="0">
	 <tr>
            <td height="25" align="right" class="MainTableTop">LOGO标志</td>
            <td colspan="3"><%if SiteLogo<>"" then%><img onclick="this.src=document.getElementById('SiteLogo').value;" src="<%=SiteLogo%>" style="width:136px;cursor:pointer;float:right;vertical-align:middle" title="LOGO预览 "><%end if%><input name="SiteLogo" type="text" class="Input" id="SiteLogo" value="<%=SiteLogo%>" size="28" style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">公司名称</td>
            <td><input name="SiteName" type="text" class="Input" id="SiteName" value="<%=SiteName%>" size="50"  style="vertical-align:middle"/></td>
            <td>域名</td>
            <td><input name="SiteUrl" type="text" class="Input" id="SiteUrl" value="<%=SiteUrl%>" size="32"  style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">电话</td>
            <td><input name="Tel" type="text" class="Input" id="Tel" value="<%=Tel%>" size="32"  style="vertical-align:middle"/></td>
            <td>400热线</td>
            <td>
			<input name="Tel400" type="text" class="Input" id="Tel400" value="<%=Tel400%>" size="32"  style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">Email</td>
            <td><input name="Email" type="text" class="Input" id="Email" value="<%=Email%>" size="32"  style="vertical-align:middle"/></td>
            <td>传真</td>
            <td>
			<input name="Fax" type="text" class="Input" id="Fax" value="<%=Fax%>" size="32"  style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">联系人</td>
            <td><input name="Lianxiren" type="text" class="Input" id="Lianxiren" value="<%=Lianxiren%>" size="32"  style="vertical-align:middle"/></td>
            <td>备案号</td>
            <td>
			<input name="beian" type="text" class="Input" id="beian" value="<%=beian%>" size="32"  style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">地址</td>
            <td colspan="3"><input name="Add" type="text" class="Input" id="Add" value="<%=Add%>" size="110"  style="vertical-align:middle"/>&nbsp;</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">SEO标题</td>
            <td colspan="3"><input name="SiteSeoTitle" type="text" class="Input" id="SiteSeoTitle" value="<%=SiteSeoTitle%>" size="110"  style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">SEO关键词</td>
            <td colspan="3"><input name="SiteKeyword" type="text" class="Input" id="SiteKeyword" value="<%=SiteKeyword%>" size="110"  style="vertical-align:middle"/></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">SEO描述</td>
            <td colspan="3"><input name="SiteDescription" type="text" class="Input" id="SiteDescription" value="<%=SiteDescription%>" size="110"  style="vertical-align:middle"/></td>
        </tr>
        <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">非法字符过滤</td>
            <td>
			<input name="SiteDelWord" class="Input" type="radio" id="SiteDelWord" value="1"<%=FKFun.BeCheck(SiteDelWord,1)%>  style="vertical-align:middle"/>开启
            <input type="radio" name="SiteDelWord" class="Input" id="SiteDelWord" value="0"<%=FKFun.BeCheck(SiteDelWord,0)%>  style="vertical-align:middle"/>关闭</td>
            <td>每页条数</td>
            <td><select class="Input" name="PageSizes" id="PageSizes" style="vertical-align:middle">
<%
	For i=1 To 50
%>
                    <option value="<%=i%>"<%=FKFun.BeSelect(PageSizes,i)%>><%=i%>条</option>
<%
	Next
%>
                </select></td>
        </tr>
        <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">在线客服账号</td>
            <td><input name="Kfid" type="text" class="Input" id="Kfid" value="<%If curhost<>"localhost" and curhost<>"127.0.0.1" Then response.write Kfid Else response.write "anyya" End if%>" size="15" readonly style="vertical-align:middle"/><input class="Button" type="button" value="自动获取" onclick="autoid('kfid');" <%If curhost="localhost" or curhost="127.0.0.1" Then response.write "disabled"%> name="B3" style="vertical-align:middle"><%If curhost="localhost" or curhost="127.0.0.1" Then response.write " <b style='color:red'>传到服务器后才能获取</b>" %>
</td>
            <td>统计账号</td>
            <td>
			<input name="Tjid" type="text" class="Input" id="Tjid" value="<%If curhost<>"localhost" and curhost<>"127.0.0.1" Then response.write Tjid Else response.write "544" End if%>" size="15" readonly style="vertical-align:middle"/><input class="Button" type="button" value="自动获取" onclick="autoid('tjid');"  <%If curhost="localhost" or curhost="127.0.0.1" Then response.write "disabled"%> name="B3" style="vertical-align:middle"><%If curhost="localhost" or curhost="127.0.0.1" Then response.write " <b style='color:red'>传到服务器后才能获取</b>" %></td>
        </tr>
        <tr <%call AdminY()%> style="display:none">
            <td height="25" align="right" class="MainTableTop">飞信号码</td>
            <td><input name="FetionNum" type="text" class="Input" id="FetionNum" value="<%=FetionNum%>" size="20"  style="vertical-align:middle"/>&nbsp;&nbsp;&nbsp;<span id="Test" style="color:#F00;"><a href="javascript:void(0);" onclick="SetRContent('Test','SiteSet.asp?Type=3')">设置后测试</a></span></td>
            <td>飞信密码</td>
            <td><input name="FetionPass" type="password" class="Input" id="FetionPass" value="<%=FetionPass%>" size="20"  style="vertical-align:middle"/></td>
        </tr>
        
        <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">提取缩略</td>
            <td><input name="SiteMini" type="text" class="Input" id="SiteMini" value="<%=SiteMini%>" size="20"  style="vertical-align:middle"/>个字符</td>
            <td>站点开放</td>
            <td>
			<input name="SiteOpen" class="Input" type="radio" id="SiteOpen" value="1"<%=FKFun.BeCheck(SiteOpen,1)%>  style="vertical-align:middle"/>开放
            <input type="radio" name="SiteOpen" class="Input" id="SiteOpen" value="0"<%=FKFun.BeCheck(SiteOpen,0)%>  style="vertical-align:middle"/>关闭</td>
        </tr>
        <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">模板</td>
            <td><select class="Input" name="SiteTemplate" id="SiteTemplate" style="vertical-align:middle">
    <%
	Dim ObjFloders,ObjFloder
	Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
	Set F=Fso.GetFolder(Server.MapPath("../Skin/"))
	Set ObjFloders=F.Subfolders
	For Each ObjFloder In ObjFloders
%>
                <option value="<%=ObjFloder.Name%>"<%=FKFun.BeSelect(SiteTemplate,ObjFloder.Name)%>><%=ObjFloder.Name%></option>
    <%
	Next
	Set ObjFloders=Nothing
	Set F=Nothing
	Set Fso=Nothing
%>
                </select></td>
            <td>反垃圾</td>
            <td>
			<input name="SiteNoTrash" class="Input" type="radio" id="SiteNoTrash" value="1"<%=FKFun.BeCheck(SiteNoTrash,1)%>  style="vertical-align:middle"/>开启
            <input type="radio" name="SiteNoTrash" class="Input" id="SiteNoTrash" value="0"<%=FKFun.BeCheck(SiteNoTrash,0)%>  style="vertical-align:middle"/>关闭</td>
        </tr>
        <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">文件名生成</td>
            <td><input name="SiteToPinyin" class="Input" type="radio" id="SiteToPinyin" value="1"<%=FKFun.BeCheck(SiteToPinyin,1)%>  style="vertical-align:middle"/>自动生成拼音
            <input type="radio" name="SiteToPinyin" class="Input" id="SiteToPinyin" value="0"<%=FKFun.BeCheck(SiteToPinyin,0)%>  style="vertical-align:middle"/>不自动生成</td>
            <td>客服浮窗</td>
            <td>
			<input name="SiteQQ" class="Input" type="radio" id="SiteQQ" value="1"<%=FKFun.BeCheck(SiteQQ,1)%>  style="vertical-align:middle"/>开启
            <input type="radio" name="SiteQQ" class="Input" id="SiteQQ" value="0"<%=FKFun.BeCheck(SiteQQ,0)%>  style="vertical-align:middle"/>关闭</td>
        </tr>
	    <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">是否生成</td>
            <td>
			<input name="SiteHtml" class="Input" type="radio" id="SiteHtml" value="1"<%=FKFun.BeCheck(SiteHtml,1)%>  style="vertical-align:middle"/>生成
                <input type="radio" name="SiteHtml" class="Input" id="SiteHtml" value="0"<%=FKFun.BeCheck(SiteHtml,0)%>  style="vertical-align:middle"/>不生成</td>
            <td>模板调试</td>
            <td>
			<input name="SiteTest" class="Input" type="radio" id="SiteTest" value="1"<%=FKFun.BeCheck(SiteTest,1)%> style="vertical-align:middle"/>开启
            <input type="radio" name="SiteTest" class="Input" id="SiteTest" value="0"<%=FKFun.BeCheck(SiteTest,0)%>  style="vertical-align:middle"/>关闭</td>
        </tr>
        <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">Index后缀</td>
            <td><input name="SiteFlash" class="Input" type="radio" id="SiteFlash" value="1"<%=FKFun.BeCheck(SiteFlash,1)%> style="vertical-align:middle"/>开启
            <input type="radio" name="SiteFlash" class="Input" id="SiteFlash" value="0"<%=FKFun.BeCheck(SiteFlash,0)%>  style="vertical-align:middle"/>关闭</td>
            <td>编辑器</td>
            <td><select class="Input" name="Bianjiqi" id="Bianjiqi" style="vertical-align:middle">
                    <option value="xheditor"<%=FKFun.BeSelect(Bianjiqi,"xheditor")%>>xheditor</option>
                    <option value="kinediter"<%=FKFun.BeSelect(Bianjiqi,"kinediter")%>>kinediter</option>
                    <option value="ueditor"<%=FKFun.BeSelect(Bianjiqi,"ueditor")%>>ueditor</option>
                </select></td>
        </tr>
        <tr <%call AdminY()%>>
            <td height="25" align="left" class="MainTableTop" colspan="4">301跳转设置： 跳转至域名 <input type="text" class="Input" id="mainDomain" name="mainDomain" size="15"  value="<%=mHosts%>" style="vertical-align:middle"/>
              ，需跳转域名 <input name="Site301" size="30" class="Input" type="text" id="Site301" value="<%=toUrls%>" style="vertical-align:middle;"/> 
            填写正式域名,多个以英文逗号隔开,格式:qebang.cn</td>
        </tr>
	    <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">图片资源CDN</td>
            <td>
			<input type="text" class="Input" id="ImgCdnUrl" name="ImgCdnUrl" size="45"  value="<%=ImgCdnUrl%>" style="vertical-align:middle"/></td>
            <td>JS资源CDN</td>
            <td>
			<input type="text" class="Input" id="JsCdnUrl" name="JsCdnUrl" size="45"  value="<%=JsCdnUrl%>" style="vertical-align:middle"/></td>
        </tr>
	    <tr <%call AdminY()%>>
            <td height="25" align="right" class="MainTableTop">CSS资源CDN</td>
            <td>
			<input type="text" class="Input" id="CssCdnUrl" name="CssCdnUrl" size="45"  value="<%=CssCdnUrl%>" style="vertical-align:middle"/></td>
            <td>文件资源CDN</td>
            <td>
			<input type="text" class="Input" id="FileCdnUrl" name="FileCdnUrl" size="45"  value="<%=FileCdnUrl%>" style="vertical-align:middle"/></td>
        </tr>
    </table>
	<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0" id="table002" style="display:<%
	if Snr=1 then
	response.write "none"
	else
	response.write "block"
	end if
	%>" >
      <tr>
            <td align="right" rowspan="5"></td>
            <td colspan="3"><%if Sitepic1<>"" then%><img src="<%=Sitepic1%>" onclick="this.src=document.getElementById('Sitepic1').value;" style="cursor:hand;max-height:100px;width:300px;float:right;padding:1px;border:1px #CCCCCC solid;" title="小图预览[上传后点击更新] "><%end if%>图一：<input name="Sitepic1" type="text" class="Input sitepics" id="Sitepic1" value="<%=Sitepic1%>" size="28"  style="vertical-align:middle"/><br>
			链接：<input name="Sitepicurl1" type="text" class="Input" id="Sitepicurl1" value="<%=Sitepicurl1%>" size="50"  style="vertical-align:middle"/>&nbsp;<br>
			文字：<input name="Sitepictext1" type="text" class="Input" id="Sitepictext1" value="<%=Sitepictext1%>" size="50"  style="vertical-align:middle"/></td>
        </tr>
         <tr>
            <td colspan="3"><%if Sitepic2<>"" then%><img src="<%=Sitepic2%>" onclick="this.src=document.getElementById('Sitepic2').value;"  style="cursor:hand;width:300px;max-height:100px;float:right;padding:1px;border:1px #CCCCCC solid;" title="小图预览[上传后点击更新] "><%end if%>图二：<input name="Sitepic2" type="text" class="Input sitepics" id="Sitepic2" value="<%=Sitepic2%>" size="28"  style="vertical-align:middle"/><br>
			链接：<input name="Sitepicurl2" type="text" class="Input" id="Sitepicurl2" value="<%=Sitepicurl2%>" size="50"  style="vertical-align:middle"/><br>
			文字：<input name="Sitepictext2" type="text" class="Input" id="Sitepictext2" value="<%=Sitepictext2%>" size="50"  style="vertical-align:middle"/></td>
        </tr>
         <tr>
            <td colspan="3"><%if Sitepic3<>"" then%><img src="<%=Sitepic3%>" onclick="this.src=document.getElementById('Sitepic3').value;"  style="cursor:hand;width:300px;max-height:100px;float:right;padding:1px;border:1px #CCCCCC solid;" title="小图预览[上传后点击更新] "><%end if%>图三：<input name="Sitepic3" type="text" class="Input sitepics" id="Sitepic3" value="<%=Sitepic3%>" size="28"  style="vertical-align:middle"/><br>
			链接：<input name="Sitepicurl3" type="text" class="Input" id="Sitepicurl3" value="<%=Sitepicurl3%>" size="50"  style="vertical-align:middle"/><br>
			文字：<input name="Sitepictext3" type="text" class="Input" id="Sitepictext3" value="<%=Sitepictext3%>" size="50"  style="vertical-align:middle"/></td>
        </tr>
         <tr <%call AdminY()%>>
            <td colspan="3"><%if Sitepic4<>"" then%><img src="<%=Sitepic4%>" onclick="this.src=document.getElementById('Sitepic4').value;"  style="cursor:hand;width:300px;max-height:100px;float:right;padding:1px;border:1px #CCCCCC solid;vertical-align:middle" title="小图预览[上传后点击更新] "><%end if%>图四：<input name="Sitepic4" type="text" class="Input sitepics" id="Sitepic4" value="<%=Sitepic4%>" size="28"  style="vertical-align:middle"/><br>
			链接：<input name="Sitepicurl4" type="text" class="Input" id="Sitepicurl4" value="<%=Sitepicurl4%>" size="50"  style="vertical-align:middle"/><br>
			文字：<input name="Sitepictext4" type="text" class="Input" id="Sitepictext4" value="<%=Sitepictext4%>" size="50"  style="vertical-align:middle"/></td>
        </tr>
         <tr <%call AdminY()%>>
            <td colspan="3"><%if Sitepic5<>"" then%><img src="<%=Sitepic5%>" onclick="this.src=document.getElementById('Sitepic5').value;"  style="cursor:hand;width:300px; max-height:100px; float:right;padding:1px;border:1px #CCCCCC solid;vertical-align:middle" title="小图预览[上传后点击更新] "><%end if%>图五：<input name="Sitepic5" type="text" class="Input sitepics" id="Sitepic5" value="<%=Sitepic5%>" size="28"  style="vertical-align:middle"/><br>
			链接：<input name="Sitepicurl5" type="text" class="Input" id="Sitepicurl5" value="<%=Sitepicurl5%>" size="50"  style="vertical-align:middle"/><br>
			文字：<input name="Sitepictext5" type="text" class="Input" id="Sitepictext5" value="<%=Sitepictext5%>" size="50"  style="vertical-align:middle"/></td>
        </tr>
        
    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('SystemSet','SiteSet.asp?Type=2',0,'',0,0,'','');" class="Button" name="button" id="button" value="保存设置" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名SiteSetDo()
'作    用系统设置操作
'参    数
'==========================================
Sub SiteSetDo()
	Dim OldSiteTemplate,ObjFile,ObjFiles,mainDomain
	OldSiteTemplate=SiteTemplate
	SiteName=FKFun.HTMLEncode(Trim(Request.Form("SiteName")))
	SiteSeoTitle=FKFun.HTMLEncode(Trim(Request.Form("SiteSeoTitle")))
	SiteKeyword=FKFun.HTMLEncode(Trim(Request.Form("SiteKeyword")))
	SiteDescription=FKFun.HTMLEncode(Trim(Request.Form("SiteDescription")))
	SiteUrl=FKFun.HTMLEncode(Trim(Request.Form("SiteUrl")))
	FetionNum=FKFun.HTMLEncode(Trim(Request.Form("FetionNum")))
	FetionPass=FKFun.HTMLEncode(Trim(Request.Form("FetionPass")))
	SiteOpen=Trim(Request.Form("SiteOpen"))
	SiteHtml=Trim(Request.Form("SiteHtml"))
	SiteTemplate=FKFun.HTMLEncode(Trim(Request.Form("SiteTemplate")))
	PageSizes=Trim(Request.Form("PageSizes"))
	SiteToPinyin=Trim(Request.Form("SiteToPinyin"))
	SiteQQ=Trim(Request.Form("SiteQQ"))
	SiteNoTrash=Trim(Request.Form("SiteNoTrash"))
	SiteMini=Trim(Request.Form("SiteMini"))
	SiteDelWord=Trim(Request.Form("SiteDelWord"))
	'SiteTest=Trim(Request.Form("SiteTest"))
	SiteTest=1
	SiteFlash=Trim(Request.Form("SiteFlash"))
	mainDomain=Trim(Request.Form("mainDomain"))
	Site301=Trim(Request.Form("Site301"))
	'自定义增加部分start---------------------------------
	Tel=FKFun.HTMLEncode(Trim(Request.Form("Tel")))
	Tel400=FKFun.HTMLEncode(Trim(Request.Form("Tel400")))
	Fax=FKFun.HTMLEncode(Trim(Request.Form("Fax")))
	Lianxiren=FKFun.HTMLEncode(Trim(Request.Form("Lianxiren")))
	Email=FKFun.HTMLEncode(Trim(Request.Form("Email")))
	Beian=FKFun.HTMLEncode(Trim(Request.Form("Beian")))
	Add=FKFun.HTMLEncode(Trim(Request.Form("Add")))
	Tjid=FKFun.HTMLEncode(Trim(Request.Form("Tjid")))
	Kfid=FKFun.HTMLEncode(Trim(Request.Form("Kfid")))
	SiteLogo=FKFun.HTMLEncode(Trim(Request.Form("SiteLogo")))
	Sitepic1=FKFun.HTMLEncode(Trim(Request.Form("Sitepic1")))
	Sitepicurl1=FKFun.HTMLEncode(Trim(Request.Form("Sitepicurl1")))
	Sitepic2=FKFun.HTMLEncode(Trim(Request.Form("Sitepic2")))
	Sitepicurl2=FKFun.HTMLEncode(Trim(Request.Form("Sitepicurl2")))
	Sitepic3=FKFun.HTMLEncode(Trim(Request.Form("Sitepic3")))
	Sitepicurl3=FKFun.HTMLEncode(Trim(Request.Form("Sitepicurl3")))
	Sitepic4=FKFun.HTMLEncode(Trim(Request.Form("Sitepic4")))
	Sitepicurl4=FKFun.HTMLEncode(Trim(Request.Form("Sitepicurl4")))
	Sitepic5=FKFun.HTMLEncode(Trim(Request.Form("Sitepic5")))
	Sitepicurl5=FKFun.HTMLEncode(Trim(Request.Form("Sitepicurl5")))
	Sitepictext1=FKFun.HTMLEncode(Trim(Request.Form("Sitepictext1")))
	Sitepictext2=FKFun.HTMLEncode(Trim(Request.Form("Sitepictext2")))
	Sitepictext3=FKFun.HTMLEncode(Trim(Request.Form("Sitepictext3")))
	Sitepictext4=FKFun.HTMLEncode(Trim(Request.Form("Sitepictext4")))
	Sitepictext5=FKFun.HTMLEncode(Trim(Request.Form("Sitepictext5")))
	
	Bianjiqi=FKFun.HTMLEncode(Trim(Request.Form("Bianjiqi")))
	
	ImgCdnUrl=FKFun.HTMLEncode(Trim(Request.Form("ImgCdnUrl")))
	JsCdnUrl=FKFun.HTMLEncode(Trim(Request.Form("JsCdnUrl")))
	CssCdnUrl=FKFun.HTMLEncode(Trim(Request.Form("CssCdnUrl")))
	FileCdnUrl=FKFun.HTMLEncode(Trim(Request.Form("FileCdnUrl")))
	'end---------------------------------------------
	Call FKFun.ShowString(SiteName,1,50,0,"请输入站点名称！","站点名称不能大于50个字符！")
	
	Call FKFun.ShowString(Tel,1,50,0,"请输入电话号码！","电话号码不能大于50个字符！")
	Call FKFun.ShowString(Tel400,1,50,0,"请输入400热线号码！","400热线号码不能大于50个字符！")
	Call FKFun.ShowString(Fax,1,50,0,"请输入传真号码！","传真号码不能大于50个字符！")
	Call FKFun.ShowString(Lianxiren,1,50,0,"请输入联系人！","联系人不能大于50个字符！")
	Call FKFun.ShowString(Email,1,50,0,"请输入Email！","Email不能大于50个字符！")
	Call FKFun.ShowString(Beian,1,50,0,"请输入备案号！","备案号不能大于50个字符！")
	Call FKFun.ShowString(Add,1,200,0,"请输入地址！","地址不能大于200个字符！")
	Call FKFun.ShowString(Tjid,1,50,0,"请输入统计账号！","统计账号不能大于50个字符！")
    Call FKFun.ShowString(Kfid,1,50,0,"请输入客服账号！","客服账号不能大于50个字符！")

	Call FKFun.ShowString(SiteSeoTitle,1,255,0,"请输入站点Seo标题！","站点Seo标题不能大于255个字符！")
	Call FKFun.ShowString(SiteKeyword,1,255,0,"请输入站点关键字！","站点关键字不能大于255个字符！")
	Call FKFun.ShowString(SiteDescription,1,255,0,"请输入站点介绍！","站点描述不能大于255个字符！")
	Call FKFun.ShowString(SiteLogo,1,255,0,"您还未设置LOGO","请点击选择上传LOGO")
	Call FKFun.ShowString(SiteUrl,1,255,0,"请输入站点域名！","站点域名不能大于255个字符！")
	Call FKFun.ShowString(ImgCdnUrl,0,255,0,"请输入图片资源CDN地址！","图片资源CDN地址不能大于255个字符！")
	Call FKFun.ShowString(JsCdnUrl,0,255,0,"请输入JS资源CDN地址！","JS资源CDN地址不能大于255个字符！")
	Call FKFun.ShowString(CssCdnUrl,0,255,0,"请输入Css资源CDN地址！","Css资源CDN地址不能大于255个字符！")
	Call FKFun.ShowString(FileCdnUrl,0,255,0,"请输入文件资源CDN地址！","文件资源CDN地址不能大于255个字符！")
	Call FKFun.ShowString(SiteTemplate,1,50,0,"兄弟啊，你怎么能把模板删完啊！","模板文件夹名称不能大于50个字符！")
	Call FKFun.ShowNum(SiteHtml,"请选择是否生成HTML！")
	Call FKFun.ShowNum(SiteOpen,"请选择站点是否开放！")
	Call FKFun.ShowNum(PageSizes,"请选择每页条数！")
	Call FKFun.ShowNum(SiteToPinyin,"请选是否自动生成拼音文件名！")
	Call FKFun.ShowNum(SiteQQ,"请选择是否开启客服浮窗！")
	Call FKFun.ShowNum(SiteNoTrash,"请选择是否开启反垃圾留言功能！")
	Call FKFun.ShowNum(SiteMini,"读取缩略字符数必须是数字！")
	Call FKFun.ShowNum(SiteDelWord,"请选择是否进行非法字符过滤！")
	Call FKFun.ShowNum(SiteTest,"请选择是否开启模板调试模式！")
	Call FKFun.ShowNum(SiteFlash,"请选择是否开启动态地址带Index.asp！")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",7,"SiteName="""&SiteName&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",8,"SiteUrl="""&SiteUrl&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",9,"SiteKeyword="""&SiteKeyword&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",10,"SiteDescription="""&SiteDescription&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",11,"SiteOpen="&SiteOpen&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",12,"SiteTemplate="""&SiteTemplate&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",13,"SiteHtml="&SiteHtml&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",14,"PageSizes="&PageSizes&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",15,"SiteToPinyin="&SiteToPinyin&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",16,"FetionNum="""&FetionNum&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",17,"FetionPass="""&FetionPass&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",18,"SiteQQ="&SiteQQ&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",19,"SiteNoTrash="&SiteNoTrash&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",20,"SiteMini="&SiteMini&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",21,"SiteDelWord="&SiteDelWord&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",22,"SiteTest="&SiteTest&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",23,"SiteFlash="&SiteFlash&"")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",24,"Tel="""&Tel&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",25,"Tel400="""&Tel400&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",26,"Fax="""&Fax&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",27,"Email="""&Email&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",28,"Lianxiren="""&Lianxiren&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",29,"Beian="""&Beian&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",30,"Add="""&Add&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",31,"Tjid="""&Tjid&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",32,"Kfid="""&Kfid&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",33,"SiteLogo="""&SiteLogo&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",34,"Sitepic1="""&Sitepic1&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",35,"Sitepic2="""&Sitepic2&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",36,"Sitepic3="""&Sitepic3&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",37,"Sitepic4="""&Sitepic4&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",38,"Sitepic5="""&Sitepic5&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",39,"Sitepicurl1="""&Sitepicurl1&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",40,"Sitepicurl2="""&Sitepicurl2&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",41,"Sitepicurl3="""&Sitepicurl3&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",42,"Sitepicurl4="""&Sitepicurl4&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",43,"Sitepicurl5="""&Sitepicurl5&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",44,"Sitepictext1="""&Sitepictext1&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",45,"Sitepictext2="""&Sitepictext2&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",46,"Sitepictext3="""&Sitepictext3&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",47,"Sitepictext4="""&Sitepictext4&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",48,"Sitepictext5="""&Sitepictext5&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",49,"Bianjiqi="""&Bianjiqi&"""")
	if len(site301)>0 and len(mainDomain)>0 then
	Call FKFso.FsoLineWrite("../Inc/Site.asp",54,"Site301="""&mainDomain&"|"&Site301&"""")
	else
	Call FKFso.FsoLineWrite("../Inc/Site.asp",54,"Site301=""""")
	end if
	Call FKFso.FsoLineWrite("../Inc/Site.asp",55,"SiteSeoTitle="""&SiteSeoTitle&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",60,"ImgCdnUrl="""&ImgCdnUrl&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",61,"JsCdnUrl="""&JsCdnUrl&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",62,"CssCdnUrl="""&CssCdnUrl&"""")
	Call FKFso.FsoLineWrite("../Inc/Site.asp",63,"FileCdnUrl="""&FileCdnUrl&"""")


	If OldSiteTemplate<>SiteTemplate Then
		Application.Lock()
		Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
		Set F=Fso.GetFolder(Server.MapPath("../Skin/"&SiteTemplate))
		Set ObjFiles=F.Files
		For Each ObjFile In ObjFiles
			If LCase(Split(ObjFile.Name,".")(UBound(Split(ObjFile.Name,"."))))="html" Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Name='"&Replace(LCase(ObjFile.Name),".html","")&"'"
				Rs.Open Sqlstr,Conn,1,3
				If Not Rs.Eof Then
					Rs("Fk_Template_Content")=FKFso.FsoFileRead("../Skin/"&SiteTemplate&"/"&ObjFile.Name)
					Rs.Update()
				Else
					Rs.AddNew()
					Rs("Fk_Template_Name")=Replace(LCase(ObjFile.Name),".html","")
					Rs("Fk_Template_Content")=FKFso.FsoFileRead("../Skin/"&SiteTemplate&"/"&ObjFile.Name)
					Rs.Update()
				End If
				Rs.Close
			End If
		Next
		Set ObjFiles=Nothing
		Set F=Nothing
		Set Fso=Nothing
		Application.UnLock()
	End If
	Response.Write("修改成功！")
	'插入日志
	on error resume next
	dim log_content,log_ip,log_user
	log_content="修改基础信息"
	log_user=Request.Cookies("FkAdminName")
		
	log_ip=FKFun.getIP()
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
End Sub

'==========================================
'函 数 名TestFetion()
'作    用飞信接口测试
'参    数
'==========================================
Sub TestFetion()
	If FetionNum<>"" And FetionPass<>"" Then
		Temp=FKFun.SmsGo("测试飞信接口，如果您收到本短信则飞信接口已经正常运作！")
		Response.Write(Temp)
	Else
		Response.Write("请先设置飞信号！")
	End If
End Sub
%><!--#Include File="../Code.asp"-->