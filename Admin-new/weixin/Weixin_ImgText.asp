<!--#Include File="CheckToken.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<%
'==========================================
'文 件 名：weixin_ImgText.asp
'文件用途：微信图文管理拉取页面
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
Dim Fk_imgText_Title,Fk_imgText_Pic,Fk_imgText_Summary,Fk_imgText_status,Fk_imgText_px,Fk_imgText_url,Fk_imgText_Content,Fk_imgText_Id_List

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WeixinImgTextList() '微信图文列表
	Case 2
		Call WeixinImgTextAdd() '添加微信图文
	Case 3
		Call WeixinImgTextAddDo() '添加微信图文
	Case 4
		Call WeixinImgTextEditForm() '修改微信图文
	Case 5
		Call WeixinImgTextEditDo() '执行修改微信图文
	Case 6
		Call WeixinImgTextDelDo() '执行删除微信图文
	Case 7
		Call WeixinImgTextPx() '执行批量排序
	Case 8
		Call WeixinImgTextOpen() '执行批量启用
	Case 9
		Call WeixinImgTextClose() '执行批量禁用
	Case 10
		Call WeixinImgTextYulan()  '微信图文预览
	Case Else
		Response.Write("没有找到此功能项！")
End Select

sub WeixinImgTextOpen()	
	Id=Trim(Request("Id"))
	if id<>"" then
		if instr(id,",")>0 then
			dim arr,arrpx
			arr=split(id,",")
			for i=0 to ubound(arr)			
				conn.execute("update [weixin_imageText] set imgText_status=0 where id="&arr(i))
			next
		else
			conn.execute("update [weixin_imageText] set imgText_status=0 where id="&Id)
		end if	
		Response.Write("批量启用成功！")
	end if
end sub


sub WeixinImgTextClose()	
	Id=Trim(Request("Id"))
	if id<>"" then
		if instr(id,",")>0 then
			dim arr,arrpx
			arr=split(id,",")
			for i=0 to ubound(arr)			
				conn.execute("update [weixin_imageText] set imgText_status=1 where id="&arr(i))
			next
		else
			conn.execute("update [weixin_imageText] set imgText_status=1 where id="&Id)
		end if	
		Response.Write("批量禁用成功！")
	end if
end sub

sub WeixinImgTextPx()	
	dim px
	Id=Trim(Request("Id"))
	px=Trim(Request("px"))
	if id<>"" then
		if instr(id,",")>0 then
			dim arr,arrpx
			arr=split(id,",")
			arrpx=split(px,",")
			for i=0 to ubound(arr)			
				conn.execute("update [weixin_imageText] set imgText_px="&arrpx(i)&" where id="&arr(i))
			next
		else
			conn.execute("update [weixin_imageText] set imgText_px="&px&" where id="&Id)
		end if	
		Response.Write("批量排序成功！")
	end if
end sub

'==========================================
'函 数 名：WeixinImgTextList()
'作    用：微信图文列表
'参    数：
'==========================================
Sub WeixinImgTextList()
Session("NowPage")=FkFun.GetNowUrl()
%>


<style type="text/css">
.layui-layer-btn{background-color:#fff}
.layui-layer input.Input, input.Input.zffs{margin-left:0px;}
</style>
            <div class="pageright xgzs ssyqtj" style="border-top:0;">
            	<div class="xgzstop" style="border:1px solid #ceecff; border-bottom:0;">
                	<div class="xgzstop_top gnsztopbtn">
                        <a class="no4" href="javascript:void(0);" onclick="ShowBox('/admin-new/weixin/weixin_ImgText.asp?Type=2','添加图文','1000px','450px');return false;">添加</a>
                    </div>
                </div>
            	<div class="xgzsbtm" style="padding-top:0;">
                    <form name="DelList" id="DelList" method="post" action="Down.asp?Type=7" onsubmit="return false;">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" class="khdtj">
                    	<tr>
                    	<th align="center" class="ListTdTop" width="80">选</th>
                        <th align="center" class="ListTdTop" width="80">编号</th>
                        <th align="center" class="ListTdTop" width="160">素材</th>
                        <th align="center" class="ListTdTop" width="160">标题</th>
                        <th align="center" class="ListTdTop" width="80">排序</th>
                        <th align="center" class="ListTdTop" width="200">时间</th>
                        <th align="center" class="ListTdTop" width="80">状态</th>
                        <th align="center" class="ListTdTop">操作</th>
                        </tr>
                        <%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_imageText] Order By imgText_Px Desc,id desc"
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
            <td align="center"><%=Rs("id")%></td>
			<td align="center"><%if Rs("imgText_Pic")>"" then response.write "<a href="""&Rs("imgText_Pic")&""" target=""_blank""><img class=""teebox"" width=""45"" bimg="""&Rs("imgText_Pic")&""" src="""&Rs("imgText_Pic")&""" title="""&Rs("imgText_Title")&""" /></a>"%></td>
            <td align="center"><%=Rs("imgText_Title")%></td>
            <td align="center"><input style="width:30px" type="text" value="<%=Rs("imgText_px")%>" class="Input" name="px" size=2 style="text-align:center"/></td>
            <td height="20" align="center"><%=Rs("imgText_addtime")%></td>
            <td align="center"><%if Rs("imgText_status")=0 then:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_1.gif' title='启用'>":else:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif' title='禁用'>":end if%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('/admin-new/weixin/weixin_ImgText.asp?Type=4&Id=<%=Rs("id")%>','修改图文','1000px','450px');return false;"><img src="/admin-new/images/edit.png" title="编辑"></a> &nbsp;<a href="javascript:void(0);" onclick="ShowBox('/admin-new/weixin/weixin_ImgText.asp?Type=10&Id=<%=Rs("id")%>','预览');return false;"  title="预览"><img src="/admin-new/images/yulan.png"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>        <tr>
            <td height="30" colspan="8">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style=' margin-left:34px;  margin-right:2px; vertical-align:middle'> 全选
            <input type="submit" value="排序" class="Button" onClick="if($('input.Checks:checked').length<1){layer.msg('请先选择要批量操作的数据！');return false};Sends('DelList','/admin-new/weixin/weixin_ImgText.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle;  margin-right:5px; margin-left:10px;'>
			<input type="submit" value="启用" class="Button" onClick="if($('input.Checks:checked').length<1){layer.msg('请先选择要批量操作的数据！');return false};Sends('DelList','/admin-new/weixin/weixin_ImgText.asp?Type=8',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle; margin-right:5px;'>
			<input type="submit" value="禁用" class="Button" onClick="if($('input.Checks:checked').length<1){layer.msg('请先选择要批量操作的数据！');return false};Sends('DelList','/admin-new/weixin/weixin_ImgText.asp?Type=9',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle; margin-right:5px;'>
			<input type="submit" value="删除" class="Button" onClick="if($('input.Checks:checked').length<1){layer.msg('请先选择要批量操作的数据！');return false};Sends('DelList','/admin-new/weixin/weixin_ImgText.asp?Type=6',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle; margin-right:5px;'></td>
        </tr>
		 <tr>
            <td height="30" colspan="8" align="center"><%Call FKFun.ShowPageCode("/admin-new/weixin/weixin_ImgText.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="8" align="center">暂无图文</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
    </form>
                </div>
            </div>
<%
End Sub

'==========================================
'函 数 名：WeixinImgTextAdd()
'作    用：添加微信图文
'参    数：
'==========================================
Sub WeixinImgTextAdd()
%>

<style type="text/css">
	.table_form .explain-col {
		margin: 5px 0;
		min-height: 50px;
	}
	.explain-col ul li {
		margin: 5px 0;
		padding-bottom: 5px;
		border-bottom: 1px dotted #D6D6D6;
	}
	.explain-col ul li .item {
		width: 350px;
		height: 35px;
		line-height: 35px;
		/*
		background: url('http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif') top right no-repeat;
		cursor: pointer;
		*/
		padding-left: 10px;
	}
	
	.news_pic {
		max-width: 200px;
	}
	.rndBtn.plus {
background-position: -1100px 0;
}
.rndBtn {
height: 30px;
width: 30px;
display: inline-block;
background: url("http://image001.dgcloud01.qebang.cn/website/weixin/userMenuButtons.png") no-repeat scroll 0 0 transparent;
}
	.rndBtn.plus:hover {
background-position: -1100px -50px;
}
.fr {
float: right;
}
	.explain-col {
margin: 5px 0;
min-height: 50px;
}
	.explain-col {
border: 1px solid #ffbe7a;
zoom: 1;
background: #fffced;
padding: 8px 10px;
line-height: 20px;
}
.alert-col {
color: #999;
}
.blue, .blue a {
color: #004499;
}

.rndBtn.ext.on:hover {
background-position: -650px -50px;
}
.rndBtn.ext.on {
background-position: -650px 0;
}
.rndBtn.ext:hover {
background-position: -450px -50px;
}
.rndBtn.blkFrd {
background-position: -500px 0;
}
.rndBtn.blkFrd:hover {
background-position: -500px -50px;
}
table {
border-collapse: collapse;
border-spacing: 0;
}
.layui-layer input.Input, input.Input.zffs{margin-left:0px;}
</style>
<script language="javascript">	
	var id = 0;
	/**
	 * 添加图文
	 * @param	string	type
	 * @param	integer	id
	 * @return
	 */
	function add_news() {
	}

	
    $(document).ready(function() {

		// 选择素材
		$('.icon_ui_btn').click('click', function() {
			var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
			
			$.get('/admin-new/weixin/weixin_getSucaiList.asp?type=1&id=0',
				function(data){
					layer.close(layer1);
					//document.getElementById('BoxContent').innerHTML=data;
					var layopen=layer.open({
						type: 1,
						title: '选择素材',
						shadeClose: true,
						btn:['确定','取消'],
						yes: function(index, layero){ //或者使用btn1
							if(layero.find("input.Checks:checked").length>0){
								var c;
								c=layero.find("input.Checks:checked").next("#picurl").val();
								$("#Fk_imgText_Pic").val(c);
								var url=c;
								var thisDIV = $("#Fk_imgText_Pic");
								if(url) {			
									if(thisDIV.prev('p').length < 1) {
										var html = '<p><a href="' + url + '" target="_blank" title="点击查看原图"><img class="news_pic" src="' + url + '" /></a><br /><br /></p>';
										thisDIV.before(html);
									}else if(url != $(this).prev('p').find('img').attr('src')) {
										thisDIV.prev('p').find('img').attr('src', url);
									}	
								}else{
									$(this).prev('p').remove();
								}
								layer.close(layopen);
							}
						},
						shade: 0.5,
						area: ['650px', 'auto'],
						zIndex:88888,
						content: data
					});
				}
			)
			return false;
		});
		
		$(".selectNewsUrl").click(function(){
			var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
			
			$.get('/admin-new/weixin/Weixin_GetArticle.asp?type=1&id=0',
				function(data){
					layer.close(layer1);
					//document.getElementById('BoxContent').innerHTML=data;
					var layopen=layer.open({
						type: 1,
						title: '选择图文',
						shadeClose: true,
						btn:['确定','取消'],
						yes: function(index, layero){ //或者使用btn1
							if(layero.find("input.Checks:checked").length>0){
								var c;
								c=layero.find("input.Checks:checked").next(".hid").val();
								$("#Fk_imgText_url").val(c);
								
								layer.close(layopen);
							}
						},
						shade: 0.5,
						area: ['700px', 'auto'],
						zIndex:88888,
						content: data
					});
				}
			)
			return false;
				
		});
		
		/*/ 更新封面
		$('#Fk_imgText_Pic').blur(function() {
			var url = $(this).val();
			if(url) {			
				if($(this).prev('p').length < 1) {
					var html = '<p><a href="' + url + '" target="_blank" title="点击查看原图"><img class="news_pic" src="' + url + '" /></a><br /><br /></p>';
					$(this).before(html);
				}else if(url != $(this).prev('p').find('img').attr('src')) {
					$(this).prev('p').find('img').attr('src', url);
				}	
			}else{
				$(this).prev('p').remove();
			}
		});
*/

		// 移出图文
		$('.item > .blkFrd').live('click', function() {
			$(this).parent().parent().remove();
		});
		// 下移图文
		$('.item > .ext.on').live('click', function() {
			var parent = $(this).parent().parent();
			if(parent.next('li').length > 0) {
				parent.before(parent.next('li'));
			}
		});
		
		// 多图文
		$('.items_expanded .plus').click(function() {
			var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
			
			$.get('/admin-new/weixin/weixin_getNewsList.asp?type=1&id='+id,
				function(data){
					layer.close(layer1);
					//document.getElementById('BoxContent').innerHTML=data;
					var layopen=layer.open({
						type: 1,
						title: '选择图文',
						shadeClose: true,
						btn:['确定','关闭'],
						yes: function(index, layero){ //或者使用btn1
							if(layero.find("input.Checks:checked").length>0){
								var c,box;
								c=layero.find("input.Checks:checked");
								box = $('.items_expanded > ul');
								for(i = 0; i < c.length; i++) {
									if(c[i].type == 'checkbox' && c[i].name == 'ListId' && c[i].checked) {
										if(box.children().length > 9) {
											layer.msg('图文数量已超出');
											break;
										}
										id = c[i].value;
										if(id) {
											val = layero.find('#news_' + id).val();
											if(id && val && box.find(".item[nid='" + id + "']").length < 1) {
												html = '<li><div class="item" nid="' + id + '"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a>' + val + '</div></li>';
												box.append(html);
											}
										}
									}
								}
								layer.close(layopen);
							}
						},
						shade: 0.5,
						area: ['750px', 'auto'],
						zIndex:88888,
						content: data
					});
				}
			)
			return false;
				
		});
		

    });
	
	function updateItems(){
		if($(".items_expanded > ul > li").length>0){
			$(".items_expanded > ul > li").each(function(i){
				if(i==0){
					items=$(this).children(".item").attr("nid");
				}
				else{
					items=items+","+$(this).children(".item").attr("nid");
				}
				$("#items").val(items);
			})
		}
		else{
			$("#items").val("");
		}
	
	}
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin-new/weixin/weixin_ImgText.asp?Type=3" onsubmit="updateItems();return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right" width="100">标题：</td>
	        <td><input name="Fk_imgText_Title" type="text" class="Input" id="Fk_imgText_Title" size="60" /></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">摘要：</td>
	        <td><textarea name="Fk_imgText_Summary" id="Fk_imgText_Summary" class="Input" style="height:70px;width:740px;"></textarea></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">正文：</td>
	        <td><textarea name="Fk_imgText_Content" class="<%=bianjiqi%>" id="Fk_imgText_Content" rows="8" style="height:70px; width:740px;"></textarea></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">图文封面：</td>
	        <td><input name="Fk_imgText_Pic" class="Input" type="text" id="Fk_imgText_Pic" size="60"> &nbsp; <a href="javascript: void(0);" class="icon_ui_btn blue" for="Fk_imgText_Pic" ui_type="1" ui_tpl="0" title="选择素材">选择素材</a><br><span class="alert-col">大图片建议尺寸: 360px*200px, 文件小于200k; 推荐上传到腾讯微博再获取外链</span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">图文外链：</td>
	        <td><input name="Fk_imgText_url" class="Input" type="text" id="Fk_imgText_url" size="60"> &nbsp; <a href="javascript: void(0);" class="selectNewsUrl blue" for="Fk_imgText_url" ui_type="1" ui_tpl="0" title="选择链接">选择链接</a><br><span class="alert-col">系统会自动生成，如需要跳转到外链请填写</span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">多图文：</td>
	        <td><div class="explain-col items_expanded"> 
				<h3>
					<a href="javascript: void(0);" class="rndBtn plus fr" title="添加图文"></a>
					<input type="hidden" name="items" id="items" value="">
				</h3>
				<ul>
								</ul>
			</div></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">排序：</td>
	        <td><input name="Fk_imgText_px" class="Input" type="text" id="Fk_imgText_px" value="0"></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">状态：</td>
	        <td><input name="Fk_imgText_status" type="radio" id="Fk_imgText_status" value="0" checked="checked" />启用&nbsp;&nbsp;
            <input type="radio" name="Fk_imgText_status" id="Fk_imgText_status1" value="1" />禁用</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left;" class="tcbtm">
        <input style="margin-left:113px;" type="submit" onclick="Sends('WeixinAdd','/admin-new/weixin/weixin_ImgText.asp?Type=3',0,'',0,1,'MainRight','/admin-new/weixin/weixin_ImgText.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>

<%
End Sub

'==============================
'函 数 名：WeixinImgTextAddDo
'作    用：执行添加微信图文
'参    数：
'==============================
Sub WeixinImgTextAddDo()
	Fk_imgText_Title	= FKFun.HTMLEncode(Trim(Request.Form("Fk_imgText_Title")))
	Fk_imgText_Summary	= Trim(Request.Form("Fk_imgText_Summary"))
	Fk_imgText_url		= FKFun.HTMLEncode(Trim(Request.Form("Fk_imgText_url")))
	Fk_imgText_Pic		= FKFun.HTMLEncode(Trim(Request.Form("Fk_imgText_Pic")))
	Fk_imgText_px		= Trim(Request.Form("Fk_imgText_px"))
	Fk_imgText_status	= Trim(Request.Form("Fk_imgText_status"))
	Fk_imgText_Id_List	= Trim(Request.Form("items"))
	Fk_imgText_Content	= Trim(Request.Form("Fk_imgText_Content"))
	Call FKFun.ShowString(Fk_imgText_Title,1,100,0,"请输入标题名称！","标题名称不能大于100个字节！")
	Sqlstr="Select * From [weixin_imageText]"
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs.AddNew()
		Rs("imgText_Title")		=Fk_imgText_Title
		Rs("imgText_Summary")	=Fk_imgText_Summary
		Rs("imgText_url")		=Fk_imgText_url
		Rs("imgText_Pic")		=Fk_imgText_Pic
		Rs("imgText_px")		=Fk_imgText_px
		Rs("imgText_status")	=Fk_imgText_status
		Rs("imgText_Id_List")	=Fk_imgText_Id_List
		Rs("imgText_Content")	=Fk_imgText_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("图文添加成功！")
	Rs.Close
End Sub

function getInfo(id)
	getInfo=""
	Sqlstr="Select imgText_Title From [weixin_imageText] Where id=" & Id
	set rs=conn.execute(sqlstr)
	if not rs.eof then
	getInfo=rs("imgText_Title")
	end if
	rs.close
end function

'==========================================
'函 数 名：WeixinImgTextEditForm
'作    用：修改微信图文
'参    数：
'==========================================
Sub WeixinImgTextEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [weixin_imageText] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_imgText_Title	= Rs("imgText_Title")
		Fk_imgText_Summary	= Rs("imgText_Summary")
		Fk_imgText_url		= Rs("imgText_url")
		Fk_imgText_Pic		= Rs("imgText_Pic")
		Fk_imgText_px		= Rs("imgText_px")
		Fk_imgText_status	= Rs("imgText_status")
		Fk_imgText_Id_List	= Rs("imgText_Id_List")
		Fk_imgText_Content	= Rs("imgText_Content")
	End If
	Rs.Close
%>
<style type="text/css">
	.table_form .explain-col {
		margin: 5px 0;
		min-height: 50px;
	}
	.explain-col ul li {
		margin: 5px 0;
		padding-bottom: 5px;
		border-bottom: 1px dotted #D6D6D6;
	}
	.explain-col ul li .item {
		width: 350px;
		height: 35px;
		line-height: 35px;
		/*
		background: url('http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif') top right no-repeat;
		cursor: pointer;
		*/
		padding-left: 10px;
	}
	
	.news_pic {
		max-width: 200px;
	}
	.rndBtn.plus {
background-position: -1100px 0;
}
.rndBtn {
height: 30px;
width: 30px;
display: inline-block;
background: url("http://image001.dgcloud01.qebang.cn/website/weixin/userMenuButtons.png") no-repeat scroll 0 0 transparent;
}
	.rndBtn.plus:hover {
background-position: -1100px -50px;
}
.fr {
float: right;
}
	.explain-col {
margin: 5px 0;
min-height: 50px;
}
	.explain-col {
border: 1px solid #ffbe7a;
zoom: 1;
background: #fffced;
padding: 8px 10px;
line-height: 20px;
}
.alert-col {
color: #999;
}
.blue, .blue a {
color: #004499;
}

.rndBtn.ext.on:hover {
background-position: -650px -50px;
}
.rndBtn.ext.on {
background-position: -650px 0;
}
.rndBtn.ext:hover {
background-position: -450px -50px;
}
.rndBtn.blkFrd {
background-position: -500px 0;
}
.rndBtn.blkFrd:hover {
background-position: -500px -50px;
}
table {
border-collapse: collapse;
border-spacing: 0;
}
</style>
<script language="javascript">	
	var id = 0;
	/**
	 * 添加图文
	 * @param	string	type
	 * @param	integer	id
	 * @return
	 */
	function add_news() {
	}

	
    $(document).ready(function() {

		// 选择素材
		$('.icon_ui_btn').click('click', function() {
			var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
			
			$.get('/admin-new/weixin/weixin_getSucaiList.asp?type=1&id=0',
				function(data){
					layer.close(layer1);
					//document.getElementById('BoxContent').innerHTML=data;
					var layopen=layer.open({
						type: 1,
						title: '选择素材',
						shadeClose: true,
						btn:['确定','取消'],
						yes: function(index, layero){ //或者使用btn1
							if(layero.find("input.Checks:checked").length>0){
								var c;
								c=layero.find("input.Checks:checked").next("#picurl").val();
								$("#Fk_imgText_Pic").val(c);
								var url=c;
								var thisDIV = $("#Fk_imgText_Pic");
								if(url) {			
									if(thisDIV.prev('p').length < 1) {
										var html = '<p><a href="' + url + '" target="_blank" title="点击查看原图"><img class="news_pic" src="' + url + '" /></a><br /><br /></p>';
										thisDIV.before(html);
									}else if(url != $(this).prev('p').find('img').attr('src')) {
										thisDIV.prev('p').find('img').attr('src', url);
									}	
								}else{
									$(this).prev('p').remove();
								}
								layer.close(layopen);
							}
						},
						shade: 0.5,
						area: ['650px', 'auto'],
						zIndex:88888,
						content: data
					});
				}
			)
			return false;
		});
		
		$(".selectNewsUrl").click(function(){
			var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
			
			$.get('/admin-new/weixin/Weixin_GetArticle.asp?type=1&id=0',
				function(data){
					layer.close(layer1);
					//document.getElementById('BoxContent').innerHTML=data;
					var layopen=layer.open({
						type: 1,
						title: '选择图文',
						shadeClose: true,
						btn:['确定','取消'],
						yes: function(index, layero){ //或者使用btn1
							if(layero.find("input.Checks:checked").length>0){
								var c;
								c=layero.find("input.Checks:checked").next(".hid").val();
								$("#Fk_imgText_url").val(c);
								
								layer.close(layopen);
							}
						},
						shade: 0.5,
						area: ['700px', 'auto'],
						zIndex:88888,
						content: data
					});
				}
			)
			return false;
				
		});


		// 移出图文
		$('.item > .blkFrd').live('click', function() {
			$(this).parent().parent().remove();
		});
		// 下移图文
		$('.item > .ext.on').live('click', function() {
			var parent = $(this).parent().parent();
			if(parent.next('li').length > 0) {
				parent.before(parent.next('li'));
			}
		});
		
		// 多图文
		$('.items_expanded .plus').click(function() {
		
			var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
			
			$.get('/admin-new/weixin/weixin_getNewsList.asp?type=1&id=<%=id%>',
				function(data){
					layer.close(layer1);
					//document.getElementById('BoxContent').innerHTML=data;
					var layopen=layer.open({
						type: 1,
						title: '选择图文',
						shadeClose: true,
						btn:['确定','关闭'],
						yes: function(index, layero){ //或者使用btn1
							if(layero.find("input.Checks:checked").length>0){
								var c,box;
								c=layero.find("input.Checks:checked");
								box = $('.items_expanded > ul');
								for(i = 0; i < c.length; i++) {
									if(c[i].type == 'checkbox' && c[i].name == 'ListId' && c[i].checked) {
										if(box.children().length > 9) {
											layer.msg('图文数量已超出');
											break;
										}
										id = c[i].value;
										if(id) {
											val = layero.find('#news_' + id).val();
											if(id && val && box.find(".item[nid='" + id + "']").length < 1) {
												html = '<li><div class="item" nid="' + id + '"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a>' + val + '</div></li>';
												box.append(html);
											}
										}
									}
								}
								layer.close(layopen);
							}
						},
						shade: 0.5,
						area: ['650px', 'auto'],
						zIndex:88888,
						content: data
					});
				}
			)
			return false;
			
				
		});
		
		

    });
	
	
	
	function updateItems(){
		if($(".items_expanded > ul > li").length>0){
			$(".items_expanded > ul > li").each(function(i){
				if(i==0){
					items=$(this).children(".item").attr("nid");
				}
				else{

					items=items+","+$(this).children(".item").attr("nid");
				}
				$("#items").val(items);
			})
		}
		else{
			$("#items").val("");
		}
	
	}
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin-new/weixin/weixin_ImgText.asp?Type=3" onsubmit="updateItems();return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">标题：</td>
	        <td><input name="Fk_imgText_Title" type="text" class="Input" id="Fk_imgText_Title" size="40" value="<%=Fk_imgText_Title%>"/><input type="hidden" value="<%=id%>" name="id"/></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">摘要：</td>
	        <td><textarea name="Fk_imgText_Summary" id="Fk_imgText_Summary" class="Input" style="height:70px;width:740px;"><%=Fk_imgText_Summary%></textarea></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">正文：</td>
	        <td><textarea name="Fk_imgText_Content" class="<%=bianjiqi%>" id="Fk_imgText_Content" rows="8" style="height:70px; width:740px;"><%=Fk_imgText_Content%></textarea></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">图文封面：</td>
	        <td><input name="Fk_imgText_Pic" class="Input" type="text" id="Fk_imgText_Pic" size="60" value="<%=Fk_imgText_Pic%>"/> &nbsp; <a href="javascript: void(0);" class="icon_ui_btn blue" for="Fk_imgText_Pic" ui_type="1" ui_tpl="0" title="选择素材">选择素材</a><br><span class="alert-col">大图片建议尺寸: 360px*200px, 文件小于200k; 推荐上传到腾讯微博再获取外链</span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">图文外链：</td>
	        <td><input name="Fk_imgText_url" class="Input" type="text" id="Fk_imgText_url" size="60" value="<%=Fk_imgText_url%>" /> &nbsp; <a href="javascript: void(0);" class="selectNewsUrl blue" for="Fk_imgText_url" ui_type="1" ui_tpl="0" title="选择链接">选择链接</a><br><span class="alert-col" >系统会自动生成，如需要跳转到外链请填写</span></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">多图文：</td>
	        <td><div class="explain-col items_expanded"> 
				<h3>
					<a href="javascript: void(0);" class="rndBtn plus fr" title="添加图文"></a>
					<input  type="hidden" name="items" id="items"  value="<%=Fk_imgText_Id_List%>"/>
				</h3>
				<ul>
				<%if Fk_imgText_Id_List<>"" then
					if instr(Fk_imgText_Id_List,",")>0 then
					dim arr
					arr=split(Fk_imgText_Id_List,",")
					for i=0 to ubound(arr)%>
					<li><div class="item" nid="<%=arr(i)%>"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a><%=getInfo(arr(i))%><div></li>
				<%next
				else%>
				<li><div class="item" nid="<%=Fk_imgText_Id_List%>"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a><%=Fk_imgText_Id_List%><div></li>
				<%end if
				end if%>
				</ul>
			</div></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">排序：</td>
	        <td><input name="Fk_imgText_px" class="Input" type="text" id="Fk_imgText_px"  value="<%=Fk_imgText_px%>"></td>
	        </tr>
	    <tr>
	        <td height="25" align="right">状态：</td>
	        <td><input name="Fk_imgText_status" type="radio" id="Fk_imgText_status" value="0" checked="checked" <%if Fk_imgText_status=0 then response.write "checked"%>/>启用&nbsp;&nbsp;
            <input type="radio" name="Fk_imgText_status" id="Fk_imgText_status1" value="1" />禁用</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
        <input type="submit" style="margin-left:113px" onclick="Sends('WeixinAdd','/admin-new/weixin/weixin_ImgText.asp?Type=5',0,'',0,1,'MainRight','/admin-new/weixin/weixin_ImgText.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinImgTextEditDo
'作    用：执行修改图文
'参    数：
'==============================
Sub WeixinImgTextEditDo()
	Id					= Trim(Request.Form("Id"))
	Fk_imgText_Title	= FKFun.HTMLEncode(Trim(Request.Form("Fk_imgText_Title")))
	Fk_imgText_Summary	= Trim(Request.Form("Fk_imgText_Summary"))
	Fk_imgText_url		= FKFun.HTMLEncode(Trim(Request.Form("Fk_imgText_url")))
	Fk_imgText_Pic		= FKFun.HTMLEncode(Trim(Request.Form("Fk_imgText_Pic")))
	Fk_imgText_px		= Trim(Request.Form("Fk_imgText_px"))
	Fk_imgText_status	= Trim(Request.Form("Fk_imgText_status"))
	Fk_imgText_Id_List	= Trim(Request.Form("items"))
	Fk_imgText_Content	= Trim(Request.Form("Fk_imgText_Content"))
	Call FKFun.ShowString(Fk_imgText_Title,1,100,0,"请输入标题名称！","标题名称不能大于100个字节！")
	Sqlstr="Select * From [weixin_imageText] where id="&id
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs("imgText_Title")		=Fk_imgText_Title
		Rs("imgText_Summary")	=Fk_imgText_Summary
		Rs("imgText_url")		=Fk_imgText_url
		Rs("imgText_Pic")		=Fk_imgText_Pic
		Rs("imgText_px")		=Fk_imgText_px
		Rs("imgText_status")	=Fk_imgText_status
		Rs("imgText_Id_List")	=Fk_imgText_Id_List
		Rs("imgText_Content")	=Fk_imgText_Content
		Rs.Update()
		Application.UnLock()
		Response.Write("图文修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：WeixinImgTextDelDo
'作    用：执行删除微信图文
'参    数：
'==============================
Sub WeixinImgTextDelDo()
	Id=Trim(Request("Id"))
	Sqlstr="Select * From [weixin_imageText] Where id in("& Id &")"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("微信图文删除成功！")
	Else
		Response.Write("微信图文不存在！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：WeixinImgTextYulan
'作    用：微信图文预览
'参    数：
'==========================================
Sub WeixinImgTextYulan()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [weixin_imageText] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_imgText_Title	= Rs("imgText_Title")
		Fk_imgText_Summary	= Rs("imgText_Summary")
		Fk_imgText_url		= Rs("imgText_url")
		Fk_imgText_Pic		= Rs("imgText_Pic")
		Fk_imgText_px		= Rs("imgText_px")
		Fk_imgText_status	= Rs("imgText_status")
		Fk_imgText_Id_List	= Rs("imgText_Id_List")
		Fk_imgText_Content	= Rs("imgText_Content")
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

	z-index: 1005;
}

.chat_box .chat_wrap {
 margin:0 auto;
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

<div id="BoxContents" style="width:93%; padding-top:20px;">
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
						<img onerror="this.parentNode.removeChild(this)" src="<%=Fk_imgText_pic%>" />
					</div>
										<div class="mediaImgFooter">
						<p class="mesgTitleTitle left"><%=Fk_imgText_Title%></p>
						<div class="clr"></div>
					</div>
				</div>
				</a>
			
					<%if Fk_imgText_Id_List<>"" then
					set rs=conn.execute("select * from weixin_imageText where id in("&Fk_imgText_Id_List&") order by imgText_px desc")
					if not rs.eof then%>
					<div class="mediaContent">
					<%do while not rs.eof%>
					<a href="#">
					<div class="mediaMesg">
						<span class="mediaMesgDot"></span>
						<div class="mediaMesgTitle left">
							<p class="left"><%=rs("imgText_Title")%></p>
							<div class="clr"></div>
						</div>
						<div class="mediaMesgIcon right">
							<img onerror="this.parentNode.removeChild(this)" src="<%=rs("imgText_pic")%>" />
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
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:center" class="tcbtm">
        <input style="margin-left:0;" type="submit" onclick="Sends('WeixinAdd','/admin-new/weixin/weixin_ImgText.asp?Type=5',0,'',0,1,'MainRight','/admin-new/weixin/weixin_ImgText.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
<%
End Sub
%><!--#Include File="../../Code.asp"-->























