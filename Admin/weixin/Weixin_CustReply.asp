<!--#Include File="../AdminCheck.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<%
'==========================================
'文 件 名：Weixin_CustReply.asp
'文件用途：微信回复管理拉取页面
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
Dim Fk_reply_qtitle,Fk_reply_qanswerText,Fk_reply_qanswerNews,Fk_reply_qanswerResource,Fk_reply_type

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WeixinCustReplyList() '微信回复列表
	Case 2
		Call WeixinCustReplyAdd() '添加微信回复
	Case 3
		Call WeixinCustReplyAddDo() '添加微信回复
	Case 4
		Call WeixinCustReplyEditForm() '修改微信自定义回复
	Case 5
		Call WeixinCustReplyEditDo() '执行修改微信自定义回复
	Case 6
		Call WeixinCustReplyDelDo() '执行删除微信自定义回复
	Case 7
		Call WeixinCustReplyMake() '生成微信自定义回复
	Case 8
		Call WeixinCustReplyMakeDo() '执行生成微信自定义回复
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：WeixinCustReplyList()
'作    用：微信回复列表
'参    数：
'==========================================
Sub WeixinCustReplyList()
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
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false;">自定义回复</a></li>
        <li><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/Weixin_CustReply.asp?Type=2');return false;">添加</a></li>
    </ul>
</div>
<div id="ListTop">自定义回复模块：<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle;"/>&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','/admin/weixin/Weixin_CustReply.asp?Type=1&SearchStr='+escape(document.all.SearchStr.value));" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/>&nbsp;&nbsp;请选择模块：
<select name="D1" id="D1" onChange="window.execScript(this.options[this.selectedIndex].value);" style="vertical-align:middle;">
      <option value="alert('请选择模块');">请选择模块</option>
</select>
</div>
<div id="ListContent">
    <form name="DelList" id="DelList" method="post" action="Down.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop">编号</td>
            <td align="center" class="ListTdTop">问题</td>
            <td align="center" class="ListTdTop">时间</td>
            <td align="center" class="ListTdTop">排序</td>
            <td align="center" class="ListTdTop">状态</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Dim Rs
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Weixin_CustReply] where 1=1"
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And reply_qtitle Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By px Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		
		Dim DownTemplate
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
            <td>&nbsp;<%=Rs("id")%></td>
            <td align="center"><%=Rs("reply_qtitle")%><%if Rs("reply_type")=0 then response.write "<br>"&Rs("reply_qanswerText")%></td>
            <td align="center"><%=Rs("add_time")%></td>
            <td height="20" align="center"><%=Rs("px")%></td>
            <td align="center"><%if Rs("Status")=0 then:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_1.gif' title='启用'>":else:response.write "<img src='http://image001.dgcloud01.qebang.cn/website/weixin/status_0.gif' title='禁用'>":end if%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/Weixin_CustReply.asp?Type=4&Id=<%=Rs("id")%>');return false;"><img src="/admin/images/edit.png"></a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("reply_qtitle")%>”，此操作不可逆！','/admin/weixin/Weixin_CustReply.asp?Type=6&Id=<%=Rs("id")%>','MainRight','<%=Session("NowPage")%>');return false;"><img src="/admin/images/del.png"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>
        <tr>
            <td height="30" colspan="8">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style='text-indent:10px;vertical-align:middle'> 全选
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除选中的下载吗？')){Sends('DelList','/admin/weixin/Weixin_CustReply.asp?Type=6',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}" style='vertical-align:middle'></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="8" align="center">暂无数据</td>
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
'函 数 名：WeixinCustReplyAdd()
'作    用：添加微信回复
'参    数：
'==========================================
Sub WeixinCustReplyAdd()
%>
<style type="text/css">

.rndBtn {
height: 30px;
width: 30px;
display: inline-block;
background: url("http://image001.dgcloud01.qebang.cn/website/weixin/userMenuButtons.png") no-repeat scroll 0 0 transparent;
}
.rndBtn.plus {
background-position: -1100px 0;
}
.fr {
float: right;
}
td ul li {
margin: 1px 0;
padding-bottom: 3px;
}
.alert-col {
color: #666;
}

.explain-col {
border: 1px solid #ffbe7a;
zoom: 1;
background: #fffced;
padding: 8px 10px;
line-height: 20px;
}

.hide {
display: none;
}

.reply_text textarea{height:50px;background:none;}

	.emot_box {
		/*
		height: 160px;
		*/
		height: 100%;
		overflow: auto;
	}
		.emot_box .emj {
			float: left;
			width: 102px;
			height: 40px;
			padding-top: 50px;
			background-attachment: scroll;
			background-position: top center;
			background-repeat: no-repeat;
			background-size: 38px 38px;
			text-align: center;
		}
		.explain-col ul li .item {
width: 350px;
height: 35px;
line-height: 35px;
padding-left: 10px;
}
.explain-col ul li {
margin: 5px 0;
padding-bottom: 5px;
border-bottom: 1px dotted #D6D6D6;
}
		.emot_box{width:712px;margin-top:10px;}
		.rndBtn.blkFrd {
background-position: -500px 0;
}
		.rndBtn.blkFrd:hover {
background-position: -500px -50px;
}
.rndBtn.ext.on {
background-position: -650px 0;
}
.rndBtn.ext.on:hover {
background-position: -650px -50px;
}

.table_form .explain-col {
margin: 30px 0 5px 0;
min-height: 50px;
}
</style>
<script type="text/javascript">
	/**
	 * 添加图文
	 * @return
	 */
	function add_news() {
		// 多图文
				
		}


	/**
	 * 添加语音
	 * @return
	 */
	function add_music() {
	
			
	}
	
$(document).ready(function(){
		// 回复类型切换
		$('#reply_type1, #reply_type2, #reply_type3').click(function() {
			var id = $(this).attr('id');
			
			$('.reply_text, .reply_news, .reply_rs').hide();

			if(id == 'reply_type2') {
				// 图文
				$('.reply_news').show();
			}else if(id == 'reply_type3') {
				// 语音
				$('.reply_rs').show();
			}else {
				// 文本
				$('.reply_text').show();
			}
		})
		
				// 插入表情代码
		$('#add_emot').click(function() {
			if($(this).is(':checked')) {
				$('.emot_box').removeClass('hide').show();
			}else {
				$('.emot_box').addClass('hide').hide();
			}
		})
		
		// 优先显示图文/语音回复/文本
		if($('.reply_news > ul > li').length > 0) {
			$('#reply_type2').click();
		}else if($('.reply_rs > ul > li').length > 0) {
			$('#reply_type3').click();
		}else if($('.reply_text > ul > li').length > 0) {
			$('#reply_type1').click();
		}
		
		// 移出
		$('.item > .blkFrd').live('click', function() {
			$(this).parent().parent().remove();
		});
		// 下移
		$('.item > .ext.on').live('click', function() {
			var parent = $(this).parent().parent();
			if(parent.next('li').length > 0) {
				parent.before(parent.next('li'));
			}
		});
		
				// 追加问答
		$('.plus').click(function() {
			var parent, html;
			if($(this).hasClass('answer') && $('#reply_type2').is(':checked')) {
				// 图文
				ymPrompt.win({message:'/admin/weixin/weixin_getNewsList.asp?type=1&id=0',
					width : 600,
					height : 350,
					title:'选择图文',
					btn: [['确定','ok'],['关闭','close']],
					maxBtn : true,
					minBtn : true,
					closeBtn : true,
					iframe : true,handler:function(msg){
						if (msg == 'error') {
						
						}else if(msg == 'ok'){ 
							$("#news").val("");
							var items="";
							if($("iframe").contents().find("input.Checks:checked").length>0){
							
								var html;
								var id, val, box,c;
								box = $('.reply_news > ul');
								c=$("iframe").contents().find("input.Checks:checked");
								for(i = 0; i < c.length; i++) {
									if(c[i].type == 'checkbox' && c[i].name == 'ListId' && c[i].checked) {
										if(box.children().length > 9) {
											alert('图文数量已超出');
											break;
										}
										id = c[i].value;
										if(id) {
											val = $("iframe").contents().find('#news_' + id).val();
											if(id && val && box.find(".item[nid='" + id + "']").length < 1) {
												html = '<li><div class="item" nid="' + id + '"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a>' + val + '</div></li>';
												box.append(html);
											}
										}
									}
								}	
							
								
								
								$("#news").val(items);
							}
						}
					}
				})
				return false;
					
				

			}else if($(this).hasClass('answer') && $('#reply_type3').is(':checked')) {
				// 语音
				ymPrompt.win({message:'/admin/weixin/weixin_getMusicList.asp?type=1&id=0',
					width : 600,
					height : 350,
					title:'选择语音',
					btn: [['确定','ok'],['关闭','close']],
					maxBtn : true,
					minBtn : true,
					closeBtn : true,
					iframe : true,handler:function(msg){
						if (msg == 'error') {
						
						}else if(msg == 'ok'){ 
							$("#resource").val("");
							var items="";
							if($("iframe").contents().find("input.Checks:checked").length>0){
							
								var html;
								var id, val, box,c;
								box = $('.reply_rs > ul');
								c=$("iframe").contents().find("input.Checks:checked");
								for(i = 0; i < c.length; i++) {
									if(c[i].type == 'checkbox' && c[i].name == 'ListId' && c[i].checked) {
										if(box.children().length > 9) {
											alert('语音数量已超出');
											break;
										}
										id = c[i].value;
										if(id) {
											val = $("iframe").contents().find('#news_' + id).val();
											if(id && val && box.find(".item[nid='" + id + "']").length < 1) {
												html = '<li><div class="item" nid="' + id + '"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a>' + val + '</div></li>';
												box.append(html);
											}
										}
									}
								}	
							
								
								$("#resource").val(items);
							}
						}
					}
				})
				return false;
			}else {
				// 文本问答
				parent = $(this).hasClass('answer') ? $(this).next().children('ul'): $(this).next('ul');
				if(parent.children().length >= 5) {
					return;
				}
				html = '<li>' + parent.children('li:first').html() + '</li>'
				parent.append(html).find('li:last > input').val('').removeAttr('readonly');
			}

		})
})


	function updateItems(){
		if($(".reply_news > ul > li").length>0){
			$(".reply_news > ul > li").each(function(i){
				if(i==0){
					items=$(this).children(".item").attr("nid");
				}
				else{
					items=items+","+$(this).children(".item").attr("nid");
				}
				$("#news").val(items);
			})
		}
		else{
			$("#news").val("");
		}
	
		if($(".reply_rs > ul > li").length>0){
			$(".reply_rs > ul > li").each(function(i){
				if(i==0){
					items=$(this).children(".item").attr("nid");
				}
				else{
					items=items+","+$(this).children(".item").attr("nid");
				}
				$("#resource").val(items);
			})
		}
		else{
			$("#resource").val("");
		}
	}
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/Weixin.asp?Type=3" onsubmit="updateItems();return false;">
<div id="BoxTop" style="width:98%;"><span>添加自定义回复</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" class="table_form" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">问题：</td>
	        <td><a href="javascript: void(0);" class="rndBtn plus fr question" title="添加问题"></a>
			<ul>
			<li><input name="question[]" type="text" class="Input" size="60" maxlength="200" value=""/></li>
			</ul>
			<div class="alert-col">点击右边的添加按钮，可设置多个相似问题。 </div></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">回复类型：</td>
	        <td><span class="fr">支持同时编辑多种回复类型，回复优先级: 图文 > 语音 > 文本</span><input name="reply_type" type="radio" id="reply_type1" class="Input" value="0" style="verticle-align:middle"/> <label for="reply_type1" style="verticle-align:middle">文本</label> &nbsp; <input name="reply_type" type="radio" id="reply_type2" class="Input" value="2" style="verticle-align:middle"/> <label for="reply_type2" style="verticle-align:middle">图文</label> &nbsp; <input name="reply_type" type="radio" id="reply_type3" class="Input" value="3" style="verticle-align:middle"/> <label for="reply_type3" style="verticle-align:middle">语音</label> </td>
	        </tr>
	    <tr>
	        <td height="25" align="right">回复内容：</td>
	        <td><a href="javascript: void(0);" class="rndBtn plus fr answer" title="添加回复"></a>
				
				<div class="reply_text" style="display: block;">
				<ul>
									<li><textarea name="answer[]" class="Input" rows="2" cols="56"></textarea></li>
								</ul>
				<div class="alert-col">
					可设置多个备选回复(默认为随机回复，修改<em>机器人_随机回复</em>可关闭随机回复机制) &nbsp; 
					<input type="checkbox" id="add_emot" value="" class="hide"> <label for="add_emot" class="hide">插入表情代码</label> &nbsp; 
								</div>

				<!--textarea name="answer" id="answer" rows="3" cols="58" class="validate[required]"></textarea-->
					
				<div class="emot_box explain-col hide" title="官方表情发送到微信端后会自动转成微信表情">
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/哈哈.png)" code="哈哈" title="哈哈" class="emj">[emot=哈哈/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/得瑟.png)" code="得瑟" title="得瑟" class="emj">[emot=得瑟/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/示好.png)" code="示好" title="示好" class="emj">[emot=示好/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/哼.png)" code="哼" title="哼" class="emj">[emot=哼/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/嚎啕大哭.png)" code="嚎啕大哭" title="嚎啕大哭" class="emj">[emot=嚎啕大哭/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/不懂.png)" code="不懂" title="不懂" class="emj">[emot=不懂/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/俏皮.png)" code="俏皮" title="俏皮" class="emj">[emot=俏皮/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/别惹我.png)" code="别惹我" title="别惹我" class="emj">[emot=别惹我/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/呜呜.png)" code="呜呜" title="呜呜" class="emj">[emot=呜呜/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/打瞌睡.png)" code="打瞌睡" title="打瞌睡" class="emj">[emot=打瞌睡/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/雷倒.png)" code="雷倒" title="雷倒" class="emj">[emot=雷倒/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/再见啦.png)" code="再见啦" title="再见啦" class="emj">[emot=再见啦/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/汗到.png)" code="汗到" title="汗到" class="emj">[emot=汗到/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/晕翻.png)" code="晕翻" title="晕翻" class="emj">[emot=晕翻/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/被激怒.png)" code="被激怒" title="被激怒" class="emj">[emot=被激怒/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/不好意思.png)" code="不好意思" title="不好意思" class="emj">[emot=不好意思/]</div>
					<div class="fl" style="width: 100%;">
						备注：以上表情适用于网页客服与微信平台，发送到微信端后会自动转成微信表情 &nbsp; 
						<a href="/web/help/index.html?item=bot_cmd" class="green" target="_blank">获取更多微信表情</a>
					</div>				
				</div>
			</div>
				<div class="explain-col reply_news hide" style="display: block;">

				<input type="hidden" name="news" id="news" value="">
				<ul>
								</ul>
					<div class="alert-col">可设置多个备选回复(默认为随机回复，关闭随机将优先排列靠前的回复) &nbsp; 没有图文请到 微信管理-&gt;<a href="javascript:void(0);" class="green" onclick="$('#Boxs').hide();$('select').show();SetRContent('MainRight','weixin/weixin_ImgText.asp?Type=1');return false;">图文管理</a> 添加</div>
				</div>
				<div class="explain-col reply_rs hide" style="display: block;">
				<input type="hidden" name="resource" id="resource" value="">
				<ul>
								</ul>
				<div class="alert-col">可设置多个备选回复(默认为随机回复，关闭随机将优先排列靠前的回复) &nbsp; 没有语音素材请到 内容管理-&gt;素材管理--&gt;<a href="javascript:void(0);" onclick="$('#Boxs').hide();$('select').show();SetRContent('MainRight','weixin/Weixin_Sucai.asp?Type=1');return false;" class="green">素材列表</a> 添加</div>
			</div>
			</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/Weixin_CustReply.asp?Type=3',0,'',0,1,'MainRight','/admin/weixin/Weixin_CustReply.asp?Type=1');" class="Button" name="button" id="button" value="提 交" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：WeixinCustReplyAddDo
'作    用：执行添加微信自定义回复
'参    数：
'==============================
Sub WeixinCustReplyAddDo()
	Fk_reply_qtitle			 = FKFun.HTMLEncode(Trim(Request.Form("question[]")))
	Fk_reply_qanswerText	 = FKFun.HTMLEncode(Trim(Request.Form("answer[]")))
	Fk_reply_qanswerNews	 = FKFun.HTMLEncode(Trim(Request.Form("news")))
	Fk_reply_qanswerResource = FKFun.HTMLEncode(Trim(Request.Form("resource")))
	Fk_reply_type			 = Trim(Request.Form("reply_type"))
	Call FKFun.ShowString(replace(Fk_reply_qtitle,",",""),1,100,0,"请输入问题！","问题不能大于100个字节！")
	if Fk_reply_type="0" then
		Call FKFun.ShowString(replace(Fk_reply_qanswerText,",",""),1,500,0,"请输入回复内容！","回复内容不能大于500个字节！")
	elseif Fk_reply_type="2" then
		Call FKFun.ShowString(replace(Fk_reply_qanswerNews,",",""),1,100,0,"请选择回复图文！","回复图文不能大于100个字节！")
	elseif Fk_reply_type="3" then
		Call FKFun.ShowString(replace(Fk_reply_qanswerResource,",",""),1,100,0,"请输入回复语音！","回复语音不能大于100个字节！")
	end if
	if instr(Fk_reply_qtitle,", ,")>0 then Fk_reply_qtitle=trim(replace(Fk_reply_qtitle,", ,",""))
	if instr(Fk_reply_qtitle&",",",,")>0 then Fk_reply_qtitle=trim(replace(Fk_reply_qtitle&",",",,",""))
	if instr(Fk_reply_qanswerText,", ,")>0 then Fk_reply_qanswerText=trim(replace(Fk_reply_qanswerText,", ,",""))
	if instr(Fk_reply_qanswerText&",",",,")>0 then Fk_reply_qanswerText=trim(replace(Fk_reply_qanswerText&",",",,",""))
	Sqlstr="Select * From [Weixin_CustReply]"
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs.AddNew()
		Rs("reply_qtitle")=Fk_reply_qtitle
		Rs("reply_qanswerText")=Fk_reply_qanswerText
		Rs("reply_qanswerNews")=Fk_reply_qanswerNews
		Rs("reply_qanswerResource")=Fk_reply_qanswerResource
		Rs("reply_type")=Fk_reply_type
		Rs.Update()
		Application.UnLock()
		Response.Write("自定义回复添加成功！")
	Rs.Close
End Sub

'==========================================
'函 数 名：WeixinCustReplyEditForm
'作    用：修改微信自定义回复
'参    数：
'==========================================
Sub WeixinCustReplyEditForm()
	Id=Clng(Request.QueryString("Id"))
	Sqlstr="Select * From [Weixin_CustReply] Where id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_reply_qtitle			= Rs("reply_qtitle")
		Fk_reply_qanswerText	= Rs("reply_qanswerText")
		Fk_reply_qanswerNews	= Rs("reply_qanswerNews")
		Fk_reply_qanswerResource= Rs("reply_qanswerResource")
		Fk_reply_type			= Rs("reply_type")
	End If
	Rs.Close
%>
<style type="text/css">
.rndBtn {
height: 30px;
width: 30px;
display: inline-block;
background: url("http://image001.dgcloud01.qebang.cn/website/weixin/userMenuButtons.png") no-repeat scroll 0 0 transparent;
}
.rndBtn.plus {
background-position: -1100px 0;
}
.fr {
float: right;
}
td ul li {
margin: 1px 0;
padding-bottom: 3px;
}
.alert-col {
color: #666;
}

.explain-col {
border: 1px solid #ffbe7a;
zoom: 1;
background: #fffced;
padding: 8px 10px;
line-height: 20px;
}

.hide {
display: none;
}

.reply_text textarea{height:50px;background:none;}

	.emot_box {
		/*
		height: 160px;
		*/
		height: 100%;
		overflow: auto;
	}
		.emot_box .emj {
			float: left;
			width: 102px;
			height: 40px;
			padding-top: 50px;
			background-attachment: scroll;
			background-position: top center;
			background-repeat: no-repeat;
			background-size: 38px 38px;
			text-align: center;
		}
		.explain-col ul li .item {
width: 350px;
height: 35px;
line-height: 35px;
padding-left: 10px;
}
.explain-col ul li {
margin: 5px 0;
padding-bottom: 5px;
border-bottom: 1px dotted #D6D6D6;
}
		.emot_box{width:712px;margin-top:10px;}
		.rndBtn.blkFrd {
background-position: -500px 0;
}
		.rndBtn.blkFrd:hover {
background-position: -500px -50px;
}
.rndBtn.ext.on {
background-position: -650px 0;
}
.rndBtn.ext.on:hover {
background-position: -650px -50px;
}

.table_form .explain-col {
margin: 30px 0 5px 0;
min-height: 50px;
}
</style>
<script type="text/javascript">
	/**
	 * 添加图文
	 * @return
	 */
	function add_news() {
		// 多图文
				
		}


	/**
	 * 添加语音
	 * @return
	 */
	function add_music() {
	}
	
$(document).ready(function(){
		// 回复类型切换
		$('#reply_type1, #reply_type2, #reply_type3').click(function() {
			var id = $(this).attr('id');
			
			$('.reply_text, .reply_news, .reply_rs').hide();

			if(id == 'reply_type2') {
				// 图文
				$('.reply_news').show();
			}else if(id == 'reply_type3') {
				// 语音
				$('.reply_rs').show();
			}else {
				// 文本
				$('.reply_text').show();
			}
		})
		
				// 插入表情代码
		$('#add_emot').click(function() {
			if($(this).is(':checked')) {
				$('.emot_box').removeClass('hide').show();
			}else {
				$('.emot_box').addClass('hide').hide();
			}
		})
		
		// 优先显示图文/语音回复/文本
		if($('.reply_news > ul > li').length > 0) {
			$('#reply_type2').click();
		}else if($('.reply_rs > ul > li').length > 0) {
			$('#reply_type3').click();
		}else if($('.reply_text > ul > li').length > 0) {
			$('#reply_type1').click();
		}
		
		// 移出
		$('.item > .blkFrd').live('click', function() {
			$(this).parent().parent().remove();
		});
		// 下移
		$('.item > .ext.on').live('click', function() {
			var parent = $(this).parent().parent();
			if(parent.next('li').length > 0) {
				parent.before(parent.next('li'));
			}
		});
		
				// 追加问答
		$('.plus').click(function() {
			var parent, html;
			if($(this).hasClass('answer') && $('#reply_type2').is(':checked')) {
				// 图文
				ymPrompt.win({message:'/admin/weixin/weixin_getNewsList.asp?type=1&id=0',
					width : 600,
					height : 350,
					title:'选择图文',
					btn: [['确定','ok'],['关闭','close']],
					maxBtn : true,
					minBtn : true,
					closeBtn : true,
					iframe : true,handler:function(msg){
						if (msg == 'error') {
						
						}else if(msg == 'ok'){ 
						
							if($("iframe").contents().find("input.Checks:checked").length>0){
							
								var html;
								var id, val, box,c;
								box = $('.reply_news > ul');
								c=$("iframe").contents().find("input.Checks:checked");
								for(i = 0; i < c.length; i++) {
									if(c[i].type == 'checkbox' && c[i].name == 'ListId' && c[i].checked) {
										if(box.children().length > 9) {
											alert('图文数量已超出');
											break;
										}
										id = c[i].value;
										if(id) {
											val = $("iframe").contents().find('#news_' + id).val();
											if(id && val && box.find(".item[nid='" + id + "']").length < 1) {
												html = '<li><div class="item" nid="' + id + '"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a>' + val + '</div></li>';
												box.append(html);
											}
										}
									}
								}
								
							}
						}
					}
				})
				return false;

			}else if($(this).hasClass('answer') && $('#reply_type3').is(':checked')) {
				// 语音
				
				ymPrompt.win({message:'/admin/weixin/weixin_getMusicList.asp?type=1&id=0',
					width : 600,
					height : 350,
					title:'选择语音',
					btn: [['确定','ok'],['关闭','close']],
					maxBtn : true,
					minBtn : true,
					closeBtn : true,
					iframe : true,handler:function(msg){
						if (msg == 'error') {
						
						}else if(msg == 'ok'){ 
							$("#resource").val("");
							var items="";
							if($("iframe").contents().find("input.Checks:checked").length>0){
							
								var html;
								var id, val, box,c;
								box = $('.reply_rs > ul');
								c=$("iframe").contents().find("input.Checks:checked");
								for(i = 0; i < c.length; i++) {
									if(c[i].type == 'checkbox' && c[i].name == 'ListId' && c[i].checked) {
										if(box.children().length > 9) {
											alert('语音数量已超出');
											break;
										}
										id = c[i].value;
										if(id) {
											val = $("iframe").contents().find('#news_' + id).val();
											if(id && val && box.find(".item[nid='" + id + "']").length < 1) {
												html = '<li><div class="item" nid="' + id + '"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a>' + val + '</div></li>';
												box.append(html);
											}
										}
									}
								}	
							
								
								$("#resource").val(items);
							}
						}
					}
				})
				return false;
			}else {
				// 文本问答
				parent = $(this).hasClass('answer') ? $(this).next().children('ul'): $(this).next('ul');
				if(parent.children().length >= 5) {
					return;
				}
				html = '<li>' + parent.children('li:first').html() + '</li>'
				parent.append(html).find('li:last > input').val('').removeAttr('readonly');
			}

		})
})

	function updateItems(){
		if($(".reply_news > ul > li").length>0){
			$(".reply_news > ul > li").each(function(i){
				if(i==0){
					items=$(this).children(".item").attr("nid");
				}
				else{
					items=items+","+$(this).children(".item").attr("nid");
				}
				$("#news").val(items);
			})
		}
		else{
			$("#news").val("");
		}
		
		
		if($(".reply_rs > ul > li").length>0){
			$(".reply_rs > ul > li").each(function(i){
				if(i==0){
					items=$(this).children(".item").attr("nid");
				}
				else{
					items=items+","+$(this).children(".item").attr("nid");
				}
				$("#resource").val(items);
			})
		}
		else{
			$("#resource").val("");
		}
	
	}
</script>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/Weixin_CustReply.asp?Type=5" onsubmit="updateItems();return false;"><input type="hidden" value="<%=id%>" name="id"/>
<div id="BoxTop" style="width:98%;"><span>修改自定义回复</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" class="table_form" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">问题：</td>
	        <td><a href="javascript: void(0);" class="rndBtn plus fr question" title="添加问题"></a>
			<ul>
				<%if Fk_reply_qtitle<>"" then
					if instr(Fk_reply_qtitle,",")>0 then
					dim arr
					arr=split(Fk_reply_qtitle,",")
					for i=0 to ubound(arr)%>
					<li><input name="question[]" type="text" class="Input" size="60" maxlength="200" value="<%=replace(arr(i)," ","")%>"/></li>
				<%next
				else%>
				<li><input name="question[]" type="text" class="Input" size="60" maxlength="200" value="<%=Fk_reply_qtitle%>"/></li>
				<%end if
				end if%>
			</ul>
			<div class="alert-col">点击右边的添加按钮，可设置多个相似问题。 </div></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">回复类型：</td>
	        <td><span class="fr">支持同时编辑多种回复类型，回复优先级: 图文 > 语音 > 文本</span><input name="reply_type" type="radio" id="reply_type1" class="Input" value="0" style="verticle-align:middle"/> <label for="reply_type1" style="verticle-align:middle">文本</label> &nbsp; <input name="reply_type" type="radio" id="reply_type2" class="Input" value="2" style="verticle-align:middle"/> <label for="reply_type2" style="verticle-align:middle">图文</label> &nbsp; <input name="reply_type" type="radio" id="reply_type3" class="Input" value="3" style="verticle-align:middle"/> <label for="reply_type3" style="verticle-align:middle">语音</label> </td>
	        </tr>
	    <tr>
	        <td height="25" align="right">回复内容：</td>
	        <td><a href="javascript: void(0);" class="rndBtn plus fr answer" title="添加回复"></a>
				
				<div class="reply_text" style="display: block;">
				<ul>
				<%if Fk_reply_qanswerText<>"" then
					if instr(Fk_reply_qanswerText,", ")>0 then
					arr=split(Fk_reply_qanswerText,", ")
					for i=0 to ubound(arr)%>
					<li><textarea name="answer[]" class="Input" rows="2" cols="56"><%=replace(arr(i)," ","")%></textarea></li>
				<%next
				else%>
				<li><textarea name="answer[]" class="Input" rows="2" cols="56"><%=Fk_reply_qanswerText%></textarea></li>
				<%end if
				else%>				
				<li><textarea name="answer[]" class="Input" rows="2" cols="56"></textarea></li>
				<%end if%>
									
								</ul>
				<div class="alert-col">
					可设置多个备选回复(默认为随机回复，修改<a class="blue" href="javascript: redirect('/robot/learn_param/');"><em>机器人_随机回复</em></a>可关闭随机回复机制) &nbsp; 
					<input type="checkbox" id="add_emot" value="" class="hide"> <label for="add_emot" class="hide">插入表情代码</label> &nbsp; 
									<a href="/web/help/index.html?item=bot_cmd" class="blue hide" target="_blank">学习高级培训指令</a>
								</div>

				<!--textarea name="answer" id="answer" rows="3" cols="58" class="validate[required]"></textarea-->
					
				<div class="emot_box explain-col hide" title="官方表情发送到微信端后会自动转成微信表情">
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/哈哈.png)" code="哈哈" title="哈哈" class="emj">[emot=哈哈/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/得瑟.png)" code="得瑟" title="得瑟" class="emj">[emot=得瑟/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/示好.png)" code="示好" title="示好" class="emj">[emot=示好/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/哼.png)" code="哼" title="哼" class="emj">[emot=哼/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/嚎啕大哭.png)" code="嚎啕大哭" title="嚎啕大哭" class="emj">[emot=嚎啕大哭/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/不懂.png)" code="不懂" title="不懂" class="emj">[emot=不懂/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/俏皮.png)" code="俏皮" title="俏皮" class="emj">[emot=俏皮/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/别惹我.png)" code="别惹我" title="别惹我" class="emj">[emot=别惹我/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/呜呜.png)" code="呜呜" title="呜呜" class="emj">[emot=呜呜/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/打瞌睡.png)" code="打瞌睡" title="打瞌睡" class="emj">[emot=打瞌睡/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/雷倒.png)" code="雷倒" title="雷倒" class="emj">[emot=雷倒/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/再见啦.png)" code="再见啦" title="再见啦" class="emj">[emot=再见啦/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/汗到.png)" code="汗到" title="汗到" class="emj">[emot=汗到/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/晕翻.png)" code="晕翻" title="晕翻" class="emj">[emot=晕翻/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/被激怒.png)" code="被激怒" title="被激怒" class="emj">[emot=被激怒/]</div>
<div style="background-image:url(http://static.v5kf.com/images/emoji/default/不好意思.png)" code="不好意思" title="不好意思" class="emj">[emot=不好意思/]</div>
					<div class="fl" style="width: 100%;">
						备注：以上表情适用于网页客服与微信平台，发送到微信端后会自动转成微信表情 &nbsp; 
						<a href="/web/help/index.html?item=bot_cmd" class="green" target="_blank">获取更多微信表情</a>
					</div>				
				</div>
			</div>
				<div class="explain-col reply_news hide">

				<input type="hidden" name="news" id="news" value="<%=Fk_reply_qanswerNews%>">
				<ul>
				
				<%if Fk_reply_qanswerNews<>"" then
					if instr(Fk_reply_qanswerNews,",")>0 then
					arr=split(Fk_reply_qanswerNews,",")
					for i=0 to ubound(arr)%>
					<li><div class="item" nid="<%=arr(i)%>"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a><%=getInfo(arr(i))%><div></li>
				<%next
				else%>
				<li><div class="item" nid="<%=Fk_reply_qanswerNews%>"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a><%=getInfo(Fk_reply_qanswerNews)%><div></li>
				<%end if
				end if%>
								</ul>
					<div class="alert-col">可设置多个备选回复(默认为随机回复，关闭随机将优先排列靠前的回复) &nbsp; <!--可设置多个备选回复，优先匹配排列靠前的回复，超过有效期将选择备选图文。-->没有图文请到 微信管理-&gt;<a href="javascript:void(0);" onclick="$('#Boxs').hide();$('select').show();SetRContent('MainRight','weixin/weixin_ImgText.asp?Type=1');return false;" class="green">图文管理</a> 添加</div>
				</div>
				<div class="explain-col reply_rs hide">
				<input type="hidden" name="resource" id="resource" value="">
				<ul>
				
				<%if Fk_reply_qanswerResource<>"" then
					if instr(Fk_reply_qanswerResource,",")>0 then
					arr=split(Fk_reply_qanswerResource,",")
					for i=0 to ubound(arr)%>
					<li><div class="item" nid="<%=arr(i)%>"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a><%=getMInfo(arr(i))%><div></li>
				<%next
				else%>
				<li><div class="item" nid="<%=Fk_reply_qanswerResource%>"><a class="rndBtn blkFrd fr" title="移出"></a><a class="rndBtn ext on fr" title="下移"></a><%=getMInfo(Fk_reply_qanswerResource)%><div></li>
				<%end if
				end if%>
								</ul>
				<div class="alert-col">可设置多个备选回复(默认为随机回复，关闭随机将优先排列靠前的回复) &nbsp; 没有语音素材请到 内容管理-&gt;素材管理--&gt;<a href="javascript:void(0);" onclick="$('#Boxs').hide();$('select').show();SetRContent('MainRight','weixin/Weixin_Sucai.asp?Type=1');return false;" class="green">素材列表</a>素材列表</a> 添加</div>
			</div>
			</td>
	        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/Weixin_CustReply.asp?Type=5',0,'',0,1,'MainRight','/admin/weixin/Weixin_CustReply.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
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

function getMInfo(id)
	getMInfo=""
	Sqlstr="Select Sucai_title From [weixin_Sucai] Where id=" & Id
	set rs=conn.execute(sqlstr)
	if not rs.eof then
	getMInfo=rs("Sucai_title")
	end if
	rs.close
end function
'==========================================
'函 数 名：WeixinCustReplyMake()
'作    用：生成微信回复
'参    数：
'==========================================
Sub WeixinCustReplyMake()
dim wx_AppId,wx_AppSecret
set rs=conn.execute("select wx_AppId,wx_AppSecret from weixin_config")
if not rs.eof then
	wx_AppId	 = rs("wx_AppId")
	wx_AppSecret = rs("wx_AppSecret")
end if
rs.close
%>
<form id="WeixinAdd" name="WeixinAdd" method="post" action="/admin/weixin/Weixin_CustReply.asp?Type=8" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>生成回复</span></div>
<div id="BoxContents" style="width:98%;">
	<div style="margin:20px;padding:10px;word-wrap:break-word;word-break:break-all;border: 1px solid #ffbe7a;background: #fffced;">
	   请首先确保已创建回复<br>
	   请到公众号官方后台->高级功能->开发模式 中获取以下信息
	   </div>
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right">AppId</td>
	        <td>&nbsp;<input name="Fk_wx_AppId" type="text" class="Input" id="Fk_wx_AppId" size="40" value="<%=wx_AppId%>"/></td>
	    </tr>
	    <tr>
	        <td height="25" align="right">AppSecret</td>
	        <td>&nbsp;<input name="Fk_wx_AppSecret" class="Input" type="text" id="Fk_wx_AppSecret" value="<%=wx_AppSecret%>" size="40"></td>
	    </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('WeixinAdd','/admin/weixin/Weixin_CustReply.asp?Type=8',0,'',0,1,'MainRight','/admin/weixin/Weixin_CustReply.asp?Type=1');" class="Button" name="button" id="button" value="生 成" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub


'==============================
'函 数 名：WeixinCustReplyMakeDo
'作    用：执行生成自定义回复
'参    数：
'==============================
Sub WeixinCustReplyMakeDo()
	dim wx_AppId,wx_AppSecret
	wx_AppId	= FKFun.HTMLEncode(Trim(Request.Form("Fk_wx_AppId")))
	wx_AppSecret= FKFun.HTMLEncode(Trim(Request.Form("Fk_wx_AppSecret")))
	Call FKFun.ShowString(wx_AppId,1,50,0,"请输入AppId！","AppId不能大于8个字节(4个汉字)！")
	Call FKFun.ShowString(wx_AppSecret,1,50,0,"请输入AppSecret！","AppSecret不能大于50个字符！")
	' Sqlstr="Select * From [weixin_config]"
	' Rs.Open Sqlstr,Conn,1,3
	' Application.Lock()
	' if rs.eof then
		' rs.addnew()
	' end if
	' Rs("wx_AppId")=wx_AppId
	' Rs("wx_AppSecret")=wx_AppSecret
	' Rs.Update()
	' Application.UnLock()
	' Rs.Close
	dim jsonHtml,subrs,i,j
	jsonHtml="{"&vbcrlf
	set rs=conn.execute("select * from Weixin_CustReply where menuParent=0")
	if not rs.eof then
		jsonHtml=jsonHtml&" ""button"":["&vbcrlf
		i=0
		do while not rs.eof
			if i<>0 then jsonHtml=jsonHtml&","&vbcrlf
			set subrs=conn.execute("select * from Weixin_CustReply where menuParent="&rs("id"))
			if not subrs.eof then	'存在子自定义回复
				jsonHtml=jsonHtml&"{"&vbcrlf
				jsonHtml=jsonHtml&"""name"":"""&rs("menuName")&""","&vbcrlf
				jsonHtml=jsonHtml&"""sub_button"":["&vbcrlf
				j=0
				do while not subrs.eof
					if j<>0 then jsonHtml=jsonHtml&","&vbcrlf
					jsonHtml=jsonHtml&"{"&vbcrlf
					jsonHtml=jsonHtml&"""type"":"""&subrs("menuType")&""","&vbcrlf
					jsonHtml=jsonHtml&"""name"":"""&subrs("menuName")&""","&vbcrlf
					if subrs("menuType")="view" then
						jsonHtml=jsonHtml&"""url"":"""&subrs("menuOnEvent")&""""&vbcrlf
					else
						jsonHtml=jsonHtml&"""key"":"""&subrs("menuOnEvent")&""""&vbcrlf
					end if
					jsonHtml=jsonHtml&"}"&vbcrlf
					
					j=j+1
				subrs.movenext
				loop
				jsonHtml=jsonHtml&"]"&vbcrlf
				jsonHtml=jsonHtml&"}"&vbcrlf
			else
			
				jsonHtml=jsonHtml&"{"&vbcrlf
				jsonHtml=jsonHtml&"""type"":"""&rs("menuType")&""","&vbcrlf
				jsonHtml=jsonHtml&"""name"":"""&rs("menuName")&""","&vbcrlf
				if rs("menuType")="view" then
					jsonHtml=jsonHtml&"""url"":"""&rs("menuOnEvent")&""""&vbcrlf
				else
					jsonHtml=jsonHtml&"""key"":"""&rs("menuOnEvent")&""""&vbcrlf
				end if
				jsonHtml=jsonHtml&"}"&vbcrlf
				
			end if
			subrs.close
			i=i+1
		rs.movenext
		loop
		jsonHtml=jsonHtml&"]"&vbcrlf
		jsonHtml=jsonHtml&"}"&vbcrlf
		dim access_token,returnMsg
		access_token=DoGet("https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid="&wx_AppId&"&secret="&wx_AppSecret)
		access_token=strCut(access_token,"access_token"":""","""",2)
		returnMsg=DoPost("https://api.weixin.qq.com/cgi-bin/menu/create?access_token="&access_token,jsonHtml)
		if returnMsg="{""errcode"":0,""errmsg"":""ok""}" then
			Response.Write("自定义回复生成成功！请重启微信查看自定义回复效果")
		else
			Response.Write("自定义回复生成失败！请重试")
		end if
	else
		response.write "还未创建回复，请先创建好自定义回复再生成"
	end if
	rs.close

End Sub

Function ByteToStr(vIn)
	Dim strReturn,i,ThisCharCode,innerCode,Hight8,Low8,NextCharCode
	strReturn = "" 
	For i = 1 To LenB(vIn)
	ThisCharCode = AscB(MidB(vIn,i,1))
	If ThisCharCode < &H80 Then
	strReturn = strReturn & Chr(ThisCharCode)
	Else
	NextCharCode = AscB(MidB(vIn,i+1,1))
	strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
	i = i + 1
	End If
	Next
	ByteToStr = strReturn 
End Function

Function DoGet(url)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	'.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	.Send()
	DoGet = .ResponseText
	End With
	Set Http = Nothing
	'DoPost=ByteToStr(DoPost)
	if err then 
		err.clear
		DoGet=""
	end if
End Function

Function DoPost(url,PostStr)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	.Send(PostStr)
	DoPost = .ResponseBody
	End With
	Set Http = Nothing
	DoPost=ByteToStr(DoPost)
	if err then 
		err.clear
		DoPost=""
	end if
End Function

	'写入文件法调试
	public Function WriteFile(content)
		dim filepath,fso,fopen
		filepath=server.mappath(".")&"\wx.txt"
		Set fso = Server.CreateObject("scripting.FileSystemObject")
		set fopen=fso.OpenTextFile(filepath, 8 ,true)
		content = content&vbcrlf&"************line seperate("&now()&")*****************"
		fopen.writeline(content)
		set fso=nothing
		set fopen=Nothing
	End Function

Function strCut(strContent,StartStr,EndStr,CutType)
	Dim strHtml,S1,S2
	strHtml = strContent
	On Error Resume Next
	Select Case CutType
	Case 1
		S1 = InStr(strHtml,StartStr)
		S2 = InStr(S1,strHtml,EndStr)+Len(EndStr)
	Case 2
		S1 = InStr(strHtml,StartStr)+Len(StartStr)
		S2 = InStr(S1,strHtml,EndStr)
	End Select
	If Err Then
		strCute = ""
		Err.Clear
		Exit Function
	Else
		strCut = Mid(strHtml,S1,S2-S1)
	End If
End Function

'==============================
'函 数 名：WeixinCustReplyEditDo
'作    用：执行修改自定义回复
'参    数：
'==============================
Sub WeixinCustReplyEditDo()
	Id=Trim(Request.Form("Id"))
	Fk_reply_qtitle			 = FKFun.HTMLEncode(Trim(Request.Form("question[]")))
	Fk_reply_qanswerText	 = FKFun.HTMLEncode(Trim(Request.Form("answer[]")))
	Fk_reply_qanswerNews	 = FKFun.HTMLEncode(Trim(Request.Form("news")))
	Fk_reply_qanswerResource = FKFun.HTMLEncode(Trim(Request.Form("resource")))
	Fk_reply_type			 = Trim(Request.Form("reply_type"))
	Call FKFun.ShowString(replace(Fk_reply_qtitle,",",""),1,100,0,"请输入问题！","问题不能大于100个字节！")
	if Fk_reply_type="0" then
		Call FKFun.ShowString(replace(Fk_reply_qanswerText,",",""),1,500,0,"请输入回复内容！","回复内容不能大于500个字节！")
	elseif Fk_reply_type="2" then
		Call FKFun.ShowString(replace(Fk_reply_qanswerNews,",",""),1,100,0,"请选择回复图文！","回复图文不能大于100个字节！")
	elseif Fk_reply_type="3" then
		Call FKFun.ShowString(replace(Fk_reply_qanswerResource,",",""),1,100,0,"请输入回复语音！","回复语音不能大于100个字节！")
	end if
	if instr(Fk_reply_qtitle,", ,")>0 then Fk_reply_qtitle=trim(replace(Fk_reply_qtitle,", ,",""))
	if instr(Fk_reply_qtitle&",",",,")>0 then Fk_reply_qtitle=trim(replace(Fk_reply_qtitle&",",",,",""))
	if instr(Fk_reply_qanswerText,", ,")>0 then Fk_reply_qanswerText=trim(replace(Fk_reply_qanswerText,", ,",""))
	if instr(Fk_reply_qanswerText&",",",,")>0 then Fk_reply_qanswerText=trim(replace(Fk_reply_qanswerText&",",",,",""))
	Sqlstr="Select * From [Weixin_CustReply] where id="&id
	Rs.Open Sqlstr,Conn,1,3
		Application.Lock()
		Rs("reply_qtitle")=Fk_reply_qtitle
		Rs("reply_qanswerText")=Fk_reply_qanswerText
		Rs("reply_qanswerNews")=Fk_reply_qanswerNews
		Rs("reply_qanswerResource")=Fk_reply_qanswerResource
		Rs("reply_type")=Fk_reply_type
		Rs.Update()
		Application.UnLock()
		Response.Write("自定义回复修改成功！")
	Rs.Close
End Sub

'==============================
'函 数 名：WeixinCustReplyDelDo
'作    用：执行删除微信自定义回复
'参    数：
'==============================
Sub WeixinCustReplyDelDo()
	Id=Trim(Request("Id"))
	Call FKFun.ShowNum(Id,"系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Weixin_CustReply] Where id in("& Id &")"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("微信自定义回复删除成功！")
	Else
		Response.Write("微信自定义回复不存在！")
	End If
	Rs.Close
End Sub
%><!--#Include File="../../Code.asp"-->