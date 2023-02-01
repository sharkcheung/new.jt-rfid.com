<!--#Include File="../AdminCheck.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<%
'==========================================
'文 件 名：Weixin_Set.asp
'文件用途：微信接口设置拉取页面
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
dim wx_token,wx_raw_id,wx_AppId,wx_AppSecret,wx_url,wx_Subscribe,wx_Repeat,wx_Random

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call WeixinSet() '微信接口配置
	Case 2
		Call WeixinSetDo() '微信接口配置执行修改
	Case Else
		Response.Write("没有找到此功能项！")
End Select

Function randKey(obj) 
	 Dim char_array(80) 
	 Dim temp 
	 For i = 0 To 9  
	  char_array(i) = Cstr(i) 
	 Next 
	 For i = 10 To 35 
	  char_array(i) = Chr(i + 55) 
	 Next 
	 For i = 36 To 61 
	  char_array(i) = Chr(i + 61) 
	 Next 
	 Randomize 
	 For i = 1 To obj 
	  'rnd函数返回的随机数在0~1之间，可等于0，但不等于1 
	  '公式：int((上限-下限+1)*Rnd+下限)可取得从下限到上限之间的数，可等于下限但不可等于上限 
	  temp = temp&char_array(int(62 - 0 + 1)*Rnd + 0) 
	 Next 
	 randKey = temp 
End Function



'==============================
'函 数 名：WeixinSetDo
'作    用：执行微信接口配置修改
'参    数：
'==============================
Sub WeixinSetDo()
	wx_token	= FKFun.HTMLEncode(Trim(Request.Form("wx_token")))
	wx_raw_id	= FKFun.HTMLEncode(Trim(Request.Form("wx_raw_id")))
	wx_AppId	= FKFun.HTMLEncode(Trim(Request.Form("wx_AppId")))
	wx_AppSecret= FKFun.HTMLEncode(Trim(Request.Form("wx_AppSecret")))
	wx_url		= Trim(Request.Form("wx_url"))
	wx_Subscribe= Trim(Request.Form("wx_Subscribe"))
	wx_Repeat	= Trim(Request.Form("wx_Repeat"))
	wx_Random	= Trim(Request.Form("wx_Random"))
	
	Call FKFun.ShowString(wx_raw_id,1,50,0,"微信原始账号为必填项","微信原始账号不能大于50个字符！")
	Sqlstr="Select * From [weixin_config]"
	Rs.Open Sqlstr,Conn,1,3
	Application.Lock()
	If Rs.Eof Then
		Rs.AddNew()
	End If
	Rs("wx_token")		= wx_token
	Rs("wx_raw_id")		= wx_raw_id
	Rs("wx_AppId")		= wx_AppId
	Rs("wx_AppSecret")	= wx_AppSecret
	Rs("wx_url")		= wx_url
	Rs("wx_Subscribe")	= wx_Subscribe
	Rs("wx_Repeat")		= wx_Repeat
	Rs("wx_Random")		= wx_Random
	Rs.Update()
	Application.UnLock()
	Response.Write("设置成功！")
	Rs.Close
End Sub

'==========================================
'函 数 名 WeixinSet()
'作    用 读取微信接口设置
'参    数
'==========================================
Sub WeixinSet()
set rs=conn.execute("select * from weixin_config")
if not rs.eof then
	wx_token	= trim(rs("wx_token")&" ")
	wx_raw_id	= trim(rs("wx_raw_id")&" ")
	wx_AppId	= trim(rs("wx_AppId")&" ")
	wx_AppSecret= trim(rs("wx_AppSecret")&" ")
	wx_Repeat	= trim(rs("wx_Repeat")&" ")
	wx_Random	= trim(rs("wx_Random")&" ")
	wx_Subscribe	= trim(rs("wx_Subscribe")&" ")
end if
rs.close
set rs=nothing
if wx_Repeat="" then
	wx_Repeat= "0"
end if
if wx_Random="" then
	wx_Random= "0"
end if
if wx_token="" then
	wx_token=ucase(randKey(32))
end if
%>
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
		$('.icon_ui_btn').live('click', function() {
			search_ui($(this));
		});
		
		// 更新封面
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
		$('.addNews').click(function() {
			ymPrompt.win({message:'/admin/weixin/weixin_getNewslist.asp?type=1&id=0',
				width : 600,
				height :350,
				title:'选择图文',
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
				
		});
		
		// 文字
		$('.addText').click(function() {
			ymPrompt.win({message:'/admin/weixin/weixin_SetSubcrib.asp?type=1&id=0',
				width : 400,
				height : 250,
				title:'编辑信息',
				btn: [['确定','ok'],['关闭','close']],
				maxBtn : true,
				minBtn : true,
				closeBtn : true,
				iframe : true,handler:function(msg){
					if (msg == 'error') {
					
					}else if(msg == 'ok'){ 
						if($("iframe").contents().find("textarea").length>0){
							var text;
							text=$("iframe").contents().find("textarea").val();
							$('#wx_Subscribe').val(text);
						}
					}
				}
			});return false;
				
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

<form id="SystemSet" name="SystemSet" method="post" action="Weixin_Set.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>微信接口设置</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="/admin/images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
    
       <div style="margin:20px;padding:10px;word-wrap:break-word;word-break:break-all;border: 1px solid #ffbe7a;background: #fffced;">
	   <pre>温馨提示：
系统提供微信公众平台上自动回复的接口，请按以下步骤操作 
1. 请先登录 微信公众平台 开通公众账号 
2. 进入[高级功能]，在关闭[编辑模式]后，开启并进入[开发模式]配置页面 
3. 把下列的接口配置信息 <span style='color: #0C0;'>URL</span> 和 <span style='color: #0C0;'>Token</span> 写入微信公众平台的配置页面后提交即可 
4. 如果您的公众号拥有高级接口权限，请将公众平台中的 <span style='color: #0C0;'>开发者凭据(AppId和AppSecret)</span> 设置在此页面</pre>
	   </div>
	<table width="90%" border="0" align="center" id="table001"  cellpadding="0" cellspacing="0">
        <tr>
            <td height="25" align="right" class="MainTableTop">URL</td>
            <td><input name="wx_url" type="text" class="Input" id="wx_url" value="http://<%=Request.ServerVariables("Http_Host")%>/admin/weixin/" size="40"  style="vertical-align:middle" readonly/></td>
            <td></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">Token</td>
            <td><input name="wx_token" type="text" class="Input" id="wx_token" value="<%=wx_token%>" size="40"  style="vertical-align:middle" readonly/></td>
            <td></td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">微信原始账号</td>
            <td><input name="wx_raw_id" type="text" class="Input" id="wx_raw_id" value="<%=wx_raw_id%>" size="40"  style="vertical-align:middle"/></td>
            <td>即微信后台账户信息中显示的原始ID</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">AppId</td>
            <td><input name="wx_AppId" type="text" class="Input" id="wx_AppId" value="<%=wx_AppId%>" size="40"  style="vertical-align:middle"/></td>
            <td>选填，请到微信后台->高级功能->开发模式的开发者凭据中获取</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">AppSecret</td>
            <td><input name="wx_AppSecret" type="text" class="Input" id="wx_AppSecret" size="40" value="<%=wx_AppSecret%>"/></td>
            <td>同上，配置AppId和AppSecret可以配置自定义菜单和调用相应的高级接口</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">微信开场白</td>
            <td><textarea name="wx_Subscribe" id="wx_Subscribe" rows="4" cols="35" readonly><%=wx_Subscribe%></textarea> <br><input type="button" value="回复文本" class="Button addText"> <input type="button" value="回复图文" class="Button addNews"></td>
            <td>注：用户关注您的微信账号后将收到此内容，留空则不做处理</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">随机回复</td>
            <td><input name="wx_Random" type="radio" class="Input" id="wx_Random" value="1" style="vertical-align:middle;" <%if wx_Random="1" then response.write "checked"%>/><label for="wx_Random" style="vertical-align:middle;">开启</label> &nbsp;<input name="wx_Random" type="radio" class="Input" style="vertical-align:middle;" id="wx_Random1" value="0" <%if wx_Random="0" then response.write "checked"%>/><label for="wx_Random1" style="vertical-align:middle;">关闭</label></td>
            <td>注：关闭则默认回复第一条内容，开启则随机挑选回复内容</td>
        </tr>
        <tr>
            <td height="25" align="right" class="MainTableTop">重复机制</td>
            <td><input name="wx_Repeat" type="text" class="Input" id="wx_Repeat" size="40" value="<%=wx_Repeat%>"/></td>
            <td> 注：0=允许重复提问，1=重复一定次数后不再处理</td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('SystemSet','/admin/weixin/Weixin_Set.asp?Type=2',0,'',0,0,'','');" class="Button" name="button" id="button" value="保存设置" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub
%><!--#Include File="../../Code.asp"-->