<!--#Include File="../AdminCheck.asp"-->
<!--#Include File="CheckUpdate.asp"-->
<!--#Include File="XMLUpload.asp"-->
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
dim wx_token,wx_raw_id,wx_AppId,wx_AppSecret,wx_url,wx_Subscribe,wx_Repeat,wx_Random,wx_NoneReply
err.clear
on error resume next
set rs=conn.execute("select id from weixin_MassSend")
if err then
err.clear
conn.execute("Create TABLE weixin_MassSend(id integer primary key AUTO_INCREMENT, [字段名2] MEMO, [字段名3] COUNTER NOT NULL, [字段名4] DATETIME, [字段名5] TEXT(200), [字段名6] TEXT(200)")
end if
'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call Weixin_MassSendList() '微信群发
	Case 2
		Call Weixin_MassSend() '微信群发
	Case 3
		Call Weixin_MassSendDo() '微信群发执行
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

private Function GetURL(url)
    On Error Resume Next 
	dim objXML
    Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")  
	objXML.open "GET",url,false  
	objXML.send()  
	GetURL=objXML.responseText
	if err then
		err.clear
		GetURL=""
	end if
End Function

private Function ByteToStr(vIn)
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

private Function DoPost(url,PostStr)
	dim Http
	on error resume next
	Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
	With Http
	.Open "POST", url, false ,"" ,""
	.setRequestHeader "Content-Type","application/x-www-form-urlencoded;charset=utf-8"
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
private sub WriteFile(content)
	dim fso,fopen,filepath
	on error resume next
	filepath=server.mappath(".")&"\wx.txt"
	Set fso = Server.CreateObject("scripting.FileSystemObject")
	set fopen=fso.OpenTextFile(filepath, 8 ,true)
	content = content&vbcrlf&"************line seperate("&now()&")*****************"
	fopen.writeline(content)
	if err then
		fopen.writeline(err.description)
		err.clear
		response.end
	end if
	set fso=nothing
	set fopen=Nothing
	
End sub

'==============================
'函 数 名：WeixinSetDo
'作    用：执行微信群发
'参    数：
'==============================
Sub Weixin_MassSendDo()
	on error resume next
	'获取access_token
	dim wx_AppId,wx_AppSecret,MEDIA_ID
	set rs=conn.execute("select top 1 wx_AppId,wx_AppSecret from weixin_config")
	if not rs.eof then
		wx_AppId=rs(0)
		wx_AppSecret=rs(1)
	end if
	rs.close
	dim access_token,obj,url,returnstr,strdata
	url="https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid="&wx_AppId&"&secret="&wx_AppSecret
	returnstr=GetURL(url)
	Set obj = parseJSON(returnstr)
	access_token=obj.access_token
	'wxObj.WriteFile("access_token===>"&access_token)
	set obj=nothing
	
	' url="https://api.weixin.qq.com/cgi-bin/getcallbackip?access_token="&access_token
	' returnstr=GetURL(url)
	' call WriteFile(returnstr)
	' response.write returnstr
	' response.end
	
	' dim strreturn,strdata
	' strdata="{\n\r$$begin_date$$:$$2011-12-02$$,\n\r$$end_date$$:$$2014-12-07$$\n\r}\n\r}"
	' returnstr=DoPost("https://api.weixin.qq.com/datacube/getusersummary?access_token="&access_token&"",strdata)
	
	'添加客服接口
	strdata="{\n\r$$kf_account$$:$$test1@test$$,\n\r$$nickname$$:$$kefu1$$,\n\r$$password$$:$$pswmd5$$\n\r}"
	strdata=replace(strdata,"$$","""")
	strdata=replace(strdata,"\n\r",vbcrlf)
	returnstr=DoPost("https://api.weixin.qq.com/customservice/kfaccount/add?access_token="&access_token&"",strdata)
	response.write returnstr
	
	'群发预览接口
	' strdata="{\n\r$$touser$$:$$oCulat7l7wx_XpkJRyKH09rmmDMo$$,\n\r$$text$$:{\n\r$$content$$:$$123$$\n\r},\n\r$$msgtype$$:$$text$$\n\r}"
	' strdata=replace(strdata,"$$","""")
	' strdata=replace(strdata,"\n\r",vbcrlf)
	' strreturn=DoPost("https://api.weixin.qq.com/cgi-bin/message/mass/preview?access_token="&access_token&"",strdata)
	' Set obj = parseJSON(strreturn)
	' access_token=obj.access_token
	' set obj=nothing
	
	'上传多媒体文件接口示例
	' Dim UploadData
	' Set UploadData = New XMLUploadImpl
	' UploadData.Charset = "utf-8"
	'语音
	' UploadData.AddFile "ImgFile", Server.MapPath("655928671080064.mp3"), "audio/mp3", GetFileBinary(Server.MapPath("655928671080064.mp3"))
	'图片
	UploadData.AddFile "ImgFile", Server.MapPath("uploadify-cancel.png"), "image/png", GetFileBinary(Server.MapPath("uploadify-cancel.png"))
	' returnstr= UploadData.Upload("http://file.api.weixin.qq.com/cgi-bin/media/upload?access_token="&access_token&"&type=voice")
	' Set UploadData = Nothing
	' Set obj = parseJSON(returnstr)
	' MEDIA_ID=obj.MEDIA_ID
	' set obj=nothing
	
	'下载多媒体文件接口示例
	' url="http://file.api.weixin.qq.com/cgi-bin/media/get?access_token="&access_token&"&media_id="&MEDIA_ID	
	' dim objXML,contenttype,objAdostream
    ' Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")  
	' objXML.open "GET",url,false  
	' objXML.send()  
	' returnstr=objXML.ResponseBody
	' contenttype = lcase(objXML.getResponseHeader("content-type"))'获取响应头类型
	' set objXML=nothing
	
	' if instr(contenttype,"image/")=1 then '为图片
		' set objAdostream=server.createobject("ADODB.Stream")
		' objAdostream.Open()
		' objAdostream.type=1
		' objAdostream.Write(returnstr)
		' objAdostream.SaveToFile(server.MapPath(MEDIA_ID&".jpg"))
		' objAdostream.SetEOS
		' set objAdostream=nothing
	' end if
	
	
	
	' response.ContentType=contenttype'设置响应头
    ' response.BinaryWrite returnstr'输出图片2进制数据
	
	'call WriteFile(returnstr)
	' wx_token	= FKFun.HTMLEncode(Trim(Request.Form("wx_token")))
	' wx_raw_id	= FKFun.HTMLEncode(Trim(Request.Form("wx_raw_id")))
	' wx_AppId	= FKFun.HTMLEncode(Trim(Request.Form("wx_AppId")))
	' wx_AppSecret= FKFun.HTMLEncode(Trim(Request.Form("wx_AppSecret")))
	' wx_url		= Trim(Request.Form("wx_url"))
	' wx_Subscribe= Trim(Request.Form("wx_Subscribe"))
	' wx_NoneReply= Trim(Request.Form("wx_NoneReply"))
	' wx_Repeat	= Trim(Request.Form("wx_Repeat"))
	' wx_Random	= Trim(Request.Form("wx_Random"))
	
	' Call FKFun.ShowString(wx_raw_id,1,50,0,"微信原始账号为必填项","微信原始账号不能大于50个字符！")
	' Sqlstr="Select * From [weixin_config]"
	' Rs.Open Sqlstr,Conn,1,3
	' Application.Lock()
	' If Rs.Eof Then
		' Rs.AddNew()
	' End If
	' Rs("wx_token")		= wx_token
	' Rs("wx_raw_id")		= wx_raw_id
	' Rs("wx_AppId")		= wx_AppId
	' Rs("wx_AppSecret")	= wx_AppSecret
	' Rs("wx_url")		= wx_url
	' Rs("wx_Subscribe")	= wx_Subscribe
	' Rs("wx_NoneReply")	= wx_NoneReply
	' Rs("wx_Repeat")		= wx_Repeat
	' Rs("wx_Random")		= wx_Random
	' Rs.Update()
	' Application.UnLock()
	' Response.Write("设置成功！")
	' Rs.Close
End Sub

private function CheckFields(FieldsName,TableName)
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

'==========================================
'函 数 名：WeixinImgTextList()
'作    用：微信图文列表
'参    数：
'==========================================
Sub Weixin_MassSendList()
Session("NowPage")=FkFun.GetNowUrl()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false;">群发列表</a></li>
        <li><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/weixin_MassSend.asp?Type=2');return false;">添加</a></li>
    </ul>
</div>
<div id="ListContent">
    <form name="DelList" id="DelList" method="post" action="Down.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">内容</td>
            <td align="center" class="ListTdTop">状态</td>
            <td align="center" class="ListTdTop">时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	on error resume next
	set rs=conn.execute("Select * From [weixin_MassSendList]")
	if err then
		conn.execute("create table weixin_MassSendList(id AUTOINCREMENT(1,1) PRIMARY KEY,mass_type int,mass_content memo null,mass_statu int null ,mass_sendtime date null,mass_sendNums int null,mass_sendok int null,mass_sendfail int null)")
	end if
	rs.close
	err.clear
	Set Rs=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [weixin_MassSendList] Order By id desc"
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
            <td height="20" align="center"><%if Rs("mass_type")=0 then response.write "[文字]"%><%=Rs("mass_content")%></td>
            <td ><%=Rs("mass_statu")%></td>
            <td align="center"><%=Rs("mass_sendtime")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/weixin_ImgText.asp?Type=4&Id=<%=Rs("id")%>');return false;"><img src="/admin/images/edit.png" title="编辑"></a> &nbsp;<a href="javascript:void(0);" onclick="ShowBox('/admin/weixin/weixin_ImgText.asp?Type=10&Id=<%=Rs("id")%>');return false;"  title="预览"><img src="/admin/images/yulan.png"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>        <tr>
            <td height="30" colspan="8">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style='text-indent:10px;vertical-align:middle'> 全选
            <input type="submit" value="排序" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_ImgText.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'>
			<input type="submit" value="启用" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_ImgText.asp?Type=8',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'>
			<input type="submit" value="禁用" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_ImgText.asp?Type=9',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'>
			<input type="submit" value="删除" class="Button" onClick="if($('input.Checks:checked').length<1){alert('请先选择要批量操作的数据！');return false};Sends('DelList','/admin/weixin/weixin_ImgText.asp?Type=6',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" style='vertical-align:middle'></td>
        </tr>
		 <tr>
            <td height="30" colspan="8"><%Call FKFun.ShowPageCode("/admin/weixin/weixin_MassSend.asp?Type=1&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="4" align="center">暂无群发消息</td>
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
'函 数 名 WeixinMassSend()
'作    用 微信群发
'参    数
'==========================================
Sub Weixin_MassSend()
if not CheckFields("wx_NoneReply","weixin_config") then
	conn.execute("alter table weixin_config add column wx_NoneReply varchar(200) null")
end if
set rs=conn.execute("select * from weixin_config")
if not rs.eof then
	wx_token	= trim(rs("wx_token")&" ")
	wx_raw_id	= trim(rs("wx_raw_id")&" ")
	wx_AppId	= trim(rs("wx_AppId")&" ")
	wx_AppSecret= trim(rs("wx_AppSecret")&" ")
	wx_Repeat	= trim(rs("wx_Repeat")&" ")
	wx_Random	= trim(rs("wx_Random")&" ")
	wx_Subscribe	= trim(rs("wx_Subscribe")&" ")
	wx_NoneReply	= trim(rs("wx_NoneReply")&" ")
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

		var counter = $("#wx_MassMsg").val().length; //获取文本域的字符串长度
		//$(".editor_tip span").text(600 - counter);

		$("#wx_MassMsg").keyup(function() {
			var text = $("#wx_MassMsg").val();
			var counter = text.length;
			if(counter>600){
				$(".editor_tip").html("已超出<span style=\"color: #C00\">"+counter+"</span>字");
			}
			else{
				$(".editor_tip").html("你还可以输入<span>"+(600 - counter)+"</span>字");
			}
		});

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

		$(".bntType .Button").click(function(){
			$(this).blur();
			$(".Button").removeClass("hover");
			$(this).addClass("hover");
		})
		
		// 无匹配回复多图文
		$('.addVoice').click(function() {
			ymPrompt.win({message:'/admin/weixin/weixin_getNewslist.asp?type=2&id=0',
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
							$('#wx_NoneReply').val("[wx_news="+id+"]");
						}
					}
				}
			});return false;
				
		});
		
		// 多图文
		$('.addNews').click(function() {
			$("#wx_MassMsg").hide();
			$(".wx_MassMsg").show();
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
			$("#wx_MassMsg").show();
			$(".wx_MassMsg").hide();
		});
		
		
		
		
		// 无匹配回复文字
		$('.addNoneText').click(function() {
			ymPrompt.win({message:'/admin/weixin/weixin_SetSubcrib.asp?type=2&id=0',
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
							$('#wx_NoneReply').val(text);
						}
					}
				}
			});return false;
				
		});
		
		// 提交前判断
		$('#button').click(function() {
			var text = $("#wx_MassMsg").val();
			var counter = text.length;
			if(counter>600 || counter<1){
				alert("文字必须为1-600个字");
				return false;
			};
			Sends('SystemSet','/admin/weixin/weixin_MassSend.asp?Type=3',0,'',0,1,'MainRight','/admin/weixin/weixin_MassSend.asp?Type=1');
			
				
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
<style type="text/css">
.wx_MassMsg{width:500px;height:200px;border:solid 1px #ccc;}
.Button{background:none;cursor:pointer;}
.hover{background-color:#E8F4F7;}
</style>
<form id="SystemSet" name="SystemSet" method="post" action="Weixin_MassSend.asp?Type=3" onsubmit="return false;">

<div id="BoxTop" style="width:98%;"><span>添加图文</span></div>
<div id="BoxContents" style="width:98%;">
<table width="90%" border="0"  cellpadding="0" cellspacing="0" style="margin-top:15px;">
        <tr>
            <td align="right">群发对象</td>
            <td><select name="userType">
	<option value="-1">全部用户</option>
	<option value="0">未分组</option>
	<option value="1">黑名单</option>
	<option value="2">星标组</option>
	</select> &nbsp; 性别：<select name="userType">
	<option value="-1">全部用户</option>
	<option value="0">未分组</option>
	<option value="1">黑名单</option>
	<option value="2">星标组</option>
	</select></td>
        </tr>
        <tr>
            <td align="right">群发内容</td>
            <td><div class="bntType"><input type="button" value="文本" class="Button addText hover"> <input type="button" value="图片" class="Button addPic"> <input type="button" value="语音" class="Button addVoice"> <input type="button" value="视频" class="Button addVideo"> <input type="button" value="图文" class="Button addNews"></div><div class="wx_MassMsg" style="width:500px;height:200px;border:solid 1px #ccc;display:none;"></div><textarea name="wx_MassMsg" id="wx_MassMsg" style="width:500px;height:200px;border:solid 1px #ccc;padding:5px;"></textarea><p class="editor_tip">你还可以输入<span>600</span>个字符。</p> </td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:96%;">
	<input type="submit" class="Button" name="button" id="button" value="群 发" />
	<input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button1" value="关 闭" />
</div>
</form>
<%
End Sub
%>
<script language="jscript" runat="server">  
	Array.prototype.get = function(x) { return this[x]; };  
	function parseJSON(strJSON) { return eval("(" + strJSON + ")"); }  
</script>
<!--#Include File="../../Code.asp"-->