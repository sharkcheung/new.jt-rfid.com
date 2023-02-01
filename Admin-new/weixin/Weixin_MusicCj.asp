<!--#Include File="../AdminCheck.asp"-->
<link href="/admin/Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="/Js/function.js"></script>
<%
'==========================================
'文 件 名：Weixin_Menu.asp
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
Dim Fk_menuName,Fk_menuType,Fk_menuEvent,Fk_menuStatus,Fk_menuPx,Fk_menuParent

'获取参数
id=Clng(Request.QueryString("id"))


	Session("NowPage")=FkFun.GetNowUrl()
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	
	dim keyword,r,body,pt,pagecont,cur_page,url,wstr
   pt=clng(Request.querystring("pt"))
   if pt=2 then
		response.Charset="gb2312"
		keyword = trim(Request.querystring("sp"))
		response.Charset="utf-8"
	else
		keyword = trim(Request.querystring("sp"))
	end if
%>
<script type="text/javascript">
 function closeWin()
   {
    window.parent.ymPrompt.doHandler("error",true);
   }
</script>
<style type="text/css">
body{padding:4px;}
#ListTop,#ListContent{width:100%;margin:0 auto}
#ListContent table{border-right:0}
#ListContent table td{line-height:34px;}
</style>
<div id="ListTop">
  <form id="frm_search" method="get" name="frm_search">
    输入要采集信息的关键词：
    <input id="sp" name="sp" type="text" value="<%=keyword%>" />
    &nbsp;
    <select name="pt">
      <option value="0" <%if pt=0 then response.write "selected"%>>搜搜问问</option>
      <option value="1" <%if pt=1 then response.write "selected"%>>360问答</option>
      <option value="2" <%if pt=2 then response.write "selected"%>>百度知道</option>
    </select>
    <input type="submit" value=" 语音采集 " onclick="this.disabled=true;this.value='正在搜索中...';frm_search.submit();"/>
  </form>
    
</div>
<div id="ListContent">
<%
if keyword<>"" then
   			
  
   
   cur_page=clng(Request.querystring("pg"))
   
   if cur_page="" or cur_page=0 then
   		if pt=0 then
   			url="http://so.111ttt.com/cse/search?s=10087588629572173360&nsid=1&q="&server.URLEncode(keyword)
		elseif pt=1 then
   			url="http://wenda.so.com/search/?q="&server.URLEncode(keyword)
		elseif pt=2 then
			response.Charset="gb2312"
   			url="http://zhidao.baidu.com/search?lm=0&rn=10&pn=0&fr=search&ie=utf8&word="&server.URLEncode(keyword)&""
			response.Charset="utf-8"
		else
			response.end
		end if 
   'url="http://www.soso.com/q?pid=s.idx&w="&keyword
   else
   		if pt=0 then
   			url="http://wenwen.soso.com/z/Search.e?sp=S"&server.URLEncode(keyword)&"&sci=0&pg="&cur_page&""
		elseif pt=1 then
   			url="http://wenda.so.com/search/?q="&server.URLEncode(keyword)&"&pn="&cur_page
		elseif pt=2 then
			response.Charset="gb2312"
   			url="http://zhidao.baidu.com/search?&ie=gbk&word="&server.URLEncode(keyword)&"&site=-1&sites=0&date=0&pn="&cur_page
			response.Charset="utf-8"
		else
			response.end
		end if
   end if
   
   
   select case pt
  		case 0			
   			wstr=FkFun.GetHttpPage(url,"utf-8")
   			body=strCut(wstr,"<div id=""results"" class=""content-main"">","<div class=""extra"">",2)
			
			response.write body
			response.end
			dim pagination
			pagination=strCut(wstr,"<div class=""pagination"">","</div>",1)
			
   			body=strCut(body,"<ol class=""result_list"">","</ol>",1)
			
   			r = "(<div class=""question_info"">[\s\S]*?</div>)*"
   			body = py_z_replace(body,r,"")
			
   			r = "(<div class=""info"">[\s\S]*?</div>)*"
   			body = py_z_replace(body,r,"")
			
   			r = "(<ul class=""baike_result""[\s\S]*?</ul>)*"
   			body = py_z_replace(body,r,"")
			
   			'r = "(<div class=""baike""[\s\S]*?</div>)*"
   			'body = py_z_replace(body,r,"")
			
   			r = "(<li[\s\S]*?<h3 ch=""w\.search\.baike[\s\S]*?</li>)*"
   			body = py_z_replace(body,r,"")
			
			
			  ' response.Write server.htmlencode(body)
			  ' response.end
			
   			'body=replace(body,"<span class=""yl1"" onFocus=""blur();"">预览</span>","")
   			r = "(<span class=""solved_time"">.*?</span>)*"
   			body = py_z_replace(body,r,"")
   			'body=replace(body,"<span>-</span>","")
   			r = "(<a *? href=""/z/.*?.htm"">.*?</a>)*"
   			body = py_z_replace(body,r,"")
   			'r = "(<a target=""_blank"" .*?>.*?</a>.*?)*"
   			'body = py_z_replace(body,r,"")
			body=body&pagination
   			body = replace(body,"/z/Search.e?sp=S",CUrl(1)&"?sp=")
   			body = replace(body,".htm?w=",".htm&w=")
   			body = replace(body,"/z/q",ServerPath&"detail.asp?t=0&p=/z/q")
   			'body = replace(body,"<script>setTimeout(setPosition,50);</script>","")
   			body = replace(body,"_blank","_self")
   			'body=strCut(body,"<ol class=""result_list"">","</ol>",1)
   			'body = replace(body,"</div>","")

			'   r = "(<cite>.*?</cite>-)*"
			'   body = py_z_replace(body,r,"")
			'   r="(<div class=""url""><a href=.*?>网页快照</a></div>)*"
			'   body = py_z_replace(body,r,"")
			'   r="(<div class=""result_summary"">-<span class=""preview"" id="".*?""></span></div>)*"
			'   body = py_z_replace(body,r,"")
   			'response.write "<div id=""main_s_content"">"&body&turn_page&"</div>"
		case 1
   			wstr=FkFun.GetHttpPage(url,"utf-8")
			if instr(wstr,"id=""noanswer""")>0 then
		   		body=""
			else
				pagecont=strCut(wstr,"<div class=""pagination"" id=""qaresult-page"">","</div>",1)
				pagecont = replace(pagecont,"?q=","?pt=1&sp=")
				pagecont = replace(pagecont,"pn","pg")
				body=strCut(wstr,"<div id=""qaresult"" clsss=""clearfix"">","<div class=""clearfix"">",1)
				r = "(<a *? href=""/z/.*?.htm"">.*?</a>)*"
				body=py_z_replace(body,"<div class=""qa-i-ft"".*?</div>","")
				body = replace(body,"/q/",ServerPath&"detail.asp?t=1&p=/q/")
				body = replace(body,"<div class=""clearfix"">","")
				body=body&pagecont
				body = replace(body,"_blank","_self")
			end if 
		case 2
   			wstr=FkFun.GetHttpPage(url,"gb2312")
			if instr(wstr,"id=""noanswer""")>0 then
		   		body="该问题暂无答案!"
			else
				pagecont=strCut(wstr,"<div class=""pager"" alog-alias=""pager"">","</div>",1)
				response.Charset="gb2312"
				pagecont = replace(pagecont,"/search?","?pt=2&sp="&server.URLEncode(keyword)&"&")
				response.Charset="utf-8"
				pagecont = replace(pagecont,"pn","pg")
				body=strCut(wstr,"<div class=""list"" id=""wgt-list"" data-log-area=""list"">","<div class=""list-footer"">",1)
				body=strCut(wstr,"<div class=""list"" id=""wgt-list"" data-log-area=""list"">","</div>",1)
				r = "(<span class=""f-12 f-light grid-r"">(.|[\r\n])*?</span>)*"
				body=py_z_replace(body,r,"")
				r = "(<dd class=""dd explain f-light""(.|[\r\n])*?</dd>)*"
				body=py_z_replace(body,r,"")
				r = "(<i class=""i-answer-text"">(.|[\r\n])*?</i>)*"
				body=py_z_replace(body,r,"")
				body = replace(body,"http://zhidao.baidu.com/",ServerPath&"detail.asp?t=2&p=")
				body = replace(body,"_blank","_self")
				body = replace(body,"<div class=""clearfix"">","")
				body=body&pagecont
			end if 
			
	end select
	if body="" then body="<div style='color:red;font-size:14px;'>当前搜索关键词暂时没有可采集内容，请更换关键词<div>"
	response.write "<div id=""main_s_content"">"&body&"</div>"
   
end if


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


%>
</div>
<%
%><!--#Include File="../../Code.asp"-->