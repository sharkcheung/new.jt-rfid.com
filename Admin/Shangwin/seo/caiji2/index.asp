<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../loginchk.asp"-->
<!--#include file="c_func.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false;">
<head>
<meta content="IE=EmulateIE7" http-equiv="X-UA-Compatible" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<!--
<link href="http://cache.soso.com/wenwen/css/w_base_201012271648.css" media="screen" rel="stylesheet" type="text/css" />
<link href="http://cache.soso.com/wenwen/css/search_201011241014.css" media="screen" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="http://cache.soso.com/30d/css/web/isoso9.css" />
-->
<script src="jquery-1.4.4.js" type="text/javascript"></script>
<script type="text/javascript">
$(document).ready(function(){
	//alert($(".qa-list").html());
	//$("#qaresult").html("<ul class=\"qa-list\">"+$(".qa-list").html()+"</ul>");
	$(".aside,#e_idea_wenda_leftBox").remove();
	$(window).scroll(function(){
		   $("#popup_show").css({
		      position:'absolute',
		      left: ($(window).width() - $("#popup_show").outerWidth())/2,
      			top: ($(window).height() - $('#popup_show').outerHeight())/2 + $(document).scrollTop() 
		   });
    });
	$("a").click(function(){
         $("#popup_show").empty();
         $("#popup_show").hide();
	     $("#popup_show").append("<div class='tcen'><p>页面加载中，请稍候...</p></div>");
		 $("#popup_show").show();
		   $("#popup_show").css({
		      position:'absolute',
		      left: ($(window).width() - $("#popup_show").outerWidth())/2,
      			top: ($(window).height() - $('#popup_show').outerHeight())/2 + $(document).scrollTop() 
		   });
		
	})
})
function goTopEx(){
        var obj=document.getElementById("goTopBtn");
        function getScrollTop(){
                return document.documentElement.scrollTop;
            }
        function setScrollTop(value){
                document.documentElement.scrollTop=value;
            }    
        window.onscroll=function(){getScrollTop()>0?obj.style.display="":obj.style.display="none";}
        obj.onclick=function(){
            var goTop=setInterval(scrollMove,10);
            function scrollMove(){
                    setScrollTop(getScrollTop()/1.1);
                    if(getScrollTop()<1)clearInterval(goTop);
                }
        }
    }

   $(document).ready(function(){
      $("ol li:first-child").addClass("hide");
	  $("ol li").find("h3");
//	  $("#search_kcss a").hover(function(){
//	     $(this).addClass("hover");
//	  },function(){
//	     $(this).removeClass("hover");
//	  })
   })
</script>
<style type="text/css">

</style>
<link href="../css/caiji.css" rel="stylesheet" type="text/css" />
<title></title>
<style type="text/css" media="screen" id="test">
		#popup_show .tcen{
	background:none ;
	background-color:#fff;
	padding-top:20px; padding-bottom:10px;
	border:1px #31699C solid; 
}
</style>
</head>
<body>
<div id="popup_show"></div>
<DIV style="DISPLAY: none" id=goTopBtn><IMG border=0 src="http://image001.dgcloud01.qebang.cn/caiji/to_top_blue.gif"  title="回到顶部" alt="回到顶部"></DIV>
<SCRIPT type=text/javascript>goTopEx();</SCRIPT>
<%dim keyword,r,body,pt,pagecont
   pt=clng(Request.querystring("pt"))
   if pt=0 then pt=1
   if pt=2 then
		response.Charset="gb2312"
		keyword = trim(Request.querystring("sp"))
		response.Charset="utf-8"
	else
		keyword = trim(Request.querystring("sp"))
	end if
%>
<div id="c_content">
  <form id="frm_search" method="get" name="frm_search">
    请您输入您需要运营的关键词：
    <input id="sp" name="sp" type="text" value="<%=keyword%>" />
    &nbsp;
    <select name="pt">
      <!--option value="0" <%if pt=0 then response.write "selected"%>>搜搜问问</option-->
      <option value="1" <%if pt=1 then response.write "selected"%>>360问答</option>
      <option value="2" <%if pt=2 then response.write "selected"%>>百度知道</option>
    </select>
    <input type="submit" value=" 问答类运营 " onclick="this.disabled=true;this.value='正在搜索中...';frm_search.submit();"/>
    <a href="../caiji3/?keywords=<%=keyword%>">转关键词运营</a> | 
    <a href="../caiji-baike/?sp=<%=keyword%>">转行业新闻运营</a>
  </form>
  <!--<form name="frm_search" id="frm_search" method="get"><input type="text" name="w" id="w" /> &nbsp;<input type="submit" value="搜索"/></form>-->
</div>
<%
 	cookies_key=request.Cookies("sosearch_keyword")
	if keyword<>"" then
		cookies_key=trim(request.Cookies("sosearch_keyword"))
   		
   			add="<a href='?sp="&keyword&"&ch=w.search.sb'>"&keyword&"</a> "
			if instr(Request.Cookies("sosearch_keyword"),add)>0 then
			else
				Response.Cookies("sosearch_keyword")=add&Request.Cookies("sosearch_keyword")
				Response.Cookies("sosearch_keyword").Expires=#May 10,2020#
			end if
   			if cookies_key="" then 
     		 	response.Cookies("sosearch_keyword")="<a href='?sp="&keyword&"&ch=w.search.sb'>"&keyword&"</a> "
     		end if
  
	end if
	a = trim(Request.querystring("a"))
		if a="c" then 
		   response.Cookies("sosearch_keyword")=""
		   response.redirct "./"
		end if
%>
<div id="search_kcss">您使用过的关键词：<%=Request.Cookies("sosearch_keyword")%>[<a href="?a=c" title="点击清空关键词">清空</a>]</div>
<%
	if keyword<>"" then
   			
  
   
   cur_page=clng(Request.querystring("pg"))
   
   if cur_page="" or cur_page=0 then
   		if pt=0 then
   			url="http://wenwen.sogou.com/s/?sp=S"&server.URLEncode(keyword)&"&ch=w.search.sb&w="&server.URLEncode(keyword)&"&search=搜索答案"
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
   			url="http://wenwen.sogou.com/z/Search.e?sp=S"&server.URLEncode(keyword)&"&sci=0&pg="&cur_page&""
		elseif pt=1 then
   			url="http://wenda.haosou.com/search/?q="&server.URLEncode(keyword)&"&pn="&cur_page
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
   			wstr=getHTTPPage(url,"utf-8")
			'response.write url
			' response.write server.htmlencode(wstr)
			'response.end
   			body=strCut(wstr,"<div class=""result"" id=""result"">","<div class=""result_side"" id=""result_side"">",1)
			body=replace(body,"<div class=""result_side"" id=""result_side"">","")
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
   			wstr=getHTTPPage(url,"utf-8")
			if instr(wstr,"id=""noanswer""")>0 then
		   		body=""
			else
				pagecont=strCut(wstr,"<div class=""pagination"" id=""qaresult-page"">","</div>",1)
				pagecont = replace(pagecont,"?q=","?pt=1&sp=")
				pagecont = replace(pagecont,"pn","pg")
				body=strCut(wstr,"<div id=""qaresult"" clsss=""clearfix"">","<div class=""clearfix""",2)
				r = "(<a *? href=""/z/.*?.htm"">.*?</a>)*"
				body=py_z_replace(body,"<div class=""qa-i-ft"".*?</div>","")
				body = replace(body,"/q/",ServerPath&"detail.asp?t=1&p=/q/")
				body = replace(body,"<div class=""clearfix"">","")
				body=body&pagecont
				body = replace(body,"_blank","_self")
			end if 
		case 2
   			wstr=getHTTPPage(url,"gb2312")
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
	body=py_z_replace(body,"(<form[\s\S]*?</form>)*","")
	body=py_z_replace(body,"(<style[\s\S]*?</style>)*","")
	body=py_z_replace(body,"(<script[\s\S]*?<\/script>)*","")
	body=py_z_replace(body,"(<img[\s\S]*?>)*","")
	if body="" then body="<div style='color:red;font-size:14px;'>当前搜索关键词暂时没有可采集内容，请更换关键词<div>"
	response.write "<div id=""main_s_content"">"&body&"</div>"

end if
%>
</body>
</html>
