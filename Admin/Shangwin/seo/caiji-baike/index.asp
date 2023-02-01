<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../loginchk.asp"-->
<!--#include file="c_func.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false;">
<head>
<meta content="IE=EmulateIE7" http-equiv="X-UA-Compatible" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<script src="jquery-1.4.4.js" type="text/javascript"></script>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script type="text/javascript">
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
	   
		var mod_list=$(".mod-list");
		if(mod_list.html()==null){
			mod_list=$(".search-list");
		}
		mod_list.find("a").each(function(){
			var url=$(this).attr("href");
			$(this).attr("href","/admin/shangwin/seo/caiji-baike/detail.asp?url="+url)
		})
		var mod_list=mod_list.html();
		$(".tangram-pager").find("a").each(function(){
			var url=$(this).attr("href");
			var ind=url.indexOf("#");
			var url=url.substring(ind+1);
			$(this).attr("href","/admin/shangwin/seo/caiji-baike/?sp=<%=Request("sp")%>&pg="+url)
		})
		var tangram_pager=$(".tangram-pager").html();
		//$(".hideHtml").replace("/<s"+"cript.*?>.*?<\/scr"+"ipt>/ig", "");
		re = /_百度百科/g;
		mod_list=mod_list.replace(re,"")
		re = /<br>/g;
		mod_list=mod_list.replace(re,"");
		
		if (mod_list==null){mod_list=""}else{mod_list="<div class=\"qa-list\"><dl>"+mod_list+"</dl></div>"}
		if (tangram_pager==null){tangram_pager=""}else{tangram_pager="<div class=\"pages\">"+tangram_pager+"</div>"}
	   $("#main_s_content").html(mod_list+tangram_pager);
	   $(".fs").remove();
	  // alert($("#main_s_content").html().indexOf("_百度百科"));
	   //$("#main_s_content").html($("#main_s_content").html())
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
<style type="text/css">
		#popup_show .tcen{
	background:none ;
	background-color:#fff;
	padding-top:20px; padding-bottom:10px;
	border:1px #31699C solid; 
}
 .f em{ 
color: #FF0000;
font-weight: bold;
}

#main_s_content .result-list h3 {
text-align:left;
}

#page {
    overflow: hidden;
}
.r {
    float: right;
}
.pages {
    clear: none;
    margin-left: -3px;
    padding: 0;
    text-align: left;
}
.pages {
    clear: both;
    height: 29px;
    padding: 10px;
}
.pages span {
    background: none repeat scroll 0 0 #FBFBFB;
    border: 1px solid #E7E7E7;
    color: #191919;
    cursor: default;
    display: inline-block;
    height: 21px;
    line-height: 21px;
    margin: 0 3px;
    padding: 3px 10px;
    text-align: center;
    vertical-align: middle;
}
.pages a{
    background-color: #F5FDFF;
    border: 1px solid #D2F0FB;
    color: #2376CB;
    display: inline-block;
    height: 21px;
    line-height: 21px;
    margin: 0 3px;
    padding: 3px 10px;
    text-align: center;
    vertical-align: middle;
}
.mw-search-result-heading a{color:#f00;font-weight:bolder;}
</style>
<title></title>
</head>
<body>
<div id="popup_show"></div>
<DIV style="DISPLAY: none" id=goTopBtn><IMG border=0 src="http://image001.dgcloud01.qebang.cn/caiji/to_top_blue.gif"  title="回到顶部" alt="回到顶部"></DIV>
<SCRIPT type=text/javascript>goTopEx();</SCRIPT>
<%

dim keyword,r,body,pt,pagecont
   pt=clng(Request.querystring("pt"))
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
      <option value="0" <%if pt=0 then response.write "selected"%>>百度百科</option>
      <!--option value="1" <%if pt=1 then response.write "selected"%>>互动百科</option-->
      <!--option value="2" <%if pt=2 then response.write "selected"%>>维基百科</option-->
      <!--option value="3" <%if pt=3 then response.write "selected"%>>MBA智库百科</option-->
    </select>
    <input type="submit" value=" 行业新闻运营 " onclick="this.disabled=true;this.value='正在搜索中...';frm_search.submit();"/>
    <a href="../caiji2/?sp=<%=keyword%>">转问答类运营</a> | 
    <a href="../caiji3/?keywords=<%=keyword%>">转关键词运营</a>
  </form>
  <!--<form name="frm_search" id="frm_search" method="get"><input type="text" name="w" id="w" /> &nbsp;<input type="submit" value="搜索"/></form>-->
</div>
<%
response.charset="utf-8"
session.codepage=65001
 	cookies_key=request.Cookies("bksearch_keyword")
	if keyword<>"" then
		cookies_key=trim(request.Cookies("bksearch_keyword"))
   		
   			add="<a href='?sp="&keyword&"&ch=w.search.sb'>"&keyword&"</a> "
			if instr(Request.Cookies("bksearch_keyword"),add)>0 then
			else
				Response.Cookies("bksearch_keyword")=add&Request.Cookies("bksearch_keyword")
				Response.Cookies("bksearch_keyword").Expires=#May 10,2050#
			end if
   			if cookies_key="" then 
     		 	response.Cookies("bksearch_keyword")="<a href='?sp="&keyword&"&ch=w.search.sb'>"&keyword&"</a> "
     		end if
  
	end if
	a = trim(Request.querystring("a"))
		if a="c" then 
		   response.Cookies("bksearch_keyword")=""
		   response.redirct "./"
		end if
%>
<div id="search_kcss">您使用过的关键词：<%=Request.Cookies("bksearch_keyword")%>[<a href="?a=c" title="点击清空关键词">清空</a>]</div>
<%
	if keyword<>"" then   			
   
   cur_page=clng(Request.querystring("pg"))
   
   if cur_page="" or cur_page=0 then
   		if pt=0 then
			'response.Charset="gb2312"
   			url="https://baike.baidu.com/search?word="&server.URLEncode(keyword)&"&rn=0&pn=0&enc=utf8"
			'response.Charset="utf-8"
		elseif pt=1 then
			url="http://so.baike.com/s/doc/"&server.URLEncode(keyword)&"&prd=button_doc_search"
		elseif pt=2 then
			url="http://zh.wikipedia.org/w/index.php?title=Special%3A%E6%90%9C%E7%B4%A2&limit=20&profile=default&search="&server.URLEncode(keyword)&"&fulltext=Search"
		else
			response.end
		end if 
   'url="http://www.soso.com/q?pid=s.idx&w="&keyword
   else
   		if pt=0 then
   			url="https://baike.baidu.com/search?word="&server.URLEncode(keyword)&"&pn="&(cur_page-1)*10
		elseif pt=1 then
			url="http://so.baike.com/s/"&server.URLEncode(keyword)&"/doc/"&cur_page
		elseif pt=2 then
			url="http://zh.wikipedia.org/w/index.php?title=Special:搜索&limit=20&offset="&cur_page&"&profile=default&search="&server.URLEncode(keyword)&""
		else
			response.end
		end if
   end if
   select case pt
  		case 0			
   			wstr=getHTTPPage(url,"utf-8")
   			'body=py_z_replace(wstr,"<div id=""pagerBox"">[\s\S]*?</div>","")
			'pagination=strCut(body,"<div id=""tangram-pager--pager""","</div>",1)
   			'Response.write  instr(body,"tangram-pager-main")'pagination
			'Response.end
   			'body=strCut(wstr,"<div class=""mod-list"">","<div class=""pager clearfix"">",1)
			'Response.write body&pagination
			'Response.End
   			'r = "(<div class=""pager clearfix"">>)"
   			'body = py_z_replace(body,r,"")
			
			
   			'r = "(<div class=""baike""[\s\S]*?</div>)*"
   			' body = py_z_replace(body,r,"")
			' response.write wstr
			' response.end
   			r = "(<br/>)"
   			body = py_z_replace(wstr,r,"")
			' body = replace(body,"http://baike","https://baike")
			
   			'r = "(<div class=""fs""(.|[\r\n])*?</div>)"
   			'body = py_z_replace(body,r,"")
			
   			'body=replace(body,"<span class=""yl1"" onFocus=""blur();"">预览</span>","")
   			r = "(<span class=""solved_time"">.*?</span>)*"
   			body = py_z_replace(body,r,"")
   			r = "(<span class=""result-date"">.*?</span>)*"
   			body = py_z_replace(body,r,"")
   			'r = "(<a *? href=""/z/.*?.htm"">.*?</a>)*"
   			'body = py_z_replace(body,r,"")
   			'r = "(<a target=""_blank"" .*?>.*?</a>.*?)*"
   			'body = py_z_replace(body,r,"")
   			body=py_z_replace(body,"(_百度百科)*","")
   			'body = replace(body,"http://baike.baidu.com/",ServerPath&"detail.asp?t=0&p=")
   			'body = replace(body,".htm?w=",".htm&w=")
   			'body = replace(body,"/z/q",ServerPath&"detail.asp?t=0&p=/z/q")
   			'body = replace(body,"<script>setTimeout(setPosition,50);</script>","")
   			body = replace(body,"_blank","_self")
			'body=py_z_replace(body,"(<script[\s\S]*?</script>)*","")'替换<javascript>调用
   			'body=body&pagination
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
			if instr(wstr,"<div class=""w-630"">")<0 then
		   		body=""

			else
				body=strCut(wstr,"<div class=""w-630"">","<div class=""w-300 r"" style=""margin-top:0;"">",1)
				body = replace(body,"<div class=""w-300 r"" style=""margin-top:0;"">","")
				r = "<form.*?>(.|[\r\n])*?</form>"
				body = py_z_replace(body,r,"")
				r = "<dl.*?class=""clearfix.*?>(.|[\r\n])*?</dl>"
				body = py_z_replace(body,r,"")
				r = "<a.*?>.*?<img.*?/></a>"
				body = py_z_replace(body,r,"")
				r = "<p.*?class=""gray80.*?>(.|[\r\n])*?</p>"
				body = py_z_replace(body,r,"")
				r = "<span class="".*?gray80 r"">.*?结果.*?个</span>"
				body = py_z_replace(body,r,"")
				r = "<div class=""result-list"">[\r\n]*?<h3><a.*?>.*?_百科图片(.|[\r\n])*?<ul class=""img_results"">(.|[\r\n])*?<div class=""clearfix""></div>(.|[\r\n])*?</div>"
				body = py_z_replace(body,r,"")
				' pagecont = replace(pagecont,"?q=","?pt=1&sp=")
				' pagecont = replace(pagecont,"pn","pg")
				' body=strCut(wstr,"<div id=""qaresult"" clsss=""clearfix"">","<div class=""clearfix"">",1)
				' r = "(<a *? href=""/z/.*?.htm"">.*?</a>)*"
				' body=py_z_replace(body,"<div class=""qa-i-ft"".*?</div>","")
				' body = replace(body,"/q/",ServerPath&"detail.asp?t=1&p=/q/")
				' body = replace(body,"<div class=""clearfix"">","")
				' body=body&pagecont
				wstr = replace(wstr,"_blank","_self")
				body = replace(body,"http://www.baike.com/wiki/",ServerPath&"detail.asp?t=1&p=")
				body = replace(body,"http://so.baike.com/s/"&server.URLEncode(keyword)&"/doc/",ServerPath&"?sp="&server.URLEncode(keyword)&"&pt=1&pg=")
			end if 
		case 2
   			wstr=getHTTPPage(url,"utf-8")
			if instr(wstr,"id=""noanswer""")>0 then
		   		body="该问题暂无答案!"
			else
				pagecont=strCut(wstr,"<p class='mw-search-pager-bottom'>","</p>",1)
				body=strCut(wstr,"<ul class='mw-search-results'>","</ul>",1)
				r = "(<p class=""mw-search-createlink"">(.|[\r\n])*?</p>)*"
				body=py_z_replace(body,r,"")
				' response.write body
				r = "(<p class=""mw-search-createlink"">(.|[\r\n])*?</p>)*"
				body=py_z_replace(body,r,"")
				r = "(<p class=""mw-search-createlink"">(.|[\r\n])*?</p>)*"
				pagecont=replace(pagecont,"/w/index.php?title=Special:%E6%90%9C%E7%B4%A2&amp;limit=20&amp;offset=","?pt=2&sp="&server.URLEncode(keyword)&"&pg=")
				body=replace(body,"/wiki/","detail.asp?t=2&p=")
				'response.write server.htmlencode(pagecont)
				r = "<div class='mw-search-result-data'>(.|[\r\n])*?<\/div>"
				body=py_z_replace(body,r,"")
				' response.end
				' r = "(<dd class=""dd explain f-light""(.|[\r\n])*?</dd>)*"
				' body=py_z_replace(body,r,"")
				' r = "(<i class=""i-answer-text"">(.|[\r\n])*?</i>)*"
				' body=py_z_replace(body,r,"")
				' body = replace(body,"http://zhidao.baidu.com/",ServerPath&"detail.asp?t=2&p=")
				' body = replace(body,"_blank","_self")
				' body = replace(body,"<div class=""clearfix"">","")
				body=body&pagecont
			end if 
			
	end select
	if body="" then body="<div style='color:red;font-size:14px;'>当前搜索关键词暂时没有可采集内容，请更换关键词<div>"
	response.write "<div id=""main_s_content"">加载中，请稍候...<div class=""hideHtml"" style=""display:none;"">"&body&"</div></div>"

end if
%>
</body>
</html>
