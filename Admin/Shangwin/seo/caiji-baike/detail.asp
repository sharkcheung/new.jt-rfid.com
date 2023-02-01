<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#Include File="../../loginchk.asp"-->
<!--#include file="c_func.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" oncontextmenu="return false;">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script src="jquery-1.4.4.js" type="text/javascript"></script>
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

$(window).resize(function(){
   $("#popup_show").css({
      position:'absolute',
      left: ($(window).width() - $("#popup_show").outerWidth())/2,
      top: ($(window).height() - $('#popup_show').outerHeight())/2 + $(document).scrollTop() 
   });
});
$(document).ready(function()
{
//初始化函数
$(window).resize();
	var card_container=$("#card-container").html();
	var baseInfoWrap=$(".baseInfoWrap").html();
	var lemma_main_content=$(".lemma-main-content").html();
	if(lemma_main_content==null){lemma_main_content=$(".main-content").html();}
	if (card_container==null){card_container=""}
	if (baseInfoWrap==null){baseInfoWrap=""}else{baseInfoWrap="<div class=\"baseInfoWrap\">"+baseInfoWrap+"</div>"}
	var html=card_container+baseInfoWrap+lemma_main_content;
	
	$(".top-tool,dl,.lemmaWgt-lemmaCatalog,.anchor-list,.album-list,#open-tag,.clear,.lemma-picture,.edit-prompt").remove();
	
	$("#main_s_content a").each(function(){
		//if($(this).attr("href")=='javascript:;')
		$(this).remove();
	})
	
	if(html=='null'){
		html=$(".main-content").html();
	}
	
	$(".cj_content").html(html);
	
   $("#caiji").hover(
      function(){
      $(this).addClass("hover");
   },function(){
      $(this).removeClass("hover");
   }
   );
   $("#caiji").click(function()
   {
      $("#caiji").removeClass("oso_red");
	  //alert($("#main_s_content h3").html());return false;
	  $("#main_s_content img").each(function(i){
	     if($(this).attr("src").indexOf("http://pic.wenwen.soso.com/")>-1){
	        $(this).remove();
		 }
	  });
      //$("#caiji").unbind("click");
      $("#caiji").html("正在采集，请稍等片刻...");
      $.ajax(
	  {
	  type:"POST",
	  url:"sosave.asp",
	  data:"tit="+escape($("#main_s_content h3").html())+"&cont="+escape($("#main_s_content .cj_content").html())+"&Module="+$("select[name=mdltype] option[selected]").val(),
	  async:false,
	  success:function(html)
	  {
         $("#popup_show").empty();
         $("#popup_show").hide();
	     $("#popup_show").append("<div id='ttop'><div class='ttopL'></div><div class='ttppt'>提示：</div><div class='ttopclose'></div><div class='ttopR'></div></div><div class='tcen'><p>"+html+"</p><span>确 定</span></div>");
		 $("#popup_show").show();
	  },
	  error:function(){
		alert("error");
	  }
      });
      //$("#caiji").bind("click");
      $("#caiji").html("一键采集到指定网站栏目！");
      $("#popup_show span").click(function(){
         $("#popup_show").empty();
         $("#popup_show").hide();
      })
   });
})
</script>
<link href="../css/caiji.css" rel="stylesheet" type="text/css" />
<style type="text/css">
#main_s_content .headline-2{text-align:left;font-size:16px;}

.baseInfoWrap .baseInfoLeft .biItem .biItemInner .biTitle, .baseInfoWrap .baseInfoRight .biItem .biItemInner .biTitle {
display: block;
float: left;
color: #999;
width: 78px;
padding: 0 5px 0 12px;
font-weight: bold;
}
.baseInfoWrap .baseInfoLeft .biItem .biItemInner, .baseInfoWrap .baseInfoRight .biItem .biItemInner {
zoom: 1;
}
.baseInfoWrap .baseInfoLeft .biItem, .baseInfoWrap .baseInfoRight .biItem {
zoom: 1;
line-height: 26px;
}
.baseInfoWrap {
font-size: 12px;
}
</style>
<title></title>
</head>
<body>
<div id="popup_show"></div>
<DIV style="DISPLAY: none" id=goTopBtn><IMG border=0 src="http://image001.dgcloud01.qebang.cn/caiji/to_top_blue.gif"  title="回到顶部" alt="回到顶部"></DIV>
<SCRIPT type=text/javascript>goTopEx();</SCRIPT>
<div class="main">
<%dim py,w,body,spi,sr,w8,qf,rn,qs,ch,act,tit,cont,body1,replaceStr%>
<div id="socaiji"><a class="returns" target="_self" title="返回列表 " onclick="javascript:history.go(-1);"> << 返 回</a><a id="caiji" title="点击采集 " style="color:#FF0000">一键采集到指定网站栏目！</a><%=selectlist%></div>

<script language="javascript" type="text/javascript" runat="server">　 
　function myEncodeURI(sStr){　 
　　　return encodeURI(sStr);　 
　}　 
</script> 
<%
dim t,para
t="0"
para=trim(Request.querystring("p"))
url=trim(Request.querystring("url"))
if t="0" then
	py = trim(Request.querystring("py"))
	w = trim(Request.querystring("w"))
	spi = trim(Request.querystring("spi"))
	sr = trim(Request.querystring("sr"))
	w8 = trim(Request.querystring("w8"))
	qf = trim(Request.querystring("qf"))
	rn = trim(Request.querystring("rn"))
	qs = trim(Request.querystring("qs"))
	ch = trim(Request.querystring("ch"))

	if left(url,1)="/" then
		url = "https://baike.baidu.com"&url
	end if
   'url="https://baike.baidu.com/"&para
	' if instr(url,"https")=0 then
		' url = replace(url,"http://","https://")
	' end if
	url=myEncodeURI(url)
	' url="https://baike.baidu.com/item/360%E5%BA%A6%E8%AF%84%E4%BC%B0/1712216"
   wstr=getHTTPPage(url,"utf-8")
   tit=strCut(wstr,"<title>","</title>",2)
   tit=trim(Replace(tit,"_百度百科",""))
   
	' Response.write url
	' Response.end
   	
		'body=strCut(wstr,"<body>","</body>",2)
		'body=strCut(body,"<div id=""content"" class=""col-main article"">","<div id=""side"" class=""col-sub"">",1)
		'body=replace(body,"<div id=""side"" class=""col-sub"">","")
		'body=strCut(body,"<div class=""clear""></div>","<div class=""clear""></div>",2)
   		
		'body=py_z_replace(wstr,"(<a[\s\S]*?</a>)*","")
		body=py_z_replace(wstr,"(<p[\s\S]*?><iframe[\s\S]*?</iframe></p>)*","")
		body=py_z_replace(body,"(<iframe[\s\S]*?</iframe>)*","")
		body=py_z_replace(body,"(<form[\s\S]*?</form>)*","")
		body=py_z_replace(body,"(<style[\s\S]*?</style>)*","")
		body=py_z_replace(body,"(<script[\s\S]*?</script>)*","")
		body=py_z_replace(body,"(<link[\s\S]*?</link>)*","")
  ' 		r = "<div.*?class=\""bk_title_body\"">(.|[\r\n])*?</span></div>(.|[\r\n])*?</div>"
  ' 		body = py_z_replace(body,r,"")
		
  ' 		r = "<dl.*?>(.|[\r\n])*?<\/dl>"
  ' 		body = py_z_replace(body,r,"")
		
  ' 		r = "<h1.*?>(.|[\r\n])*?<\/h1>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<h2.*?>(.|[\r\n])*?<\/h2>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<img.*?>(.|[\r\n])*?>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<p.*?></p>"
  ' 		body = py_z_replace(body,r,"")
		
  ' 		r = "^<div data-subindex=""0"".*?>(.|[\r\n])*?<span>目录</span>(.|[\r\n])*?</div></div>"
  ' 		body = py_z_replace(body,r,"")
		
  ' 		r = "</div> <span.*?>(.|[\r\n])*?<div class=""clear""></div>"
  ' 		body = py_z_replace(body,r,"")
		
  ' 		r = "(<script.*?>.*?</script>)*"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<div.*?data-subindex.*?>.*?<div id=""lemma-catalog-bottombg""><\/div>(.|[\r\n])*?<\/div>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<div.*?style=""clear:both;"">(.|[\r\n])*?</div>(.|[\r\n])*?</div>(.|[\r\n])*?</div>(.|[\r\n])*?</div>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<div.*?class=""z-catalog.*?"">(.|[\r\n])*?</div>(.|[\r\n])*?</div>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<div id=""bk-album-collection.*?"">(.|[\r\n])*?</div>(.|[\r\n])*?</div>"
  ' 		body = py_z_replace(body,r,"")
  ' 		r = "<div id=""lemmaExtend.*?"">(.|[\r\n])*?</div>(.|[\r\n])*?</div>"
  ' 		body = py_z_replace(body,r,"")
   		
		' if instr (body,"精华知识")>0 then
   			' replaceStr=strCut(body,"<h4 class=""ico_official_answer"">","</h4>",1)
		' else
   			' replaceStr=strCut(body,"<h4 class=""ico_star_answer"">","</h4>",1)
		' end if
		' body=replace(body,replaceStr,"")
   		' ' replaceStr=strCut(body,"<!--","-->",1)
		body=replace(body,"</div><div class=""clear""></div></div></div></div>","")
		'body=replace(body,"<div class=""evaluation_wrap"">","")
		
		
		if body="" then 
			body="<div style='font-size:14;color:red;'>该问答暂无采纳答案！</div>"
		else
		
			body="<div class=""hideHtml"" style=""display:none;"">"&body&"</div>"
			'body="<div class=""hideHtml"" style=""display:none;"">"&body&"</div>"
			'body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'>"& body &"</pre>"
		end if
	'if instr(body,"<pre>")>0 and instr(body,"<pre>") < instr(body,"</pre>") then
		'body=strCut(body,"<pre>","</pre>",2)
	'else
	'end if
   
elseif t="1" then
	
	
   url="http://www.baike.com/wiki/"&para

   wstr=getHTTPPage(url,"utf-8")
   body=strCut(wstr,"<body>","</body>",1)
   response.write body
   response.end
   body=strCut(body,"<div class=""l w-640"">","<div class=""r w-320"">",2)
   tit=strCut(body,"<div class=""hd"">","</h2>",1)
   tit=strCut(tit,"<h2 >","</h2>",2)	
   	body=strCut(body,"<div class=""mod-best-a","<div class=""mod-btns-answer-best",2)
	body=strCut(body,"<div class=""qa-content"">","</div>",2)
	if body="" then body="<div style='font-size:14;color:red;'>该问答暂无采纳答案！</div>"
	body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'></pre>"
   
elseif t="2" then
	
	url="http://zh.wikipedia.org/zh-cn/"&para

	wstr=getHTTPPage(url,"utf-8")
	tit=strCut(wstr,"<span dir=""auto"">","</span>",2)
	body=strCut(wstr,"<div id=""bodyContent"" class=""mw-body-content"">","<div id=""mw-navigation"">",1)
	body=replace(body,"<div id=""mw-navigation"">","")
	r = "<a.*?>.*?<img.*?/></a>"
	body=py_z_replace(body,r,"")
	r = "<img.*?>"
	body=py_z_replace(body,r,"")
	r = "<div.*?></div>"
	body=py_z_replace(body,r,"")
	r = "<div.*?>[\r\n]*?</div>"
	body=py_z_replace(body,r,"")
	r = "<div id=""siteSub"">.*?</div>"
	body=py_z_replace(body,r,"")
	r = "<div id=""jump-to-nav"" class=""mw-jump"">(.|[\r\n])*?</div>"
	body=py_z_replace(body,r,"")
	r = "<span.*?>\[</span>(.|[\r\n])*?<a.*?>编辑</a>(.|[\r\n])*?</span>"
	body=py_z_replace(body,r,"")
	r = "<div id=""toctitle"">(.|[\r\n])*?</div>"
	body=py_z_replace(body,r,"")
	r = "<div id=""toc"" class=""toc"">(.|[\r\n])*?</div>"
	body=py_z_replace(body,r,"")
	r = "<div class=""references-small"">(.|[\r\n])*?</div>"
	body=py_z_replace(body,r,"")
	r = "<table.*?>(.|[\r\n])*?</table>"
	body=py_z_replace(body,r,"")
	r = "<div class=""printfooter"">(.|[\r\n])*?</div>"
	body=py_z_replace(body,r,"")
	r = "<a class=""printfooter"">(.|[\r\n])*?</div>"
	body=py_z_replace(body,r,"")
	r = "<ol class=""references"">(.|[\r\n])*?</ol>"
	body=py_z_replace(body,r,"")
	body=LoseATag(body)
	if body="" then
		body=strCut(wstr,"<pre id=""recommend-content","</pre>",1)
		r = "(<pre id=""recommend-content.*?>)*"
		body=py_z_replace(body,r,"")
	end if
	' response.write server.htmlencode(body)
	' response.end
	if body="" then body="<div style='font-size:14;color:red;'>该问答暂无采纳答案！</div>"
	'body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'>"& body 
	
end if
response.write "<div id=""main_s_content"" style='padding-top:50px;'><h3 style='font-size:15pt;text-align:center;'>"&tit&"</h3><div class=""cj_content"">加载中..."&body&"</div></div>"
act=trim(request("act"))


Function LoseATag(ContentStr)
Dim ClsTempLoseStr,regEx
ClsTempLoseStr = Cstr(ContentStr)
Set regEx = New RegExp
regEx.Pattern = "<(\/){0,1}a[^<>]*>"
regEx.IgnoreCase = True
regEx.Global = True
ClsTempLoseStr = regEx.Replace(ClsTempLoseStr,"")
LoseATag = ClsTempLoseStr
Set regEx = Nothing
End Function
%>

</div>
</body>
</html>
