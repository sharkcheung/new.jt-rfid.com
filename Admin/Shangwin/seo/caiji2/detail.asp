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
      $("#caiji").unbind("click");
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
      $("#caiji").bind("click");
      $("#caiji").html("一键采集到指定网站栏目！");
      $("#popup_show span").click(function(){
         $("#popup_show").empty();
         $("#popup_show").hide();
      })
   });
})
</script>
<link href="../css/caiji.css" rel="stylesheet" type="text/css" />
<title></title>
</head>
<body>
<DIV style="DISPLAY: none" id=goTopBtn><IMG border=0 src="http://image001.dgcloud01.qebang.cn/caiji/to_top_blue.gif"  title="回到顶部" alt="回到顶部"></DIV>
<SCRIPT type=text/javascript>goTopEx();</SCRIPT>
<div class="main">
<%dim py,w,body,spi,sr,w8,qf,rn,qs,ch,act,tit,cont,body1,replaceStr%>
<div id="socaiji"><a class="returns" target="_self" title="返回列表 " onclick="javascript:history.go(-1);"> << 返 回</a><a id="caiji" title="点击采集 " style="color:#FF0000">一键采集到指定网站栏目！</a><%=selectlist%></div>
<%
dim t,para
t=trim(Request.querystring("t"))
para=trim(Request.querystring("p"))
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


   url="http://wenwen.soso.com"&para

   wstr=getHTTPPage(url,"utf-8")
   tit=strCut(wstr,"<h4 id=""questionTitle"">","</h4>",2)
   if tit="" then
		tit=strCut(wstr,"<h3 id=""questionTitle"">","</h3>",2)
	end if
   if instr(wstr,"<div class=""resolved_question"">")>0 then		
		'body=strCut(wstr,"<div class=""resolved_question"">","<div class=""sign_wrap satisfaction_vote"">",1)
   		body=strCut(wstr,"<div class=""answer_con"">","</div>",2)
   		' r = "(<!--[\s\S]*)-->"
   		' body = py_z_replace(body,r,"")
		
		' if instr (body,"精华知识")>0 then
   			' replaceStr=strCut(body,"<h4 class=""ico_official_answer"">","</h4>",1)
		' else
   			' replaceStr=strCut(body,"<h4 class=""ico_star_answer"">","</h4>",1)
		' end if
		' body=replace(body,replaceStr,"")
   		' ' replaceStr=strCut(body,"<!--","-->",1)
		' ' body=replace(body,replaceStr,"")
		'body=replace(body,"<div class=""evaluation_wrap"">","")
		
	elseif instr(wstr,"<div class=""answer_con"">")>0 then
   
   		body=strCut(wstr,"<div class=""answer_con"">","</div>",2)
   		if body="" then 
   			body1=strCut(wstr,"<!--------- best answers  --------->","<div class=""operation"">",2)
   			body=strCut(body1,"<div class=""answer_con"">","<div class=""evaluation_wrap"">",1)
   		end if
  		 body=replace(body,"<div class=""evaluation_wrap"">","")
   		'body=replace(body,w8,"<span>"&w8&"</span>")
   		r = "(<img class=ed_capture src=""http://pic.wenwen.soso.com/p/.*?\..*?"">)"
   		body = py_z_replace(body,r,"")
	else
		body=""
   end if
	if instr(body,"<pre>")>0 and instr(body,"<pre>") < instr(body,"</pre>") then
		body=strCut(body,"<pre>","</pre>",2)
		body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'>"& body &"</pre>"
	else
		if body="" then 
			body="<div style='font-size:14;color:red;'>该问答暂无采纳答案！</div>"
		else
			'body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'>"& body &"</pre>"
		end if
	end if
   
elseif t="1" then
	

   url="http://wenda.so.com"&para

   wstr=getHTTPPage(url,"utf-8")
   body=strCut(wstr,"<div class=""resolved-cnt"">","</div>",2)
   tit=strCut(wstr,"<h2 class=""title js-ask-title"">","</h2>",2)
   '	body=strCut(body,"<div class=""mod-best-a","<div class=""mod-btns-answer-best",2)
	'body=strCut(body,"<div class=""qa-content"">","</div>",2)
	'r = "参考文献(.|[\r\n])*?<a.*?>.*?</a>"
	'body=py_z_replace(body,r,"")
	if body="" then body="<div style='font-size:14;color:red;'>该问答暂无采纳答案！</div>"
	body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'>"& body &"</pre>"
   
elseif t="2" then
	
	url="http://zhidao.baidu.com/"&para

	wstr=getHTTPPage(url,"gb2312")
	tit=strCut(wstr,"<span class=""ask-title","</span>",1)
	r = "<span.*?>"
	tit=py_z_replace(tit,r,"")
	r = "</span>"
	tit=py_z_replace(tit,r,"")
	body=strCut(wstr,"<pre id=""best-content","</pre>",1)
	r = "(<pre id=""best-content.*?>)*"
	body=py_z_replace(body,r,"")
	if body="" then
		body=strCut(wstr,"<pre id=""recommend-content","</pre>",1)
		r = "(<pre id=""recommend-content.*?>)*"
		body=py_z_replace(body,r,"")
	end if
	if body="" then body="<div style='font-size:14;color:red;'>该问答暂无采纳答案！</div>"
	body="<pre style='white-space: pre-wrap;word-break: break-all;height:auto;word-wrap:break-word;overflow: hidden;'>"& body 
	
end if
response.write "<div id=""main_s_content"" style='padding-top:50px;'><h3 style='font-size:15pt;text-align:center;'>"&tit&"</h3><div class=""cj_content"">"&body&"</div></div>"
act=trim(request("act"))
%>

</div>
<div id="popup_show"></div>
</body>
</html>
