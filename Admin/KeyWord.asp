<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：KeyWord.asp
'文件用途：关键词库拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	'Response.Write("无权限！")
	'Call FKDB.DB_Close()
	'Session.CodePage=936
	'Response.End()
End If

Dim KeyWord
dim listkeyword,TempItem,Thiswords,Host,ThiswordSHU1,ThiswordSHU2,nowpaiming,SVci,SVb1,neilian,iisid
Dim Newstr,oArray


'获取参数
Types=Clng(Request.QueryString("Type"))
iisid=Clng(Request.QueryString("iisid"))
'增加数据库表
		Call FKDB.DB_Open()
		on error resume next
		rs.open "select SVkeywords from keywordSV",conn,1,1
		if err.number<>0 then
		Sqlstr="create table keywordSV(id COUNTER CONSTRAINT PrimaryKey PRIMARY KEY,SVkeywords text(255),SVci int,SVpaiming text(255),SVb1 text(255),SVb2 text(255),SVb3 text(255))"
		Conn.Execute(Sqlstr)
		end if
		rs.close


Select Case Types
	Case 1
		Call KeyWordBox() '读取关键词库
	Case 2
		Call KeyWordDo()  '设置关键词库
	Case 3
		Call KeyWordDel() '删除关键词库
	Case 4
		Call KeyWordSet() '设置关键词类型
	Case 5
		Call KeyWordSetLink() '设置关键词内链
	Case 6
		Call KeyWordUpSVci() '更新关键词有效访问量
	case 7
		call KeyWordImport()	'导入关键词		
End Select
%>

<div id="Boxs" style="display:none">
  <div id="BoxsContent">
    <div id="BoxContent"> </div>
  </div>
  <div id="AlphaBox" onClick="$('select').show();$('#Boxs').hide()"></div>
</div>
<%
'==========================================
'函 数 名：KeyWordBox()
'作    用：读取关键词库
'参    数：
'==========================================
Sub KeyWordBox()
	
'------------------------------------关键词去重-----------------------------------------------------------
'listkeyword=UnEscape(keyword)
'	listkeyword=replace(listkeyword," ","")
'	listkeyword=replace(listkeyword,"　","")
'	listkeyword=replace(listkeyword,"｜","|")
'	listkeyword=replace(listkeyword,"|||","|")
'	listkeyword=replace(listkeyword,"||","|")
'	listkeyword=replace(listkeyword,"&nbsp;","")
'oArray = Split(listkeyword, "|")
'Newstr = " " 											'这里的值是一个空格
'For i=0 To UBound(oArray)
'    If Instr(Newstr, " " & oArray(i) & " ") = 0 Then 	'在oArray(i)的前后加一个空格
'        Newstr = Newstr & oArray(i) & " " 				'用空格分开
'    End If
'Next
'Newstr=trim(Newstr)										'去掉首尾空格
'Newstr=replace(Newstr," ","|")							'替换空格为|
'KeyWord=Newstr
'listkeyword=Newstr
'------------------------------------关键词去重-----------------------------------------------------------
	dim krs,i,id,Sqlstr,PageNow,j,keyw,w,typ
	PageNow=Trim(Request.QueryString("Page"))
	keyw=Trim(Request.QueryString("searchK"))	
	typ=Trim(Request.QueryString("typ"))	
%>
<script type="text/javascript">
<!--//
function show(id){
    var aiin  = document.getElementById(id);
    if(aiin.style.display != 'block'){
        aiin.style.display = 'block';
    }else{
        aiin.style.display = 'none'; 
    }
}


//获得coolie 的值

 

function cookie(name){    

   var cookieArray=document.cookie.split("; "); //得到分割的cookie名值对    

   var cookie=new Object();    

   for (var i=0;i<cookieArray.length;i++){    

      var arr=cookieArray[i].split("=");       //将名和值分开    

      if(arr[0]==name)return unescape(arr[1]); //如果是指定的cookie，则返回它的值    

   } 

   return ""; 

} 

 

function delCookie(name)//删除cookie

{

   document.cookie = name+"=;expires="+(new Date(0)).toGMTString();

}

 

function getCookie(objName){//获取指定名称的cookie的值

    var arrStr = document.cookie.split("; ");

    for(var i = 0;i < arrStr.length;i ++){

        var temp = arrStr[i].split("=");

        if(temp[0] == objName) return unescape(temp[1]);

   } 

}

 

function addCookie(objName,objValue,objHours){      //添加cookie

    var str = objName + "=" + escape(objValue);

    if(objHours > 0){                               //为时不设定过期时间，浏览器关闭时cookie自动消失

        var date = new Date();

        var ms = objHours*60*1000;

        date.setTime(date.getTime() + ms);

        str += "; expires=" + date.toGMTString();

   }

   document.cookie = str;

}

 

function SetCookie(name,value)//两个参数，一个是cookie的名子，一个是值

{

    var Days = 30; //此 cookie 将被保存 30 天

    var exp = new Date();    //new Date("December 31, 9998");

    exp.setTime(exp.getTime() + Days*24*60*60*1000);

    document.cookie = name + "="+ escape (value) + ";expires=" + exp.toGMTString();

}

function getCookie(name)//取cookies函数        

{

    var arr = document.cookie.match(new RegExp("(^| )"+name+"=([^;]*)(;|$)"));

     if(arr != null) return unescape(arr[2]); return null;

 

}

function delCookie(name)//删除cookie

{

    var exp = new Date();

    exp.setTime(exp.getTime() - 1);

    var cval=getCookie(name);

    if(cval!=null) document.cookie= name + "="+cval+";expires="+exp.toGMTString();

}

var baseUrl="http://win.qebang.net/web/json/";
var TjUrl="<%=TjUrl%>";
var TjID=<%=Tjid%>;
var e=encodeURIComponent;
var es=escape;
var hst=window.location.host;
var dh=10;
var sitess;

function getServers(){
	$.ajax({
		type: "get",
        async: false,
        url: baseUrl+"GetServerTime.asp?t=s3d3d5j2er4fj3ij2e32e87we&d="+window.location.host,
        dataType: "jsonp",
        jsonp: "jsoncallback",//传递给请求处理程序或页面的，用以获得jsonp回调函数名的参数名(一般默认为:callback)
        jsonpCallback:"ServerTime",//自定义的jsonp回调函数名称，默认为jQuery自动生成的随机函数名，也可以写"?"，jQuery会自动为你处理数据
        success: function(json){
			var obj=json.svrtime;
			$(".serverTime_S-E").html(obj[0].starttime+" 至 "+obj[0].endtime);
			var time_v=obj[0].endtime.split("-")[0]-obj[0].starttime.split("-")[0];
			$(".visits_target").html(time_v*10000);
			delCookie(hst+".starttime");
			delCookie(hst+".endtime");
			delCookie(hst+".time_vs");
			addCookie(hst+".starttime",obj[0].starttime,dh);
			addCookie(hst+".endtime",obj[0].endtime,dh);
			addCookie(hst+".time_vs",time_v*10000,dh);
			getVisits();
        },
        error: function(){
			$(".serverTime_S-E").html("获取数据异常，请重试！");
			delCookie(hst+".doSearchInfos");
        }
	});
}

function getVisits(){
	$.ajax({
		type: "get",
        async: false,
        url: baseUrl+"GetVisits.asp?t=gkwSVdf2wer2c&d="+window.location.host,
        dataType: "jsonp",
        jsonp: "jsoncallback",//传递给请求处理程序或页面的，用以获得jsonp回调函数名的参数名(一般默认为:callback)
        jsonpCallback:"Visits",//自定义的jsonp回调函数名称，默认为jQuery自动生成的随机函数名，也可以写"?"，jQuery会自动为你处理数据
        success: function(json){
			var obj=json.vsts;
			var v=0;
			if(obj[0].v!=""){v=obj[0].v};
			$(".visits_done").html(v);
			delCookie(hst+".visits_done");
			addCookie(hst+".visits_done",v,dh);
			addCookie(hst+".doSearchInfos",1,dh);
        },
        error: function(){
			$(".visits_done").html("获取数据异常，请重试！");
			delCookie(hst+".doSearchInfos");
        }
	});
}

var sites;
function GetRmtKlist(typ){
	$(".importChk").attr("disabled",true);
	for(var i=0;i<4;i++){
		if(i==typ){
			$(".date"+typ).addClass("colorred");
		}
		else{
			$(".date"+i).removeClass("colorred");
		}
	};
	$("#content").html("<tr><td><img src=\"images/loading3.gif\"></td></tr>");
	try
	{
		$.ajax({
			type: "get",
        	async: false,
       	 	url: TjUrl+"/json/k.json.asp?t=6d2fvvds2sxcs3&tjid=<%=Tjid%>&typ="+typ,
        	dataType: "jsonp",
        	jsonp: "jsoncallback",//传递给请求处理程序或页面的，用以获得jsonp回调函数名的参数名(一般默认为:callback)
        	jsonpCallback:"klist",//自定义的jsonp回调函数名称，默认为jQuery自动生成的随机函数名，也可以写"?"，jQuery会自动为你处理数据
        	success: function(json){
				sites=json.jsonklists;
				//alert(jsonlists.length);
				OutputHtml();
				var str;
				switch(typ){
					case 0:
						str="本月";
						$(".exportXls").attr("href","<%=TjUrl%>/user/exportxls-shangwin.asp?id=<%=tjid%>&act=0");
						break;
					case 1:
						str="今日"
						$(".exportXls").attr("href","<%=TjUrl%>/user/exportxls-shangwin.asp?id=<%=tjid%>&act=1");
						break;
					case 2:
						str="本年"
						$(".exportXls").attr("href","<%=TjUrl%>/user/exportxls-shangwin.asp?id=<%=tjid%>&act=2");
						break;
					case 3:
						str="本周"
						$(".exportXls").attr("href","<%=TjUrl%>/user/exportxls-shangwin.asp?id=<%=tjid%>&act=3");
						break;
					default:
						str="本月"
						$(".exportXls").attr("href","<%=TjUrl%>/user/exportxls-shangwin.asp?id=<%=tjid%>&act=0");
				}
				$(".importKwd").val("导入"+str+"关键词到库");
				$(".exportXls").text("导出"+str+"关键词到Excel");
       	 	},
        	error: function(){
				$("#content").html("<tr><td>获取数据异常，请重试！</td></tr>");
        	}
		});
	}
	catch(e){
		$("#content").html("<tr><td>获取数据异常，请重试！</td></tr>");
	}
	//ymPrompt.win('<%=TjUrl%>/user/k.asp?type=6&id=<%=Tjid%>&time=2011-1-1&time2=<%=date()%>',630,475,'搜索引擎关键词来源',null,null,null,true);
}

function GotoPage(num){ //跳转页
	Page = num;
	$(".importChk").attr("disabled",true);
	OutputHtml();
} 

function zs(id){return document.getElementById(id);} //定义获取ID的方法
var PageSize = 10; //每页个数
var Page = 1; //当前页码

function OutputHtml(){
	
	
	//var obj = eval(jsonlists);  //获取JSON
	//var sites = obj.jsonklists;
	
	//获取分页总数
	var Pages = Math.floor((sites.length - 1) / PageSize) + 1; 
	if(Page < 1)Page = 1;  //如果当前页码小于1
	if(Page > Pages)Page = Pages; //如果当前页码大于总数
	var Temp = "";
	
	var BeginNO = (Page - 1) * PageSize + 1; //开始编号
	var EndNO = Page * PageSize; //结束编号
	if(EndNO > sites.length) EndNO = sites.length;
	if(EndNO == 0) BeginNO = 0;
	
	if(!(Page <= Pages)) Page = Pages;
	zs("total").innerHTML = "总共:<strong class='cff000'>" + sites.length + "</strong>&nbsp;&nbsp;当前:<strong class='cff000'>" + BeginNO + "-" + EndNO + "</strong>"; 
	
	//分页
	if(Page > 1 && Page !== 1){Temp ="<a href='javascript:void(0)' onclick='GotoPage(1);return false;'>第一页</a> <a href='javascript:void(0)' onclick='GotoPage(" + (Page - 1) + ");return false;'>上一页</a>&nbsp;"}else{Temp = "第一页 上一页&nbsp;"};
	
	//完美的翻页列表
	var PageFrontSum = 3; //当页前显示个数
	var PageBackSum = 3; //当页后显示个数
	
	var PageFront = PageFrontSum - (Page - 1);
	var PageBack = PageBackSum - (Pages - Page);
	if(PageFront > 0 && PageBack < 0)PageBackSum += PageFront; //前少后多，前剩余空位给后
	if(PageBack > 0 && PageFront < 0)PageFrontSum += PageBack; //后少前多，后剩余空位给前
	var PageFrontBegin = Page - PageFrontSum;
	if(PageFrontBegin < 1)PageFrontBegin = 1;
	var PageFrontEnd = Page + PageBackSum;
	if(PageFrontEnd > Pages)PageFrontEnd = Pages;
	
	if(PageFrontBegin != 1) Temp += '<a href="javascript:void(0)" onclick="GotoPage(' + (Page - 10) + ');return false;" title="前10页">..</a>';
	for(var i = PageFrontBegin;i < Page;i ++){
		Temp += " <a href='javascript:void(0)' onclick='GotoPage(" + i + ");return false;'>" + i + "</a>";
	}
	Temp += " <strong class='c006090'>" + Page + "</strong>";
	for(var i = Page + 1;i <= PageFrontEnd;i ++){
		Temp += " <a href='javascript:void(0)' onclick='GotoPage(" + i + ");return false;'>" + i + "</a>";
	}
	if(PageFrontEnd != Pages) Temp += " <a href='javascript:void(0)' onclick='GotoPage(" + (Page + 10) + ");return false;' title='后10页'>..</a>";
	
	if(Page != Pages){Temp += "&nbsp;&nbsp;<a href='javascript:void(0)' onclick='GotoPage(" + (Page + 1) + ");return false;'>下一页</a> <a href='javascript:void(0)' onclick='GotoPage(" + Pages + ");return false;'>尾页</a>"}else{Temp += "&nbsp;&nbsp;下一页 尾页"}
	
	zs("pagelist").innerHTML = Temp;
	
	//输出数据
	
	if(EndNO == 0){ //如果为空
		zs("content").innerHTML += "<tr><td>无数据</td></tr>";
		return;
	}
	var html = "<tr><td width='44' align='center' class='kwlt1'  style='background:#ECF5FF;border:2px #DBECF7 solid;'><input type='checkbox' name='checkbox1s' value='Check All' class='chkall' title='点击全/反选'></td><td width='322' align='center' class='kwlt1' style='background:#ECF5FF;border:2px #DBECF7 solid;'>来访关键词</td><td width='56'  class='kwlt1' style='width:56px;background:#ECF5FF;border:2px #DBECF7 solid;text-align:center;'>访问量</td></tr>";
		
	for(var i = BeginNO - 1;i < EndNO;i ++){
		html += "<tr class='mm' style='border:2px #DBECF7 solid;'>";
		html += "<td class='kwlt1' align='center' style='background:#ECF5FF;border:2px #DBECF7 solid;'><input type=\"checkbox\" name=\"chkkwd\" class=\"chkkwd\" value=\""+sites[i].sKeyword+"\"/></td>";
		html += "<td class='kwlt3' style='border:1px solid #ECF5FF'>"+sites[i].sKeyword+"</td>";
		html += "<td class='kwlt3' align='center' style='border:1px solid #ECF5FF;color:#006090'>"+sites[i].cnt+"</td>";
		//html += "<p class='url'><span>" +sites[i].Name+ "</span></p></a>";
		html += "</tr>";
		
	}
	$("#content").html(html);
	clickShow(); //调用鼠标点击事件
	
	//键盘左右键翻页
	document.onkeydown=function(e){
		var theEvent = window.event || e;
		var code = theEvent.keyCode || theEvent.which;
		if(code==37){
			if(Page > 1 && Page !== 1){
				GotoPage(Page - 1);
				
			}
		}
		if(code==39){
			if(Page != Pages){
				GotoPage(Page + 1);
			}
		}
	}
	
	
	//鼠标滚轮翻页
function handle(delta){
	if (delta > 0){
		if(Page > 1 && Page !== 1){
			GotoPage(Page - 1);
		}		
	}
	else{
		if(Page != Pages){
			GotoPage(Page + 1);
		} 
	}
}

function wheel(event){
	var delta = 0;
	if (!event) /* For IE. */
		event = window.event;
	if (event.wheelDelta) { /* IE或者Opera. */
		delta = event.wheelDelta / 120;
		/** 在Opera9中，事件处理不同于IE
		*/
	if (window.opera)
		delta = -delta;
	} else if (event.detail) { /** 兼容Mozilla. */
	/** In Mozilla, sign of delta is different than in IE.
	* Also, delta is multiple of 3.
	*/
	delta = -event.detail / 3;
	}
	/** 如果 增量不等于0则触发
	* 主要功能为测试滚轮向上滚或者是向下
	*/
	if (delta)
		handle(delta);
}

/** 初始化 */
if (window.addEventListener)
	/** Mozilla的基于DOM的滚轮事件 **/
	window.addEventListener("DOMMouseScroll", wheel, false);
	/** IE/Opera. */
	window.onmousewheel = document.onmousewheel = wheel;
	
	
}

//获取链接地址和网站名称
function showLink(source){
	var siteUrl = zs("siteurl");
	var siteName = zs("sitename");
	var description = zs("description");
	
	if(source.getAttribute("rel") == "bookmark"){
		var url = source.getAttribute("href");
		var title = source.getAttribute("title");
		siteUrl.innerHTML = "<a href='" + url + "' target='_blank'>"+ url +"</a>";
		siteName.innerHTML = title;
	}
	
}

//鼠标点击事件
function clickShow(){
	var links = zs("content").getElementsByTagName("a");
	for(var i=0; i<links.length; i++){
		var url = links[i].getAttribute("href");	
		var title = links[i].getAttribute("title");
		links[i].onclick = function(){
			showLink(this);
			return false;
		}
	}
}

function GetYXFW(){
	$(".kwd_vst a").html("<img src=\"images/loading3.gif\">");
	$(".kwd_vst").each(function(index,domEle){
		var k=$(domEle).prev(".strkwd").text();
		url=TjUrl+"/json/GetVisits.asp?t=gkwSVdf2wer2c&kwd="+es(k)+"&tjid=<%=Tjid%>&callback=?";
		$.getJSON(url,function(result){
			var totals;
			totals=result.total;
			if(totals.length>10){totals=0};
			$(domEle).children("a").html(totals);
		})
	})
}

function GetPaim(){
		$(".bd,.q360,.ss,.sg").html("<img src=\"images/loading3.gif\">");
		$(".mm").each(function(index,domEle){
			var k=$(domEle).children(".strkwd").text();
			url=baseUrl+"GetPm.asp?t=gkwSVdf2wer2c&kwd="+es(k)+"&iisid=<%=iisid%>&callback=?";
			$.getJSON(url,function(result){
				var bd,q360,ss,sg;
				bd=result.SVbd;
				q360=result.SV360;
				ss=result.SVss;
				sg=result.SVsg
				$(domEle).children(".bd").html(bd);
				$(domEle).children(".q360").html(q360);
				$(domEle).children(".ss").html(ss);
				$(domEle).children(".sg").html(sg);
			})
		})	
}

$(document).ready(function(){
//	if($(".mm").length>0){
//	$("#yxfw").click();
//	$("#sxpm").click();
//	}
	$('li.mainlevel').mousemove(function(){
 	 	$(this).find('ul').slideDown();//you can give it a speed
  	});
  	$('li.mainlevel').mouseleave(function(){
  		$(this).find('ul').slideUp("fast");
  	});
  	$(".mm").mouseover(function(){
  		$(this).css("background","#e4FBFF");
  	});
  	$(".mm").mouseout(function(){
  		$(this).css("background","#fff");
  	});
	$("input.chkkwd").die().live("click",function(){
		if($("input.chkkwd:checked").length==0){
			$(".importChk").attr("disabled",true);
		}
		else{
			$(".importChk").attr("disabled",false);
		}
	});
	$("input.chkall").die().live("click",function(){
		$("input.chkkwd").attr("checked",$("input.chkall").attr("checked"));
		if($("input.chkall").attr("checked")) {
			$(".importChk").attr("disabled",false);
		}
		else{
			$(".importChk").attr("disabled",true);
		}
	})
	$(".importKwd").click(function(){
		if(window.confirm("确定导入到关键词库吗？")){
		
		var oldmsg=$(".importKwd").val();
		$(".importKwd").val("导出中...");
		$(".importKwd").attr("disabled",true);
		var siteslst;
		for(var i=0;i<sites.length;i++){
			if(i==0){
				siteslst=sites[i].sKeyword;
			}
			else{
				siteslst+="{0}"+sites[i].sKeyword;
			}		
		}
		$.ajax({
			type: "POST",
       		async: false,
			data:"t=3&lst="+siteslst,
     		url: "KeyWord.asp?Type=7",
			//处理数据
       		success: function(msg){
				tan("操作成功！");
				$(".importKwd").val(oldmsg);
				$(".importKwd").attr("disabled",false);
				//updateSVci(lst);
     		},
      		error: function(){
				tan("获取数据异常，请重试！");
				$(".importKwd").val(oldmsg);
				$(".importKwd").attr("disabled",false);
       		}
		});
			
		}
	})
	$(".importChk").click(function(){
		var oldmsg=$(".importChk").val();
		$(".importChk").attr("disabled",true);
		$(".importChk").val("导出中...");
		var imlist="";
		$("input.chkkwd:checked").each(function(i){
			if(i==0){
				imlist=$(this).val();
			}
			else{
				imlist+="{0}"+$(this).val();
			}
		});
		$.ajax({
			type: "POST",
       		async: false,
			data:"t=3&lst="+imlist,
     		url: "KeyWord.asp?Type=7",
			//处理数据
       		success: function(msg){
				tan("操作成功！");
				$(".importChk").val(oldmsg);
				$(".importChk").attr("disabled",false);
				//updateSVci(lst);
     		},
      		error: function(){
				tan("获取数据异常，请重试！");
				$(".importChk").val(oldmsg);
				$(".importChk").attr("disabled",false);
       		}
		});
	})
	//$(".kwd_vst a").click(function(){
	//try
	//{
			//ymPrompt.win(TjUrl+"/json/GetReferrers.asp?t=gkwSVdf2wer2c&tjid="+TjID+"&kwd="+es($(this).parent().prev().text())+"&jsoncallback=?");
		//$.getJSON(TjUrl+"/json/GetReferrers.asp?t=gkwSVdf2wer2c&tjid="+TjID+"&kwd="+es($(this).parent().prev().text())+"&jsoncallback=?",function(data){
		//})
	//}
	//catch(e){
		//ymPrompt.alert("获取数据异常，请重试！"+e.message+"\n"+e.description+"\n"+e.number+"\n"+e.name);
	//}
		
	//});
	$(".pagelist-S a").live("click",function(){
		var base =  $(this).attr('num');
  		GotoPageS(base);
	})
})
function getSourceF(strK){
	Page=1;
		$.ajax({
			type: "get",
        	async: false,
       	 	url: TjUrl+"/json/GetReferrers.asp?t=gkwSVdf2wer2c&tjid="+TjID+"&kwd="+es(strK),
        	dataType: "jsonp",
        	jsonp: "jsoncallback",//传递给请求处理程序或页面的，用以获得jsonp回调函数名的参数名(一般默认为:callback)
        	jsonpCallback:"kenginelist",//自定义的jsonp回调函数名称，默认为jQuery自动生成的随机函数名，也可以写"?"，jQuery会自动为你处理数据
        	success: function(json){
				sitess=json.Referrerslist;
				OutputHtmlS(0,strK);
		//$(this).parent().parent().addClass("cur");
				//tipsWindown("提示","text:"+strh,"250","150","true","","true","msg")
				//ymPrompt.win({message:strh,width:360,height:300,title:'关键词：【'+strk+'】的搜索引擎来源及来访IP',showMask:false})
       	 	},
        	error: function(){
				ymPrompt.alert("获取数据异常，请重试！");
        	}
		});
}
function GotoPageS(num){ //跳转页
	Page = parseInt(num);
	var str=OutputHtmlS(1,"");
	$("#biboxs").html(str);
} 
function OutputHtmlS(t,strk){
	
	
				//alert(sitess.length);return;
	//var obj = eval(jsonlists);  //获取JSON
	//var sites = obj.jsonklists;
	//获取分页总数
	var Pages = Math.floor((sitess.length - 1) / PageSize) + 1; 
	if(Page < 1)Page = 1;  //如果当前页码小于1
	if(Page > Pages)Page = Pages; //如果当前页码大于总数
	var Temp = "";
	
	var BeginNO = (Page - 1) * PageSize + 1; //开始编号
	var EndNO = Page * PageSize; //结束编号
	if(EndNO > sitess.length) EndNO = sitess.length;
	if(EndNO == 0) BeginNO = 0;
	
	if(!(Page <= Pages)) Page = Pages;
	//$("#total").html("总共:<strong class='cff000'>" + sitess.length + "</strong>&nbsp;&nbsp;当前:<strong class='cff000'>" + BeginNO + "-" + EndNO + "</strong>"); 
	
	//分页
	if(Page > 1 && Page !== 1){Temp ="<a num=1>第一页</a> <a num="+(Page-1)+">上一页</a>&nbsp;"}else{Temp = "第一页 上一页&nbsp;"};
	
	//完美的翻页列表
	var PageFrontSum = 3; //当页前显示个数
	var PageBackSum = 3; //当页后显示个数
	
	var PageFront = PageFrontSum - (Page - 1);
	var PageBack = PageBackSum - (Pages - Page);
	if(PageFront > 0 && PageBack < 0)PageBackSum += PageFront; //前少后多，前剩余空位给后
	if(PageBack > 0 && PageFront < 0)PageFrontSum += PageBack; //后少前多，后剩余空位给前
	var PageFrontBegin = Page - PageFrontSum;
	if(PageFrontBegin < 1)PageFrontBegin = 1;
	var PageFrontEnd = Page + PageBackSum;
	if(PageFrontEnd > Pages)PageFrontEnd = Pages;
	
	if(PageFrontBegin != 1) Temp += '<a num='+(Page-10)+' title="前10页">..</a>';
	for(var i = PageFrontBegin;i < Page;i ++){
		Temp += " <a num="+i+">" + i + "</a>";
	}
	Temp += " <strong class='c006090'>" + Page + "</strong>";
	for(var i = Page + 1;i <= PageFrontEnd;i ++){
		Temp += " <a num="+i+">" + i + "</a>";
	}
	if(PageFrontEnd != Pages) Temp += " <a num="+(Page+10)+" title='后10页'>..</a>";
	if(Page != Pages){Temp += "&nbsp;&nbsp;<a num="+(Page+1)+">下一页</a> <a num="+(Pages)+">尾页</a>"}else{Temp += "&nbsp;&nbsp;下一页 尾页"}
	
	//$("#pagelist").html(Temp);
	
	//输出数据
	
	if(EndNO == 0){ //如果为空
		$("#jsonbox").html($("#jsonbox").html() + "<tr><td>无数据</td></tr>");
		return;
	}
	var html1 = "<tr><td width='86'  class='kwlt1' style='background:#ECF5FF;border:2px #DBECF7 solid;text-align:center;padding-left:0px'>搜索引擎</td><td width='122' align='center' class='kwlt1' style='background:#ECF5FF;border:2px #DBECF7 solid;padding-left:0px'>IP</td><td width='136'  class='kwlt1' style='background:#ECF5FF;border:2px #DBECF7 solid;text-align:center;padding-left:0px'>来访时间</td></tr>";
	var sSearchEngine;
	for(var i = BeginNO - 1;i < EndNO;i ++){
		sSearchEngine=sitess[i].sSearchEngine;
		if(sSearchEngine.indexOf("www.baidu.com")>-1){
			sSearchEngine="<font style='color:#dc0900'>百 度</font>";
		}
		else if (sSearchEngine.indexOf("www.soso.com")>-1){
			sSearchEngine="<font style='color:#0084bf'>搜 搜</font>";
		}
		else if (sSearchEngine.indexOf("www.sogou.com")>-1){
			sSearchEngine="<font style='color:#56017d'>搜 狗</font>";
		}
		else if (sSearchEngine.indexOf("www.google.com")>-1){
			sSearchEngine="<font style='color:#1548ed'>谷 歌</font>";
		}
		else if (sSearchEngine.indexOf("www.so.com")>-1){
			sSearchEngine="<font style='color:#3ba80d'>3 6 0</font>";
		}
			
		//sSearchEngine=removeHTMLTag(removeHTMLTag);
		html1 += "<tr class='mm' style='border:2px #DBECF7 solid;'>";
		html1 += "<td class='kwlt3' style='text-align:center;padding-left:0px;border:1px solid #ECF5FF;'>"+sSearchEngine+"</td>";
		html1 += "<td class='kwlt3' style='border:1px solid #ECF5FF'>"+sitess[i].sip+"</td>";		
		html1 += "<td class='kwlt3' style='border:1px solid #ECF5FF;color:#006090'>"+sitess[i].stime+"</td>";
		//html += "<p class='url'><span>" +sites[i].Name+ "</span></p></a>";
		html1 += "</tr>";
		
	}
		//ymPrompt.win(message:''+html1+'',width:360,height:300,title:'关键词：的搜索引擎来源及来访IP');
	if (t==0){
		Page=1;
		ymPrompt.win('<div id="biboxs"><table id="Sbox-D">'+html1+'</table><table><tr><td class="pagelist-S">'+Temp+'</td></tr></table></div>',360,300,'本月与上月搜索引擎来源及来访IP')
	}
	else{
	
		$("#biboxs").html('<table id="Sbox-D">'+html1+'</table><table><tr><td class="pagelist-S">'+Temp+'</td></tr></table>');
	}
	//$("#jsonbox").html("<table>"+html1+"</table>");
}	

function getHistory(){
	//ymPrompt.win('<div class=\'myContent\'>普通弹出窗口</div>',300,200,'普通弹窗测试')
	//ymPrompt.win({message:'http://localhost/?siteid=<%=iisid%>',width:500,height:300,title:'腾讯QQ官方网站',iframe:true})
	//ymPrompt.win('http://localhost/?siteid=<%=iisid%>',600,400,'历史排名',null,null,null,true)
		$.webox({
			height:280,
			width:600,
			bgvisibel:true,
			title:'iframe弹出层调用',
			iframe:'http://localhost/?siteid=<%=iisid%>&r='+Math.random()
		});
}

function importKwd(strList){
	$.ajax({
		type: "POST",
       	async: false,
		data:"t=3&lst="+strList,
     	url: "KeyWord.asp?Type=7",
		//处理数据
       	success: function(msg){
			tan("操作成功！");
			//updateSVci(lst);
     	},
      	error: function(){
			tan("获取数据异常，请重试！");
       	}
	});
}
//-->
</script>
<style type="text/css">
	.SeoAnalyseCleft{color:#333;}
	.SeoAnalyseCright{color:#D5006A;font-weight:bolder;}
	.SeoAnalyseT {border-collapse:collapse;}
	.kwlt6{width:260px;}
	.kwlt66{display:inline-block;}
	.saveResult{width:120px;display:inline-block;float:left;}
	.ainline a{display:inline;}
	.tixinginfo{color:red;}
	li{list-style-type:none;}
	.nav{width:84px;}
	.tixinginfo {font-size:14px;}
	.tixinginfo b{color:#000;font-weight:normal;width:116px;}
.tixinginfo li{margin:0px;padding:0px;text-indent:0px;line-height:26px;height:26px;}
.mainlevel { float:left;/*IE6 only*/}
	.mainlevel a {color:#000; text-decoration:none;display:block;}
.mainlevel a:hover {color:#000; text-decoration:none;}
.mainlevel ul {display:none; position:absolute;background:#ffe60c;margin:0px;padding:0px;}
.mainlevel ul li {border-top:0px solid #fff; /*IE6 only*/}

.headjson{clear:both;display:block;height:30px;line-height:30px;width:750px;}
.headjson a{font-size:14px;}
#BoxContents .colorred{color:red;}
.cff000{color:#ff0000;}
#pagelist,#pagelist a{font-size:14px;}
.c006090{color:#006090;}

.importKwd{margin-left:10px;text-align:center;}
.importKwd,.importChk{padding:0px;cursor:pointer;height:24px;line-height:24px;width:140px;}
#keywordlisttable .kwd_vst a,#biboxs td a{cursor:pointer;}
#biboxs td {padding-left:10px;}
#keywordlisttable .mm a{}
</style>

<form id="KeyWordSet" name="KeyWordSet" method="post" action="KeyWord.asp?Type=2" onSubmit="return false;">
  <div id="BoxTop" style="width:98%;"><span> <%if FkFun.CheckLimit("System22") then response.write "SEO效果分析中心" else response.Write "关键词库"%></span><a style="display:none;" onClick="$('#Boxs').hide();$('select').show();return false;"></a> </div>
  <div id="BoxContents" style="width:98%;">
  	
    <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="10" align="center"></td>
      </tr>
      <tr id="kw1" <%If not FkFun.CheckLimit("System22") Then response.write "style=""display:none;""" else response.Write "style=""display:block;"""%>>
        <td align="center"><a href="javascript:void(0);" onClick="viewcilie();return false;" class="keywdleft-b">关键词库列表</a> <%If not FkFun.CheckLimit("System22") Then%><a href="javascript:void(0);" onClick="viewciku();return false;" class="keywdleft-a">关键词库设置</a> <%else%> <a href="javascript:void(0);" onClick="viewciku1();return false;" class="keywdleft-a">关键词库(客服)</a> <a href="javascript:void(0);" onClick="viewciku();return false;" class="keywdleft-a">关键词库(客户)</a><%end if%><!--select id="typ" name="typ" <%'if not FkFun.CheckLimit("System22") then response.write "style=""display:none;"""%>><option value="0">所有</option><!--option value="1">主营词</option><option value="2">商业词</option><option value="3">重点词</option></select> <input type="text" value="" name="searchK" id="searchK" style="padding:0px;margin:0px;line-height:18px;height:18px;"/> <input type="button" value="关键词搜索"  style="padding:0px;margin:0px;border:solid #CCCCCC 1px;line-height:20px;height:23px;background-image:url(Images/BgLine.png);background-position:0 -160px;cursor:pointer;" onclick='ShowBox("KeyWord.asp?Type=1&typ="+document.getElementById("typ").value+"&searchK="+encodeURIComponent(document.getElementById("searchK").value)+"&Page=<%=PageNow%>")'/> <input type="button" value="返回关键词首页"  style="padding:0px;margin:0px;border:solid #CCCCCC 1px;line-height:20px;height:23px;background-image:url(Images/BgLine.png);background-position:0 -160px;cursor:pointer;" onclick='ShowBox("KeyWord.asp?Type=1")'/--> 
          <div class="tixinginfo" style="float:left;">
		  <ul>
		  <li><b style="color:red;">此处列表显示的是客服添加的重点优化关键词，与客户设置的关键词无关，请设置在10-20个左右,最多不能超过30个。</b></li>
		  </ul>
		  </div>
          <p>
          <table width="100%" id="keywordlisttable">
            <tr>
              <td width="44" rowspan="2" align="center" class="kwlt1"  title="点击全/反选" ><input type="checkbox" name="checkbox" value="Check All" onClick="SelectAll('chkID')" title="点击全/反选"></td>
              <!--td width="56" rowspan="2"  class="kwlt1" style="width:58px;">序号</td-->
              <td width="373" rowspan="2" align="center" class="kwlt1">关键词</td>
              <td width="94" rowspan="2" align="center" class="kwlt1"><a href="javascript:void(0);" onclick="GetYXFW();return false;" id="yxfw" title="刷新有效访问量">有效访问</a></td>
              <td height="21" colspan="4"  align="center" class="kwlt1"><a href="javascript:void(0);" onclick="GetPaim();return false;" id="sxpm" title="刷新排名">排名</a></td>
              <td width="236" rowspan="2"  align="center" class="kwlt1">操作</td>
            </tr>
            <tr>
              <td width="176" height="21"  align="center" class="kwlt1" >百度</td>
              <td width="102"  align="center" class="kwlt1">360</td>
              <td width="101"  align="center" class="kwlt1">搜狗</td>
              <td width="74"  align="center" class="kwlt1">搜搜</td>
            </tr>
            
            <%
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	if keyw<>"" then
		w=" and SVkeywords like '%%"&keyw&"%%'"
	else
	 	w=""
	end if
	if typ="" or typ="0" then
		w=w&""
	else
		w=w&" and SVb1='"&typ&"'"
	end if
	dim bd_pm,pm_360,ss_pm,sg_pm
	set krs=server.CreateObject("adodb.recordset")
	Sqlstr="select id,SVkeywords,SVpaiming,SVci,SVb1,SVb2,SVb3 from [keywordSV] where 1=1 "&w&" order by SVci desc, SVb1 desc, id"
	krs.Open Sqlstr,Conn,1,1
	If Not krs.Eof Then
		PageAll=krs.RecordCount
		i=0
		While (Not krs.Eof)
			i=i+1
			j=(PageNow-1)*PageSizes+i
			id=krs("id")
			Thiswords=krs("SVkeywords")			
			Dim ubArr
			ubArr=PageSizes%>
					<tr class="mm"><td class='kwlt1' width="44" align="center" ><input type="checkbox" name="chkID" id="chkID" value="<%=id%>"/></td>
					<td class='kwlt3 strkwd' style="text-align:left"><b><%=Thiswords%></b></td>
					<td class='kwlt3 kwd_vst' id="kwdcount<%=i%>"><a href="javascript:void(0);" onClick="getSourceF('<%=Thiswords%>');return false;" title="点击查看关键词来源">0</a></td>	
					<td class='kwlt3 bd'>0</td>
					<td class='kwlt3 q360'>0</td>
					<td class='kwlt3 ss'>0</td>
					<td class='kwlt3 sg'>0</td>
					<td class='kwlt3 ainline' valign="middle"><a href="javascript:void(0);" onClick="DelIt('确定删除吗？','KeyWord.asp?Type=3&id=<%=id%>&iisid=<%=request("iisid")%>','BoxContent','KeyWord.asp?Type=1&iisid=<%=request("iisid")%>');return false;">删除该词</a></td>
					</tr>
					<%
					
					
			
			krs.MoveNext
		Wend
	End If
	krs.Close
	
	dim KeyWorddat
	Sqlstr="select SVkeywords from [keywordSV]"
	set krs=conn.execute(Sqlstr)
	If Not krs.Eof Then
		dim m
		m=0
		listkeyword=""
		KeyWorddat=""
		do while Not krs.Eof
			if m=0 then
				listkeyword=FilterText(krs("SVkeywords"))
				KeyWorddat=FilterText(krs("SVkeywords"))
			else
				KeyWorddat=KeyWorddat&"|"&FilterText(krs("SVkeywords"))
				listkeyword=listkeyword&","&FilterText(krs("SVkeywords"))
			end if
			m=m+1
			krs.movenext
			if krs.eof then exit do
		loop
	else
		listkeyword=""
		KeyWorddat=""
	end if	
	krs.close
	listkeyword=cxarraynull(listkeyword,",")
	
	if FKFso.IsFile("KeyWordC.dat") then
		KeyWorddat=FKFso.FsoFileRead("KeyWordC.dat")
	else
		call FKFso.CreateFile("KeyWordC.dat",KeyWorddat)
	end if
	
	dim lenkwddat,arrkwddat
	if KeyWorddat="" then 
		lenkwddat=0
	else
		if instr(KeyWorddat,"|")=0 then
			lenkwddat=1
		else
			arrkwddat=split(KeyWorddat,"|")
			lenkwddat=ubound(arrkwddat)+1
		end if
	end if
            %>
          </table><b style="color:red;float:right;display:inline-block;line-height:26px;padding-right:15px;font-size:14px;">共：<%=PageAll%>个关键词</b>
          <p align="left">
          
		  <input onClick="chkdel();" type="button" class="Button" name="Submit" value="删除所选" id="id-del"> <input onClick="window.open('http://history.qebang.cn/?siteid=<%=iisid%>','历史排名','width='+ screen.width-10 +',height='+ screen.height +',top=0,left=0,toolbar=no,menubar=no,scrollbars=yes, resizable=yes,location=no, status=no,titlebar=no');" type="button" class="Button" name="view-history" value="历史排名"> </p>
				    </td>
      </tr>
      <tr id="kw2" <%if not FkFun.CheckLimit("System22") then response.write "style=""display:block;""" else response.Write "style=""display:none;"""%>>
        <td align="center"><a href="javascript:void(0);" onClick="viewcilie();return false;" class="keywdleft-a" <%if FkFun.CheckLimit("System22") then response.write "style=""display:block;""" else response.Write "style=""display:none;"""%>>关键词库列表</a><a href="javascript:void(0);" onClick="viewciku();return false;" class="keywdleft-b">关键词库设置</a><%if FkFun.CheckLimit("System22") then %><a href="syExportXLS.asp?act=1" target="_blank" class="keywdleft-a">关键词导出</a><%end if%>
          <div class="tixinginfo"><b>间隔</b>：关键词与关键词之间用半角状态下的“|”符号隔开，最后一个关键词不需要“|”符号。<br>
            <b>个数</b>：关键词在20～100个为宜。当前关键词数：<font color="red"><%=lenkwddat%></font>个</div>
          <textarea name="KeyWord" cols="99%" style="width:99%;" rows="10" class="TextArea" id="KeyWord"><%=KeyWorddat%></textarea>
          <br /></td>
      </tr>
	  <tr id="kw3" style="display:none;">
	  	<td><a href="javascript:void(0);" onClick="viewcilie();return false;" class="keywdleft-a" <%if FkFun.CheckLimit("System21") then response.write "style=""display:block;""" else response.Write "style=""display:none;"""%>>关键词库列表</a><a href="javascript:void(0);" onClick="viewciku1();return false;" class="keywdleft-a">关键词库(客服设置)</a><a class="exportXls keywdleft-a" href="" target="_blank">导出今日关键词到Excel</a>
          <br />
		  <span class="headjson">
		  	<a href="javascript:void(0);" onClick="GetRmtKlist(1);return false;" class="date1">今天</a> | <a href="javascript:void(0);" onClick="GetRmtKlist(0);return false;" class="date0">本月</a> | <a href="javascript:void(0);" onClick="GetRmtKlist(2);return false;" class="date2">本年</a> | <a href="javascript:void(0);" onClick="GetRmtKlist(3);return false;" class="date3">本周</a> <input class="importKwd Button" type="button" value="导入今日关键词到词库"/> &nbsp; <input type="button" class="importChk Button" value="导入选中关键词"/> 
		  </span>
		  <table id="content" style="clear:both">
		  </table>
			<div id="pager">
            	<div id="total"></div>
            	<div id="pagelist"></div>
        	</div>		</td>
	  </tr>
      <tr id="kw4" style="display:none;">
        <td align="center"><a href="javascript:void(0);" onClick="viewcilie();return false;" class="keywdleft-a">关键词库列表</a><a href="javascript:void(0);" onClick="viewciku1();return false;" class="keywdleft-b">关键词库(客服设置)</a>
          <%
		  If FkFun.CheckLimit("System22") Then%>
          <a onClick="viewYX();GetRmtKlist(1);return false;" class="keywdleft-a" href="javascript:void(0);">搜索来源关键词</a><a href="syExportXLS.asp?act=0" target="_blank" class="keywdleft-a">关键词导出</a>
          <%end if%>
          <div class="tixinginfo"><b>间隔</b>：关键词与关键词之间用半角状态下的“,”符号隔开，最后一个关键词不需要“,”符号。<br>
          <b>个数</b>：关键词在10～20个为宜,最多不超过30个。</div>
          <textarea name="KeyWord1" cols="99%" style="width:99%;" rows="10" class="TextArea" id="KeyWord1"><%=listkeyword%></textarea>
          <br /></td>
      </tr>
    </table>
	<table id="exceltb" style="display:none"></table>
  </div>
  <div id="BoxBottom" style="width:96%;">
    <input type="submit" onClick="chkKwNums();" class="Button" name="button" id="buttonset" <%if not FkFun.CheckLimit("System22") then response.write "style=""display:block;""" else response.write "style=""display:none;"""%> value="保 存" />
    <input style="display:none;" type="button" onClick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
  </div>
</form>
<script language="javascript"> 
<!-- 
function chkdel(){
		$("#id-del").attr("disabled","disabled");
		var str='';
		$('input[name=chkID]').each(function(){
			if(this.checked){
				if(str==''){
					str=$(this).val();
				}
				else{
					str+=','+$(this).val();
				}
			}
		});
		if(str){
			DelIt("确定删除吗？",'KeyWord.asp?Type=3&iisid=<%=request("iisid")%>&id='+str,"BoxContent","KeyWord.asp?Type=1&iisid=<%=request("iisid")%>");
			//var para="KeyWord.asp?Type=1&iisid=<%=request("iisid")%>";
			//setTimeout(ShowBox(para), 2000);
		}
		$("#id-del").attr("disabled","");
}

function chkSet(eid,strUrl,strID){
	$("#"+eid).attr("disabled","disabled");
	if(strID==""){
		alert("请选择关键词！");
		$("#"+eid).attr("disabled","");
	}
	else{
		ShowBox(strUrl);
		ShowBox("KeyWord.asp?Type=1&searchK=<%=server.URLEncode(request("searchk"))%>&Page=<%=request("page")%>&iisid=<%=request("iisid")%>");
		//window.location.href=strUrl;
	}
}

/*var __sto = setTimeout; 　　
window.setTimeout = function(callback, timeout, param) { 　　   
	var args = Array.prototype.slice.call(arguments, 2); 　　   
    var _cb = function() { 　　                   
        callback.apply(null, args); 　　             
    } 
    _sto(_cb, timeout); 　　
}*/

function doaction(strUrl){
	ShowBox(strUrl);
	var para="KeyWord.asp?Type=1&searchK=<%=server.URLEncode(request("searchk"))%>&Page=<%=request("page")%>&iisid=<%=request("iisid")%>";
	setTimeout(ShowBox(para), 2000);
}

function chkKwNums(){
	var kwd,t,kwdLen,lmtKlen;
	kwd="";
	if($("#kw4").css("display")=='block'){
		t=1;
		lmtKlen=30;
		kwd=$("#KeyWord1").val().replace("||","|");
	}
	else if ($("#kw2").css("display")=='block'){
		t=0;
		lmtKlen=1000;
		kwd=$("#KeyWord").val().replace("||","|");
	}
	//var kwd=$("#KeyWord").val();
	kwdLen=kwd.split("|").length-1;
	if(kwdLen>lmtKlen){
		alert("关键词建议不超过"+lmtKlen+"个，当前关键词个数为"+ kwdLen +"个，请将关键词设置在"+lmtKlen+"个以内！");
	}
	else{
		Sends('KeyWordSet','KeyWord.asp?Type=2&t='+t,1,'file-shangwin.asp?filename=keyword&Viewstyle=1&iisid=<%=request("iisid")%>',0,0,'','');
	}
}

//$(document).ready(function(){
	//$(".getvisits").click()
//})

function viewYX(){
document.getElementById("kw1").style.display="none";
document.getElementById("kw2").style.display="none";
document.getElementById("kw4").style.display="none";
document.getElementById("kw3").style.display="block";
document.getElementById("buttonset").style.display="none";
} 

function viewciku(){
document.getElementById("kw1").style.display="none";
document.getElementById("kw3").style.display="none";
document.getElementById("kw4").style.display="none";
document.getElementById("kw2").style.display="block";
document.getElementById("buttonset").style.display="block";
} 
 
function viewciku1(){
document.getElementById("kw1").style.display="none";
document.getElementById("kw2").style.display="none";
document.getElementById("kw3").style.display="none";
document.getElementById("kw4").style.display="block";
document.getElementById("buttonset").style.display="block";
} 
 
function viewcilie(){
document.getElementById("kw1").style.display="block";
document.getElementById("kw2").style.display="none";
document.getElementById("kw3").style.display="none";
document.getElementById("kw4").style.display="none";
document.getElementById("buttonset").style.display="none";
}
//--> 
</script>
<%
End Sub

'==========================================
'函 数 名：KeyWordDel()
'作    用：删除关键词库
'参    数：
'==========================================
Sub KeyWordDel()
	dim strID,DataToSend,u,i,delKlist,result
	strID= Request("id")
	u=request.ServerVariables("HTTP_HOST")
	if strID<>"" then
		set rs=conn.execute("select [SVkeywords] from [keywordSV] where id in("&strID&")")
		if not rs.eof then
			i=0
			do while not rs.eof
				if u<>"localhost" and u<>"127.0.0.1" then
					i=i+1
					if i=1 then
						delKlist=rs("SVkeywords")
					else
						delKlist=delKlist&","&rs("SVkeywords")
					end if
				end if
			rs.movenext
			if rs.eof then exit do 
			loop
		end if
		rs.close
		set rs=nothing
		conn.execute("delete from [keywordSV] where id in ("&strID&")")
		response.Write "关键词删除成功！"
		if delKlist<>"" then
			DataToSend = "a=del&d="&replace(u,"www.","")&"&k="&vbsEscape(delKlist)
			result=PostHttpPage("http://www.qebang.cn/","http://win.qebang.net/web/pvr/upd_t.asp",DataToSend)
			'response.Write result
		end if
	end if
End Sub

'==========================================
'函 数 名：KeyWordSet()
'作    用：设置关键词内链
'参    数：
'==========================================
Sub KeyWordSetLink()
	dim rs,id,t,msgs,u
	id=Request("id")
	t=cint(Request("t"))
	u=request.ServerVariables("HTTP_HOST")
	if t=1 then
		set rs=conn.execute("select [Fk_Word_Id] from [Fk_Word] where [Fk_Word_Name]=(select [SVkeywords] from [keywordSV] where id ="&id&")")
		if not rs.eof then
			'conn.execute("delete * from [Fk_Word_Id] where Fk_Word_Id="&rs("Fk_Word_Id"))
		end if
		rs.close
	else
		set rs=conn.execute("select [Fk_Word_Id] from [Fk_Word] where [Fk_Word_Name]=(select [SVkeywords] from [keywordSV] where id ="&id&")")
		if rs.eof then
			'conn.execute("insert into [Fk_Word] (Fk_Word_Name) Fk_Word_Id="&rs("Fk_Word_Id"))
		end if
		rs.close
	end if
	set rs=nothing
End Sub

Function GetHttpPage(HttpUrl)
If IsNull(HttpUrl)=True Or Len(HttpUrl)<18 Or HttpUrl="$False$" Then
GetHttpPage="$False$"
Exit Function
End If
Dim Http
Set Http=server.createobject("MSXML2.XMLHTTP")
Http.open "GET",HttpUrl,true
Http.Send()
If Http.Readystate<>4 then
Set Http=Nothing
GetHttpPage="$False$"
Exit function
End if
GetHTTPPage=bytesToBSTR(Http.responseBody,"UTF-8")
Set Http=Nothing
If Err.number<>0 then
Err.Clear
End If
End Function

Function PostHttpPage(RefererUrl,PostUrl,PostData)
Dim xmlHttp
Dim RetStr
Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
xmlHttp.Open "POST", PostUrl, false
XmlHTTP.setRequestHeader "Content-Length",Len(PostData)
xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
xmlHttp.setRequestHeader "Referer", RefererUrl
xmlHttp.Send PostData
If Err.Number <> 0 Then
Set xmlHttp=Nothing
PostHttpPage = "$False$"
Exit Function
End If
PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"UTF-8")
Set xmlHttp = nothing
End Function

Function BytesToBstr(Body,Cset)
Dim Objstream
Set Objstream = Server.CreateObject("adodb.stream")
objstream.Type = 1
objstream.Mode =3
objstream.Open
objstream.Write body
objstream.Position = 0
objstream.Type = 2
objstream.Charset = Cset
BytesToBstr = objstream.ReadText
objstream.Close
set objstream = nothing
End Function
'==========================================
'函 数 名：KeyWordSet()
'作    用：设置关键词类型
'参    数：
'==========================================
Sub KeyWordSet()
	dim rs,id,t,msgs,u,r,blnRemote,DataToSend,i,delKlist,result
	id=Request("id")
	t=Request("t")
	u=request.ServerVariables("HTTP_HOST")
	r="0"
	set rs=conn.execute("select [id],[SVb1],[SVkeywords] from [keywordSV] where id in("&id&")")
	if not rs.eof then
		i=0
		do while not rs.eof
				i=i+1
				if i=1 then
					delKlist=rs("SVkeywords")
				else
					delKlist=delKlist&","&rs("SVkeywords")
				end if
		rs.movenext
		if rs.eof then exit do 
		loop
	end if
	rs.close
	set rs=Nothing
	'Call FKFso.CreateFile("delKlist.dat",delKlist)
	'response.end
	if delKlist<>"" then
		conn.execute("update [keywordSV] set [SVb1]='"&t&"' where id in("&id&")")
		if u<>"localhost" then
			'r=GetHttpPage("http://win.qebang.net/web/pvr/upd_t.asp")
			DataToSend = "a=upd&d="&replace(u,"www.","")&"&t="&t&"&reff="&u&"&k="&vbsEscape(delKlist)
			'Call FKFso.CreateFile("delKlist.dat",delKlist)
			result=PostHttpPage("http://www.qebang.cn/","http://win.qebang.net/web/pvr/upd_ts.asp",DataToSend)
			'response.Write result 
		end if
	end if
End Sub

'==========================================
'函 数 名：KeyWordImport()
'作    用：关键词导入
'参    数：
'==========================================
Sub KeyWordImport()
	dim rs,lst,msgs,u,DataToSend,i,delKlist,result,arrLst
	lst=cxarraynull(Request("lst"),"{0}")
	u=request.ServerVariables("HTTP_HOST")
	if instr(lst,"{0}")>0 then
		arrLst=split(lst,"{0}")
		for i=0 to ubound(arrLst)
			set rs=conn.execute("select [id] from [keywordSV] where SVkeywords='"&FilterText(arrLst(i))&"'")
			if rs.eof then
				conn.execute("insert into [keywordSV] (SVkeywords,SVb1) values('"&FilterText(arrLst(i))&"','3')")
			end if
			rs.close
			set rs=nothing
		next
	else
		set rs=conn.execute("select [id] from [keywordSV] where SVkeywords='"&FilterText(lst)&"'")
		if rs.eof then
			conn.execute("insert into [keywordSV] (SVkeywords,SVb1) values('"&FilterText(lst)&"','3')")
		end if
		rs.close
		set rs=nothing
	end if
	'Call FKFso.CreateFile("delKlist.dat",delKlist)
	'response.end
	if lst<>"" then
		if u<>"localhost" then
			'r=GetHttpPage("http://win.qebang.net/web/pvr/upd_t.asp")
			DataToSend = "a=upd&d="&replace(u,"www.","")&"&t=3&reff="&u&"&k="&vbsEscape(Filterkwd(lst))
			'Call FKFso.CreateFile("delKlist.dat",delKlist)
			result=PostHttpPage("http://www.qebang.cn/","http://win.qebang.net/web/pvr/upd_Import.asp",DataToSend)
			response.Write result 
		end if
	end if
End Sub

'==========================================
'函 数 名：KeyWordDo()
'作    用：设置关键词库
'参    数：
'==========================================
Sub KeyWordUpSVci()
	Dim SVcis,lst,arr,ik,SVk
	lst=DecodeURI(Request("lst"))
	If InStr(lst,"{1}")>0 Then
		arr=Split(lst,"{1}")
		For ik=0 To UBound(arr)
			If InStr(arr(ik),"{0}")>0 Then 
				SVcis=Split(arr(ik),"{0}")(1)
				SVk=Split(arr(ik),"{0}")(0)
		'response.write Split(arr(ik),"{0}")(1)
		'response.end
				'response.write "update [keywordSV] set SVci="&SVcis&" where SVkeywords='"&SVk&"'"&"<br/>"
				conn.execute("update [keywordSV] set SVci="&Split(arr(ik),"{0}")(1)&" where SVkeywords='"&Split(arr(ik),"{0}")(0)&"'")
			End If 
		Next 
	Else		
		If InStr(arr,"{0}")>0 Then 
			conn.execute("update [keywordSV] set SVci="&Split(arr,"{0}")(1)&" where SVkeywords='"&Split(arr,"{0}")(0)&"'")
		End If 
	End if
End Sub

Function DecodeURI(ByVal s)
    s = UnEscape(s)
    Dim cs : cs = "GBK"
    With New RegExp
        .Pattern = "^(?:[\x00-\x7f]|[\xfc-\xff][\x80-\xbf]{5}|[\xf8-\xfb][\x80-\xbf]{4}|[\xf0-\xf7][\x80-\xbf]{3}|[\xe0-\xef][\x80-\xbf]{2}|[\xc0-\xdf][\x80-\xbf])+$"
        If .Test(s) Then cs = "UTF-8"
    End With
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Mode = 3
        .Open
        .CharSet = "iso-8859-1"
        .WriteText s
        .Position = 0
        .CharSet = cs
        DecodeURI = .ReadText(-1)
        .Close
    End With
End Function

'========================
'函数名：cxarraynull
'作  用：关键词去重
'参  数：cxstr1:要去重的关键词串;cxstr2:分割符
'========================
function cxarraynull(cxstr1,cxstr2)
dim ss,sss,cxs,cc,m
if isarray(cxstr1) then
cxarraynull = ""
Exit Function
end if
if cxstr1 = "" or isempty(cxstr1) then
cxarraynull = ""
Exit Function
end if
do while instr(cxstr1,",,")>0
cxstr1=replace(cxstr1,",,",",")
loop
if right(cxstr1,1)="," then
cxstr1=left(cxstr1,len(cxstr1)-1)
end if
ss = split(cxstr1,cxstr2)
cxs = cxstr2&ss(0)&cxstr2
sss = cxs
for m = 0 to ubound(ss)
cc = cxstr2&ss(m)&cxstr2
if instr(sss,cc)=0 then
sss = sss&ss(m)&cxstr2
end if
next
cxarraynull = right(sss,len(sss) - len(cxstr2))
cxarraynull = left(cxarraynull,len(cxarraynull) - len(cxstr2))
end function

'==========================================
'函 数 名：KeyWordDo()
'作    用：设置关键词库
'参    数：
'==========================================
Sub KeyWordDo()
	dim t
	dim arrKwd,i,rs,u,results,DataToSend,lenkwd
	t=Request("t")
	u=request.ServerVariables("HTTP_HOST")
	if t=0 then
		KeyWord=ClearSG(cxarraynull(FilterText(Request("KeyWord")),"|"))
		Call FKFso.CreateFile("KeyWordC.dat",KeyWord)
		Response.Write("关键词库修改成功！")
	elseif t=1 then
		KeyWord=trim(Request("KeyWord1"))
		KeyWord=ClearRightDh(KeyWord)		
		arrKwd=split(KeyWord,",")
		lenkwd=ubound(arrKwd)+1
		if lenkwd<10 then
			Response.Write("关键词库修改失败:关键词不能小于10个！")
			response.end
		end if
		if lenkwd>30 then
			Response.Write("关键词库修改失败:关键词不能大于30个！")
			response.end
		end if
		KeyWord=cxarraynull(KeyWord,",")
		arrKwd=split(KeyWord,",")
		for i=0 to ubound(arrKwd)
			set rs=conn.execute("select SVkeywords from [keywordSV] where SVkeywords='"&FilterText(arrKwd(i))&"'")
			if rs.eof then
				conn.execute("insert into [keywordSV] (SVkeywords) values('"&FilterText(arrKwd(i))&"')")
			end if
		next
		rs.close
		set rs=nothing
		Response.Write("关键词库修改成功！")
		if KeyWord<>"" then
			if u<>"localhost" and u<>"127.0.0.1" then
				DataToSend = "a=upd&d="&replace(u,"www.","")&"&t=3&reff="&u&"&k="&vbsEscape(KeyWord)
				'Call FKFso.CreateFile("delKlist.dat",delKlist)
				results=PostHttpPage("http://www.qebang.cn/","http://win.qebang.net/web/pvr/upd_ts.asp",DataToSend)
				Response.Write(results)
			end if
		end if
	else
		Response.Write("关键词库修改失败！")
	end if
End Sub

function ClearRightDh(str)
	while(right(str,1)=",")
		str=left(str,len(str)-1)
	wend
	ClearRightDh=str
end function 

function ClearSG(str)
	while(right(str,1)="|")
		str=left(str,len(str)-1)
	wend
	while(left(str,1)="|")
		str=mid(str,2)
	wend
	ClearSG=str
end function 

'===================================== 
'过滤字符 
'===================================== 
Function FilterText(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
t0=Replace(t0,"||","|")'  
FilterText=t0 
End Function 

'===================================== 
'过滤字符 
'===================================== 
Function Filterkwd(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
Filterkwd=t0 
End Function 

Function Easp_Escape(ByVal str)
	Dim i,c,a,s : s = ""
	If isnull(str) Then Easp_Escape = "" : Exit Function
	For i = 1 To Len(str)
		c = Mid(str,i,1)
		a = ASCW(c)
		If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
			s = s & c
		ElseIf InStr("@*_+-./",c)>0 Then
			s = s & c
		ElseIf a>0 and a<16 Then
			s = s & "%0" & Hex(a)
		ElseIf a>=16 and a<256 Then
			s = s & "%" & Hex(a)
		Else
			s = s & "%u" & Hex(a)
		End If
	Next
	Easp_Escape = s
End Function

Sub Shuffle (ByRef arrInput)
    'declare local variables:
    Dim arrIndices, iSize, x
    Dim arrOriginal

    'calculate size of given array:
    iSize = UBound(arrInput)+1

    'build array of random indices:
    arrIndices = RandomNoDuplicates(0, iSize-1, iSize)

    'copy:
    arrOriginal = CopyArray(arrInput)

    'shuffle:
    For x=0 To UBound(arrIndices)
        arrInput(x) = arrOriginal(arrIndices(x))
    Next
End Sub

Function CopyArray (arr)
    Dim result(), x
    ReDim result(UBound(arr))
    For x=0 To UBound(arr)
        If IsObject(arr(x)) Then
            Set result(x) = arr(x)
        Else
            result(x) = arr(x)
        End If
    Next
    CopyArray = result
End Function

Function RandomNoDuplicates (iMin, iMax, iElements)
    'this function will return array with "iElements" elements, each of them is random
    'integer in the range "iMin"-"iMax", no duplicates.

    'make sure we won't have infinite loop:
    If (iMax-iMin+1)>iElements Then
        Exit Function
    End If

    'declare local variables:
    Dim RndArr(), x, curRand
    Dim iCount, arrValues()

    'build array of values:
    Redim arrValues(iMax-iMin)
    For x=iMin To iMax
        arrValues(x-iMin) = x
    Next

    'initialize array to return:
    Redim RndArr(iElements-1)

    'reset:
    For x=0 To UBound(RndArr)
        RndArr(x) = iMin-1
    Next

    'initialize random numbers generator engine:
    Randomize
    iCount=0

    'loop until the array is full:
    Do Until iCount>=iElements
        'create new random number:
        curRand = arrValues(CLng((Rnd*(iElements-1))+1)-1)

        'check if already has duplicate, put it in array if not
        If Not(InArray(RndArr, curRand)) Then
            RndArr(iCount)=curRand
            iCount=iCount+1
        End If

        'maybe user gave up by now...
        If Not(Response.IsClientConnected) Then
            Exit Function
        End If
    Loop

    'assign the array as return value of the function:
    RandomNoDuplicates = RndArr
End Function

Function InArray(arr, val)
    Dim x
    InArray=True
    For x=0 To UBound(arr)
        If arr(x)=val Then
            Exit Function
        End If
    Next
    InArray=False
End Function


'************************* 
'函数:UBoundStrToArr 
'作用:检测原字符串转换为数组的最大下标值 
'参数:cCheckStr(需要检测的字符串) 
' cUBoundArr(生成数组的最大下标值) 
' cSpaceStr(间隔字符串) 
'返回:数组的最大下标值 
'************************ 
Public Function UBoundStrToArr(ByVal cCheckStr,ByVal cUBoundArr,ByVal cSpaceStr) 
On Error Resume Next

If Instr(cCheckStr,cSpaceStr)=0 Then 
UBoundStrToArr=cUBoundArr 
Exit Function 
End If 
Dim TempSpaceStr,UBoundValue 
TempSpaceStr=Mid(cCheckStr,Len(cCheckStr)-Len(cSpaceStr)+1) '获取字符串右侧间隔字符 
If TempSpaceStr=cSpaceStr Then '如果字符串最右侧存在间隔字符,则下标值需要-1 
UBoundValue=cUBoundArr-1 
Else 
UBoundValue=cUBoundArr 
End If 
UBoundStrToArr=UBoundValue 
End Function 


'********查询关键词在数据库某个表某个字段中出现的次数**********
Function Chakeywordci(Keywordsrt,Keywordslei)
 dim RSC1,RSC2,RSC3,SqlChastr
 RSC1=0:RSC2=0:RSC3=0
 select case Keywordslei
 case 2
	SqlChastr="Select Fk_Article_Title,Fk_Article_Keyword From Fk_Article Where Fk_Article_Title Like '%%"&Keywordsrt&"%%' or Fk_Article_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC1=Rs.RecordCount
	Rs.Close
	
	SqlChastr="Select Fk_Product_Title,Fk_Product_Keyword From Fk_Product Where Fk_Product_Title Like '%%"&Keywordsrt&"%%' or Fk_Product_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC2=Rs.RecordCount
	Rs.Close
case 1 	
	SqlChastr="Select Fk_Module_Keyword From Fk_Module Where Fk_Module_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC3=Rs.RecordCount
	Rs.Close
end select
	Chakeywordci=RSC1+RSC2+RSC3
End function

'****查询关键词是否有做内链**********
Function ChakeywordNLink(Keywordsrt)
	dim SqlChastr
	SqlChastr="Select Fk_Word_Name From Fk_Word Where Fk_Word_Name Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
		if not Rs.eof then
			ChakeywordNLink=1
		else
			ChakeywordNLink=0
		end if
	Rs.Close
End function

'****查询关键词在数据库中的排名记录**********
Function Chanowpaiming(Keywordsrt)
	dim SqlChastr
	SqlChastr="Select SVkeywords,SVpaiming From [keywordSV] Where SVkeywords='"&Keywordsrt&"' "
	Rs.Open SqlChastr,Conn,1,1
		if not Rs.eof then
			Chanowpaiming=Rs("SVpaiming")
		else
			Chanowpaiming="查询排名"
		end if
	Rs.Close
End function

'****查询strA中strB出现的次数**********
Function strCount(strA,strB)
lngA = Len(strA)
lngB = Len(strB)
lngC = Len(Replace(strA, strB, ""))
strCount = (lngA - lngC) / lngB
End Function

Function vbsEscape(str)
    dim i,s,c,a
    s=""
    For i=1 to Len(str)
        c=Mid(str,i,1)
        a=ASCW(c)
        If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
            s = s & c
        ElseIf InStr("@*_+-./",c)>0 Then
            s = s & c
        ElseIf a>0 and a<16 Then
            s = s & "%0" & Hex(a)
        ElseIf a>=16 and a<256 Then
            s = s & "%" & Hex(a)
        Else
            s = s & "%u" & Hex(a)
        End If
    Next
    vbsEscape = s
End Function

'页面结束
dim SqlChastr
KeyWord=""
SqlChastr=""
listkeyword=""
Newstr=""
'set Chakeywordci=nothing
set rs=nothing
%>
<!--#Include File="../Code.asp"-->
