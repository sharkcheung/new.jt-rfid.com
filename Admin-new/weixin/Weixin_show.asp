<!--#Include File="../../Include.asp"-->
<%
response.charset="utf-8"
session.codepage=65001
dim wx_nid
dim imgText_Title,imgText_addtime,imgText_Pic,imgText_Content
wx_nid=request("id")
if not isnumeric(wx_nid) then
	response.write "企帮(www.qebang.cn)"
	response.end	
end if
if wx_nid<=0 then
	response.write "企帮(www.qebang.cn)"
	response.end
end if
set rs=Conn.execute("select imgText_Title,imgText_addtime,imgText_Pic,imgText_Content from weixin_imageText where id="&wx_nid)
if not rs.eof then
	imgText_Title=rs("imgText_Title")
	imgText_addtime=rs("imgText_addtime")
	imgText_Pic=rs("imgText_Pic")
	imgText_Content=rs("imgText_Content")
%>

<!DOCTYPE HTML>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=100%, initial-scale=1.0, user-scalable=no"/>
<meta content="telephone=no" name="format-detection" />
<title><%=imgText_Title%></title>
<link rel="dns-prefetch" href="//mat1.gtimg.com">
<link rel="dns-prefetch" href="//imgcache.gtimg.cn">
<link href="http://mat1.gtimg.com/www/cssn/newsapp/wxnews/wechat20141230.css" rel="stylesheet" type="text/css">
<style>
#borderLogo .logoImg{
background-image:url(<%=SiteLogo%>);
background-size: 84px 34px;
}
body{overflow:hidden;}
#mcover {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background: rgba(0, 0, 0, 0.7);
    display: none;
    z-index: 20000;
 }
 #mcover img {
    position: fixed;
    right: 18px;
    top: 5px;
    width: 260px!important;
    height: 180px!important;
    z-index: 20001;
 }
 .text {
    margin: 15px 0;
    font-size: 14px;
    word-wrap: break-word;
    color: #727272;
 }
 #mess_share {
    margin: 15px 0;
    display: block;
 }


			 #share_1 {
				float: left;
				width: 48%;
				display: block;
			 }
			 #share_2 {
				float: right;
				width: 48%;
				display: block;
			 }
			 .clr {
				display: block;
				clear: both;
				height: 0;
				overflow: hidden;
			 }
			 .button2 {
				font-size: 16px;
				padding: 8px 0;
				border: 1px solid #adadab;
				color: #000000;
				background-color: #e8e8e8;
				background-image: linear-gradient(to top, #dbdbdb, #f4f4f4);
				box-shadow: 0 1px 1px rgba(0, 0, 0, 0.45), inset 0 1px 1px #efefef;
				text-shadow: 0.5px 0.5px 1px #fff;
				text-align: center;
				border-radius: 3px;
				width: 100%;
			 }


 #mess_share img {
    width: 22px!important;
    height: 22px!important;
    vertical-align: top;
    border: 0;
 }
</style>
<script src="http://libs.baidu.com/jquery/1.9.0/jquery.js"></script>
<script type="text/javascript"> // 
function weChat(){
$("#mcover").css("display","none");  // 点击弹出层，弹出层消失
}
function button1(){
$("#mcover").css("display","block")    // 分享给好友按钮触动函数
}
function button2(){
$("#mcover").css("display","block")  // 分享给好友圈按钮触动函数
}
</script>
</head> 

<body>
<div id="mcover" onclick="weChat()" style="display:none;">
<img src="https://mmbiz.qlogo.cn/mmbiz/vV3bdMHsLIjY2s0npKT0FaJ6iaC1MaiciakM61zfqBsNjYH14ovUG145GEuwMPafiaPjh5drSaAg8DMTic3a2I3icbLg/0" />
</div>
<div id="borderLogo"><div class="logoImg"></div></div>
<div id="content" class="main fontSize2">
<p class="title" align="left"><%=imgText_Title%></p>
<span class="src"><%=imgText_addtime%>&nbsp;<%=SiteName%></span>
<p class="text"><div class="preLoad" style="width:100%; min-height:"173px"><div class="img"><img src="<%=imgText_Pic%>" preview-src="<%=imgText_Pic%>" style="width:100%;display:block;"/></div><div class="tip"><%=SiteName%></div></div></p><p class="text"><%=imgText_Content%></p>

<br>
<div class="text" id="content">
	<div id="mess_share">
		<div id="share_1">
			<button class="button2" onclick="button1()">
				<img src="https://mmbiz.qlogo.cn/mmbiz/vV3bdMHsLIjY2s0npKT0FaJ6iaC1MaiciakIHMqX6tb7127kicbBd5vIZcey4wenREiaEe8YXshOWpFcIser6AgbsEA/0" width="64" height="64" />
				 发送给朋友
			</button>
		</div>
		<div id="share_2">
			<button class="button2" onclick="button2()">
				<img src="https://mmbiz.qlogo.cn/mmbiz/vV3bdMHsLIjY2s0npKT0FaJ6iaC1MaiciakERrBO1bHKDDzxiakMd4m2H1mmib1ShpekZ8RZm5ECazcDqF96c5wcl2w/0" width="64" height="64" />
				 分享到朋友圈
			</button>
		</div>
		<div class="clr"></div>
	</div>
</div>   
</div>   
        <script type="text/javascript">
        var ISWP = !!(navigator.userAgent.match(/Windows\sPhone/i));
        var sw = 0;

        if (ISWP){
            var profile = document.getElementById('post-user');
            if (profile){
                profile.setAttribute("href", "weixin://profile/gh_295d2b07ed62");
            }
        }
        var tid = "";
        var aid = "";
        var uin = "";
        var key = "";
        var biz = "MjM5NjU5NDkzMg==";
		var networkType;
		
        var cookie = {
            get: function(name){
                if( name=='' ){
                    return '';
                }
                var reg = new RegExp(name+'=([^;]*)');
                var res = document.cookie.match(reg);
                return (res && res[1]) || '';
            },
            set: function(name, value){
                var now = new Date();
                    now.setDate(now.getDate() + 1);
                var exp = now.toGMTString();
                document.cookie = name + '=' + value + ';expires=' + exp;
                return true;
            }
        };

        function hash(str){
            var hash = 5381;
            for(var i=0; i<str.length; i++){
                hash = ((hash<<5) + hash) + str.charCodeAt(i);
                hash &= 0x7fffffff;
            }
            return hash;
        }

        function trim(str){
            return str.replace(/^\s*|\s*$/g,'');
        }


        function parseParams(str) {
            if( !str ) return {};

            var arr = str.split('&'), obj = {}, item = '';
            for( var i=0,l=arr.length; i<l; i++ ){
                item = arr[i].split('=');
                obj[item[0]] = item[1];
            }
            return obj;
        }

        function htmlDecode(str){
            return str
                  .replace(/&#39;/g, '\'')
                  .replace(/<br\s*(\/)?\s*>/g, '\n')
                  .replace(/&nbsp;/g, ' ')
                  .replace(/&lt;/g, '<')
                  .replace(/&gt;/g, '>')
                  .replace(/&quot;/g, '"')
                  .replace(/&amp;/g, '&');
        }

        // 记住阅读位置
        (function(){
            var timeout = null;
            var val = 0;
            var url = "http://mp.weixin.qq.com/s?__biz=MjM5NjU5NDkzMg==&mid=203437929&idx=1&sn=e9cf519ae27b5a65b5482f937fd74962#rd".split('?').pop();
            var key = hash(url);
            /*
            var params = parseParams( url );
            var biz = params['__biz'].replace(/=/g, '#');
            var key = biz + params['appmsgid'] + params['itemidx'];
            */

            if( window.addEventListener ){
                window.addEventListener('load', function(){
                    val = cookie.get(key);
                    window.scrollTo(0, val);
                }, false);

                window.addEventListener('unload', function(){
                    cookie.set(key,val);
                    // 上报页面停留时间
                }, false);

                window.addEventListener('scroll', function(){
                    clearTimeout(timeout);
                    timeout = setTimeout(function(){
                        val = window.pageYOffset;
                    },500);
                }, false);

                document.addEventListener('touchmove', function(){
                    clearTimeout(timeout);
                    timeout = setTimeout(function(){
                        val = window.pageYOffset;
                    },500);
                }, false);
            }else if(window.attachEvent){
                window.attachEvent('load', function(){
                    val = cookie.get(key);
                    window.scrollTo(0, val);
                }, false);

                window.attachEvent('unload', function(){
                    cookie.set(key,val);
                    // 上报页面停留时间
                }, false);

                window.attachEvent('scroll', function(){
                    clearTimeout(timeout);
                    timeout = setTimeout(function(){
                        val = window.pageYOffset;
                    },500);
                }, false);

                document.attachEvent('touchmove', function(){
                    clearTimeout(timeout);
                    timeout = setTimeout(function(){
                        val = window.pageYOffset;
                    },500);
                }, false);
            }
        })();

    
        //弹出框中图片的切换
        (function(){
            var imgsSrc = [];
            function reviewImage(src) {
                if (typeof window.WeixinJSBridge != 'undefined') {
                    WeixinJSBridge.invoke('imagePreview', {
                        'current' : src,
                        'urls' : imgsSrc
                    });
                }
            }
            function onImgLoad() {
                var imgs = document.getElementById("img-content");
                imgs = imgs ? imgs.getElementsByTagName("img") : [];
                for( var i=0,l=imgs.length; i<l; i++ ){//忽略第一张图 是提前加载的loading图而已
                    var img = imgs.item(i);
                    var src = img.getAttribute('data-src') || img.getAttribute('src');
                    if( src ){
                        imgsSrc.push(src);
                        (function(src){
                            if (img.addEventListener){
                                img.addEventListener('click', function(){
                                    reviewImage(src);
                                });
                            }else if(img.attachEvent){
                                img.attachEvent('click', function(){
                                    reviewImage(src);
                                });
                            }
                        })(src);
                    }
                }
            }
            if( window.addEventListener ){
                window.addEventListener('load', onImgLoad, false);
            }else if(window.attachEvent){
                window.attachEvent('load', onImgLoad);
                window.attachEvent('onload', onImgLoad);
            }
        })();

        var has_click = {};
        function gdt_click(type, url, rl, apurl, traceid, group_id){
            if (has_click[traceid]){return;}
            has_click[traceid] = true;
            var loading = document.getElementById("loading_" + traceid);
            if (loading){
                loading.style.display = "inline";
            }
            var b = (+new Date())
        }

        // 图片延迟加载
        (function(){
            var timer  = null;
            var innerHeight = (window.innerHeight||document.documentElement.clientHeight);
            var height = innerHeight + 40;
            var images = [];
            function detect(){
                var scrollTop = (window.pageYOffset||document.documentElement.scrollTop) - 20;
                for( var i=0,l=images.length; i<l; i++ ){
                    var img = images[i];
                    var offsetTop = img.el.offsetTop;
                    if( !img.show && scrollTop < offsetTop+img.height && scrollTop+height > offsetTop ){
                        img.el.setAttribute('src', img.src);
                        img.show = true;
                    }
                    if (ISWP && (img.el.width*1 > sw)){//兼容WP
                        img.el.width = sw;
                    }
                }
            }

            var ping_apurl = false;
            function onLoad(){
                var imageEls = document.getElementsByTagName('img');
                var pcd = document.getElementById("page-content");
                if (pcd.currentStyle){
                    sw = pcd.currentStyle.width;
                }else if (typeof getComputedStyle != "undefined"){
                    sw = getComputedStyle(pcd).width;
                }
                sw = 1*(sw.replace("px", ""));
                for( var i=0,l=imageEls.length; i<l; i++ ){
                    var img = imageEls.item(i);
                    if(!img.getAttribute('data-src') ) continue;
                    images.push({
                        el     : img,
                        src    : img.getAttribute('data-src'),
                        height : img.offsetHeight,
                        show   : false
                    });
                }
                detect();
				// @cunjinli
            }
            if( window.addEventListener ){
                window.addEventListener('load', onLoad, false);
            }
            else {
                window.attachEvent('onload', onLoad);
            }
        })();
		// pic load report
		//@cunjinli
		function addEvent(elem, type, func) {
			if (window.addEventListener ) {
				elem.addEventListener(type, func, false);
			} else if (window.attachEvent) {
				elem.attachEvent("on" + type, (function(elem) {
					return function(e){ func.call(elem, e); };
				})(elem));
			} else {
				elem["on" + type] = func;
			}
		}
        
    </script>
	
</body>
</html>
<%
else
	response.write "企帮(www.qebang.cn)"
end if
rs.close
if err then
	err.clear
	response.write "企帮(www.qebang.cn)"
end if%>
<!--#Include File="../../Code.asp"-->