<!--#Include File="../../inc/qb_safe3.asp"-->
<!--#Include File="../../inc/conn.asp"-->
<!--#Include File="../../class/Cls_DB.asp"-->

<%
resposne.charset="utf-8"
session.codepage=65001
dim wx_nid
dim FKDB
Dim Conn,Rs,SiteData,SiteDir,SiteDBDir
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
Set FKDB=New Cls_DB
Call FKDB.DB_Open()
set rs=Conn.execute("select imgText_Title,imgText_addtime,imgText_Pic,imgText_Content from weixin_imageText where id="&wx_nid)
if not rs.eof then
	imgText_Title=rs("imgText_Title")
	imgText_addtime=rs("imgText_addtime")
	imgText_Pic=rs("imgText_Pic")
	imgText_Content=rs("imgText_Content")
%>
<!DOCTYPE html>
<html> 
<head> 
    <meta http-equiv="Content-Type" content="text/html;charset=utf-8">
    <title><%=imgText_Title%></title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=0" />
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black">
    <meta name="format-detection" content="telephone=no">
    <link rel="stylesheet" type="text/css" href="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/style/client-page1baa9e.css"/>
    <!--[if lt IE 9]>
    <link rel="stylesheet" type="text/css" href="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/style/pc-page1b2f8d.css"/>
    <![endif]-->
    <link media="screen and (min-width:1000px)" rel="stylesheet" type="text/css" href="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/style/pc-page1b2f8d.css"/>
    <style>
        body{ -webkit-touch-callout: none; -webkit-text-size-adjust: none; }
    </style>
    <style>
        #nickname{overflow:hidden;white-space:nowrap;text-overflow:ellipsis;max-width:90%;}
                ol,ul{list-style-position:inside;}
        #activity-detail .page-content .text{font-size:16px;}
            </style>
</head> 

<body id="activity-detail">
            <img width="12px" style="position: absolute;top:-1000px;" src="http://res.wx.qq.com/mmbizwap/zh_CN/htmledition/images/ico_loading1984f1.gif">
        
        <div class="page-bizinfo">
            <div class="header">
            <h1 id="activity-name"><%=imgText_Title%></h1>
            <p class="activity-info">
                <span id="post-date" class="activity-meta no-extra"><%=imgText_addtime%></span>
                            </p>
            </div>
        </div>
        
        <div id="page-content" class="page-content">
            <div id="img-content">
                                    
                        <div class="media" id="media">
            <img onerror="this.parentNode.removeChild(this)" src="<%=imgText_Pic%>" />
            </div>
                        <div class="text">
							<%=imgText_Content%>
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