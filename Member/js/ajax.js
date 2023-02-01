function shopcart()
//命令处理函数
{
//var password=encodeURIComponent(document.getElementById("shopcart").value); 
//从文本框中取得密码文本并进行编码以保证不会出现乱码
//var message="password="+password; 
//这是向ASP文件投递的数据必须是xxx=xxx格式
var theHttpRequest=getHttpObject(); 
//创建一个XMLHttpRequest
theHttpRequest.onreadystatechange=function ()
{
backAJAX();
}; 
//设定当asp文件返回数据时的处理函数为backAJAX()
theHttpRequest.open("POST","/member/shopcart.asp",true); 
//以POST方式打开XMLHttpRequest,投递地址为"check.asp",true表示以异步方式打开
theHttpRequest.setRequestHeader("Content-Type","application/x-www-form-urlencoded"); 
//我也不是很清楚，但有用
theHttpRequest.send(); 
//投递信息！


function getHttpObject()
//这是创建的函数XMLHttpReques
{
var objType=false; 
try
{
objType=new ActiveXObject('Msxml2.XMLHTTP'); 
//在较新的ie浏览器中这样创建
}catch(e)
{
try
{
objType=new ActiveXObject('Microsoft.XMLHTTP'); 
//在旧ie中这样创建
}catch(e)
{
objType=new XMLHttpRequest(); 
//如果浏览器是mozilla就这样创建
}
}
return objType; 
//返回XMLHttpReques对象
}





function backAJAX()
//处理asp返回数据的函数
{
if(theHttpRequest.readyState==4)
//4代表已经准备好
{
if(theHttpRequest.status==200)
//200代表一切正常
{
document.getElementById("shopcart").innerHTML=theHttpRequest.responseText; 
//将返回的的信息写入文档
}else 
{
document.getElementById("shopcart").innerHTML="<p>错误信息: "+theHttpRequest.statusText+"</p>"; 
//如果出现错误就输出错误信息
}
}
}
} 

