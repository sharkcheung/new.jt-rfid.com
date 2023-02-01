<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<title></title>
<SCRIPT type="text/javascript" src="/js/jquery-1.7.2.min.js"></SCRIPT>

<script type="text/javascript" src="/Js/jquery.form.min.js"></script>
<script type="text/javascript" src="layer/layer.js"></script>
<script type="text/javascript" src="Js/function.js"></script>
<link href="/admin-new/css/seo.css" rel="stylesheet" type="text/css" />
<script>
window.onerror = function(sMsg, sUrl, sLine) {
           var strlog="错误信息：" + sMsg + "\r\n";
           strlog+="出错文件：" + sUrl + "\r\n";
           strlog+="出错行号：" + sLine + "\r\n";
           // alert(strlog);
           return true;
    }
</script>

</head>

<body>
<div class="menunav">
      	<div class="center">
                      <a href="http://admin.qbt.qebang.com/index.php/home/operationManage/index?<%=tokenpara%>">运营管理信息</a><span></span>
                      <a class="active" href="http://admin.qbt.qebang.com/index.php/home/operationManage/seasonOperation?<%=tokenpara%>">PC端运营</a><span></span>
                      <a href="http://admin.qbt.qebang.com/index.php/home/operationManage/weChat?<%=tokenpara%>">移动端运营</a><span></span>
                      <a href="http://admin.qbt.qebang.com/index.php/home/operationManage/platFormAccount?<%=tokenpara%>">平台运营帐号</a><span></span>
                      <a href="http://admin.qbt.qebang.com/index.php/home/activities/index?<%=tokenpara%>">线下活动</a><span></span>
                      <a href="http://admin.qbt.qebang.com/index.php/home/operationManage/hlwsw?<%=tokenpara%>">运营学习</a>
         </div>
</div>