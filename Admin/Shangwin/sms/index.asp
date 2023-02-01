<%
username=request("smsid")
password=request("smspw")
if username="" then username="52261"
if password="" then password="52261"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>短信用户登录</title>
<link href="http://www.chinasms.com.cn/sms.css" rel="stylesheet" type="text/css" /> 
</head>
<script language=javascript>
setTimeout("document.form1.submit()",1000)//1秒后提交
</script>

<body >

<div class="user user_login" style="margin-top:30px; display:none;">
    <div class="user_bt" ><img src="http://www.chinasms.com.cn/img/userbt.gif" /></div>
    <p>
<form name="form1" method="post" action="http://www.chinasms.com.cn/login.php" target="_self">
 <label>用户名：</label>
      <input type="text" class="user_int" name="username" value="<%=username%>" size="10"  />
    </p>
    <p>
      <label>密　码：</label>
      <input type="password" class="user_int" name="password" value="<%=password%>" size="10" />
    </p>
    <p>
      &nbsp;<select name="time" id="time" size="1">
        <option value="1800">有效期</option>
        <option value="3600">一时</option>
        <option value="10800">三时</option>
        <option value="43200">12时</option>
        <option value="86400" selected>24时</option>
      </select>
    </p>
    <p><label>&nbsp;</label></p>
  </form>

</div>
<div style="width:50%;height:50px; border:1px #99CCFF solid;background:#EEF5FD; line-height:40px; color:#006699; margin-top:100px;">正在进入精准营销-短信平台...</div>
</body>
</html>
