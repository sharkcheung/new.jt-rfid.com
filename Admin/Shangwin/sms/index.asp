<%
username=request("smsid")
password=request("smspw")
if username="" then username="52261"
if password="" then password="52261"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�����û���¼</title>
<link href="http://www.chinasms.com.cn/sms.css" rel="stylesheet" type="text/css" /> 
</head>
<script language=javascript>
setTimeout("document.form1.submit()",1000)//1����ύ
</script>

<body >

<div class="user user_login" style="margin-top:30px; display:none;">
    <div class="user_bt" ><img src="http://www.chinasms.com.cn/img/userbt.gif" /></div>
    <p>
<form name="form1" method="post" action="http://www.chinasms.com.cn/login.php" target="_self">
 <label>�û�����</label>
      <input type="text" class="user_int" name="username" value="<%=username%>" size="10"  />
    </p>
    <p>
      <label>�ܡ��룺</label>
      <input type="password" class="user_int" name="password" value="<%=password%>" size="10" />
    </p>
    <p>
      &nbsp;<select name="time" id="time" size="1">
        <option value="1800">��Ч��</option>
        <option value="3600">һʱ</option>
        <option value="10800">��ʱ</option>
        <option value="43200">12ʱ</option>
        <option value="86400" selected>24ʱ</option>
      </select>
    </p>
    <p><label>&nbsp;</label></p>
  </form>

</div>
<div style="width:50%;height:50px; border:1px #99CCFF solid;background:#EEF5FD; line-height:40px; color:#006699; margin-top:100px;">���ڽ��뾫׼Ӫ��-����ƽ̨...</div>
</body>
</html>
