<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="./classes/MediPayRequestHandler.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk">
<title>�Ƹ�ͨ�н鵣��֧������ʾ��</title>
</head>
<body>
<%
'---------------------------------------------------------
'�Ƹ�ͨ�н鵣��֧������ʾ�����̻����մ��ĵ����п�������
'---------------------------------------------------------

Dim strNow
Dim randNumber

'14λʱ�䣨YYYYmmddHHMMss)
strNow = getStrNow()

'4λ�����
randNumber = getStrRandNumber(1000,9999)

Dim key
Dim chnid
Dim seller

'ƽ̨����Կ
key = "85e5ffb11e1caa8561b953a7e27a547c"

'ƽ̨���ʺ�
chnid = "1211461601"

'����
seller = "wosixiehongfu@126.com"

Dim mch_desc
Dim mch_name
Dim mch_price
Dim mch_returl
Dim mch_vno
Dim show_url
Dim transport_desc
Dim transport_fee

'����˵��
mch_desc = "����˵��"

'��Ʒ����
mch_name = "��Ʒ����"

'��Ʒ�ܼۣ���λΪ��
mch_price = "1"

'�ص�֪ͨURL
mch_returl = "http://localhost/tenpay/mch_returl.asp"

'�̼ҵĶ�����
mch_vno = strNow & randNumber

'֧������̻�֧�����չʾҳ��
show_url = "http://localhost/tenpay/show_url.asp"

'������˾��������ʽ˵��
transport_desc = ""

'������֧������������,�Է�Ϊ��λ
transport_fee = ""


'����֧���������
Dim reqHandler
Set reqHandler = new MediPayRequestHandler

'��ʼ��
reqHandler.init()

'������Կ
reqHandler.setKey(key)

'-----------------------------
'����֧������
'-----------------------------
reqHandler.setParameter "chnid", chnid						'ƽ̨���ʺ�
reqHandler.setParameter "encode_type", "1"					'�������� 1:gbk 2:utf-8
reqHandler.setParameter "mch_desc", mch_desc				'����˵��
reqHandler.setParameter "mch_name", mch_name				'��Ʒ����
reqHandler.setParameter "mch_price", mch_price				'��Ʒ�ܼۣ���λΪ��
reqHandler.setParameter "mch_returl", mch_returl			'�ص�֪ͨURL
reqHandler.setParameter "mch_type", "1"						'�������ͣ�1��ʵ�ｻ�ף�2�����⽻��
reqHandler.setParameter "mch_vno", mch_vno					'�̼ҵĶ�����
reqHandler.setParameter "need_buyerinfo", "2"				'�Ƿ���Ҫ�ڲƸ�ͨ�������Ϣ��1����Ҫ��2������Ҫ��
reqHandler.setParameter "seller", seller					'���ҲƸ�ͨ�ʺ�
reqHandler.setParameter "show_url",	show_url				'֧������̻�֧�����չʾҳ��
reqHandler.setParameter "transport_desc", transport_desc	'������˾��������ʽ˵��
reqHandler.setParameter "transport_fee", transport_fee		'������֧������������


'�����URL
Dim reqUrl
reqUrl = reqHandler.getRequestURL()

'debug��Ϣ
'Dim debugInfo
'debugInfo = reqHandler.getDebugInfo()

'Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")

'Response.Write("<br/>reqUrl" & reqUrl & "<br/>")

'reqHandler.doSend()


%>
<br/><a href="<%=reqUrl%>" target="_blank">�Ƹ�֧ͨ��</a>
</body>
</html>