<!--#include file="md5.asp"-->
<!--#include file = "../admin/admin_conn.asp" -->
<!-- #include file="../config.asp" -->
<%
payerName="dfsdfsdf"


'*
'* @Description: ��Ǯ�����֧�����ؽӿڷ���
'* @Copyright (c) �Ϻ���Ǯ��Ϣ�������޹�˾
'* @version 2.0
'*

'����������˻���
''���¼��Ǯϵͳ��ȡ�û���ţ��û���ź��01��Ϊ����������˻��š�
merchantAcctId="1001153656201"

'�����������Կ
''���ִ�Сд.�����Ǯ��ϵ��ȡ
key="ZUZNJB8MF63GA83J"

'�ַ���.�̶�ѡ��ֵ����Ϊ�ա�
''ֻ��ѡ��1��2��3.
''1����UTF-8; 2����GBK; 3����gb2312
''Ĭ��ֵΪ1
inputCharset="3"

'����֧�������ҳ���ַ.��[bgUrl]����ͬʱΪ�ա������Ǿ��Ե�ַ��
''���[bgUrl]Ϊ�գ���Ǯ��֧�����Post��[pageUrl]��Ӧ�ĵ�ַ��
''���[bgUrl]��Ϊ�գ�����[bgUrl]ҳ��ָ����<redirecturl>��ַ��Ϊ�գ���ת��<redirecturl>��Ӧ�ĵ�ַ
pageUrl=""

'����������֧������ĺ�̨��ַ.��[pageUrl]����ͬʱΪ�ա������Ǿ��Ե�ַ��
''��Ǯͨ�����������ӵķ�ʽ�����׽�����͵�[bgUrl]��Ӧ��ҳ���ַ�����̻�������ɺ������<result>���Ϊ1��ҳ���ת��<redirecturl>��Ӧ�ĵ�ַ��
''�����Ǯδ���յ�<redirecturl>��Ӧ�ĵ�ַ����Ǯ����֧�����post��[pageUrl]��Ӧ��ҳ�档
bgUrl="http://jhrj.qebang.cn/kqpay/receive.asp"
	
'���ذ汾.�̶�ֵ
''��Ǯ����ݰ汾�������ö�Ӧ�Ľӿڴ������
''������汾�Ź̶�Ϊv2.0
version="v2.0"

'��������.�̶�ѡ��ֵ��
''ֻ��ѡ��1��2��3
''1�������ģ�2����Ӣ��
''Ĭ��ֵΪ1
language="1"

'ǩ������.�̶�ֵ
''1����MD5ǩ��
''��ǰ�汾�̶�Ϊ1
signType="1"
   
'֧��������
''��Ϊ���Ļ�Ӣ���ַ�
payerName=payerName

'֧������ϵ��ʽ����.�̶�ѡ��ֵ
''ֻ��ѡ��1
''1����Email
payerContactType="1"

'֧������ϵ��ʽ
''ֻ��ѡ��Email���ֻ���
payerContact=""

'�̻�������
''����ĸ�����֡���[-][_]���
orderId=111111111111111

'�������
''�Է�Ϊ��λ����������������
''�ȷ�2������0.02Ԫ
orderAmount=1
	
'�����ύʱ��
''14λ���֡���[4λ]��[2λ]��[2λ]ʱ[2λ]��[2λ]��[2λ]
''�磻20080101010101
orderTime=getDateStr()

'��Ʒ����
''��Ϊ���Ļ�Ӣ���ַ�
productName="�����֤ȯ����ϵͳ"

'��Ʒ����
''��Ϊ�գ��ǿ�ʱ����Ϊ����
productNum=1

'��Ʒ����
''��Ϊ�ַ���������
productId=""

'��Ʒ����
productDesc=""
	
'��չ�ֶ�1
''��֧��������ԭ�����ظ��̻�
ext1=""

'��չ�ֶ�2
''��֧��������ԭ�����ظ��̻�
ext2=""
	
'֧����ʽ.�̶�ѡ��ֵ
''ֻ��ѡ��00��10��11��12��13��14
''00�����֧��������֧��ҳ����ʾ��Ǯ֧�ֵĸ���֧����ʽ���Ƽ�ʹ�ã�10�����п�֧��������֧��ҳ��ֻ��ʾ���п�֧����.11���绰����֧��������֧��ҳ��ֻ��ʾ�绰֧����.12����Ǯ�˻�֧��������֧��ҳ��ֻ��ʾ��Ǯ�˻�֧����.13������֧��������֧��ҳ��ֻ��ʾ����֧����ʽ��.14��B2B֧��������֧��ҳ��ֻ��ʾB2B֧��������Ҫ���Ǯ���뿪ͨ����ʹ�ã�
payType="00"

'���д���
''ʵ��ֱ����ת������ҳ��ȥ֧��,ֻ��payType=10ʱ�������ò���
''�������μ� �ӿ��ĵ����д����б�
bankId=""

'ͬһ������ֹ�ظ��ύ��־
''�̶�ѡ��ֵ�� 1��0
''1����ͬһ������ֻ�����ύ1�Σ�0��ʾͬһ��������û��֧���ɹ���ǰ���¿��ظ��ύ��Ρ�Ĭ��Ϊ0����ʵ�ﹺ�ﳵ�������̻�����0�������Ʒ���̻�����1
redoFlag="0"

'��Ǯ�ĺ��������˻���
''��δ�Ϳ�Ǯǩ���������Э�飬����Ҫ��д������
pid=""


   
'���ɼ���ǩ����
''����ذ�������˳��͹�����ɼ��ܴ���
	signMsgVal=appendParam(signMsgVal,"inputCharset",inputCharset)
	signMsgVal=appendParam(signMsgVal,"pageUrl",pageUrl)
	signMsgVal=appendParam(signMsgVal,"bgUrl",bgUrl)
	signMsgVal=appendParam(signMsgVal,"version",version)
	signMsgVal=appendParam(signMsgVal,"language",language)
	signMsgVal=appendParam(signMsgVal,"signType",signType)
	signMsgVal=appendParam(signMsgVal,"merchantAcctId",merchantAcctId)
	signMsgVal=appendParam(signMsgVal,"payerName",payerName)
	signMsgVal=appendParam(signMsgVal,"payerContactType",payerContactType)
	signMsgVal=appendParam(signMsgVal,"payerContact",payerContact)
	signMsgVal=appendParam(signMsgVal,"orderId",orderId)
	signMsgVal=appendParam(signMsgVal,"orderAmount",orderAmount)
	signMsgVal=appendParam(signMsgVal,"orderTime",orderTime)
	signMsgVal=appendParam(signMsgVal,"productName",productName)
	signMsgVal=appendParam(signMsgVal,"productNum",productNum)
	signMsgVal=appendParam(signMsgVal,"productId",productId)
	signMsgVal=appendParam(signMsgVal,"productDesc",productDesc)
	signMsgVal=appendParam(signMsgVal,"ext1",ext1)
	signMsgVal=appendParam(signMsgVal,"ext2",ext2)
	signMsgVal=appendParam(signMsgVal,"payType",payType)
	signMsgVal=appendParam(signMsgVal,"bankId",bankId)
	signMsgVal=appendParam(signMsgVal,"redoFlag",redoFlag)
	signMsgVal=appendParam(signMsgVal,"pid",pid)
	signMsgVal=appendParam(signMsgVal,"key",key)
signMsg= Ucase(md5(signMsgVal))
	
	'���ܺ�����������ֵ��Ϊ�յĲ�������ַ���
	Function appendParam(returnStr,paramId,paramValue)

		If returnStr <> "" Then
			If paramValue <> "" then
				returnStr=returnStr&"&"&paramId&"="&paramValue
			End if
		Else 
			If paramValue <> "" then
				returnStr=paramId&"="&paramValue
			End if
		End if
		
		appendParam=ReturnStr

	End Function
	'���ܺ�����������ֵ��Ϊ�յĲ�������ַ���������

	'���ܺ�������ȡ14λ������
	Function getDateStr() 
	dim dateStr1,dateStr2,strTemp 
	dateStr1=split(cstr(formatdatetime(now(),2)),"-") 
	dateStr2=split(cstr(formatdatetime(now(),3)),":") 

	for each StrTemp in dateStr1 
	if len(StrTemp)<2 then 
	getDateStr=getDateStr & "0" & strTemp 
	else 
	getDateStr=getDateStr & strTemp 
	end if 
	next 

	for each StrTemp in dateStr2 
	if len(StrTemp)<2 then 
	getDateStr=getDateStr & "0" & strTemp 
	else 
	getDateStr=getDateStr & strTemp 
	end if
	next
	End function 
	'���ܺ�������ȡ14λ�����ڡ�����
	
%>

<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>ʹ�ÿ�Ǯ֧��</title>
		<meta http-equiv="content-type" content="text/html; charset=gb2312" >
	</head>
	
<BODY>
	
	<div align="center">
		<table width="259" border="0" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC" >
			<tr bgcolor="#FFFFFF">
				<td width="80">֧����ʽ:</td>
				<td >��Ǯ[99bill]</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td >�������:</td>
				<td ><%=orderId %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>�������:</td>
				<td><%=orderAmount/100 %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>֧����:</td>
				<td><%=payerName %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>��Ʒ����:</td>
				<td><%=productName %></td>
			</tr>
			<tr>
				<td></td>
				<td></td>
			</tr>
	  </table>
	</div>

	<div align="center" style="font-size=12px;font-weight: bold;color=red;">
		<form name="kqPay" method="post" action="https://sandbox.99bill.com/gateway/recvMerchantInfoAction.htm">
			<input type="hidden" name="inputCharset" value="<%=inputCharset %>">
			<input type="hidden" name="bgUrl" value="<%=bgUrl %>">
			<input type="hidden" name="pageUrl" value="<%=pageUrl %>">
			<input type="hidden" name="version" value="<%=version %>">
			<input type="hidden" name="language" value="<%=language %>">
			<input type="hidden" name="signType" value="<%=signType %>">
			<input type="hidden" name="signMsg" value="<%=signMsg %>">
			<input type="hidden" name="merchantAcctId" value="<%=merchantAcctId %>">
			<input type="hidden" name="payerName" value="<%=payerName %>">
			<input type="hidden" name="payerContactType" value="<%=payerContactType %>">
			<input type="hidden" name="payerContact" value="<%=payerContact %>">
			<input type="hidden" name="orderId" value="<%=orderId %>">
			<input type="hidden" name="orderAmount" value="<%=orderAmount %>">
			<input type="hidden" name="orderTime" value="<%=orderTime %>">
			<input type="hidden" name="productName" value="<%=productName %>">
			<input type="hidden" name="productNum" value="<%=productNum %>">
			<input type="hidden" name="productId" value="<%=productId %>">
			<input type="hidden" name="productDesc" value="<%=productDesc %>">
			<input type="hidden" name="ext1" value="<%=ext1 %>">
			<input type="hidden" name="ext2" value="<%=ext2 %>">
			<input type="hidden" name="payType" value="<%=payType %>">
			<input type="hidden" name="bankId" value="<%=bankId %>">
			<input type="hidden" name="redoFlag" value="<%=redoFlag %>">
			<input type="hidden" name="pid" value="<%=pid %>">
			<input type="submit" name="submit" value="�ύ����Ǯ">
			
		</form>		
	</div>
	
</BODY>
</HTML>