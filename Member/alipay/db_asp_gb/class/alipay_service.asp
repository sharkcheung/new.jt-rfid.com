<%
	'������alipay_service
	'���ܣ�֧�����ⲿ����ӿڿ���
	'��ϸ����ҳ��������������Ĵ����ļ�������Ҫ�޸�
	'�汾��3.0
	'�޸����ڣ�2010-07-26
	'˵����
	'���´���ֻ��Ϊ�˷����̻����Զ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
	'�ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο�
%>

<!--#include file="alipay_function.asp"-->

<%

dim gateway			'���ص�ַ
dim mysign			'���ܽ����ǩ�������
dim sPara		'��Ҫ���ܵ��Ѿ����˺�Ĳ�������

'********************************************************************************

'���캯��
'�������ļ�������ļ��г�ʼ������
'inputPara ��Ҫ���ܵĲ�������
function alipay_service(inputPara)
	gateway = "https://www.alipay.com/cooperate/gateway.do?"
	sPara = para_filter(inputPara)
	sort_para = arg_sort(sPara)		'�õ�����ĸa��z�����ļ��ܲ�������
	'���ǩ�����
	mysign = build_mysign(sort_para,key,sign_type,input_charset)
end function

'********************************************************************************

'��������URL��GET��ʽ����
'��� ����url
function create_url()
	url = gateway
	sort_para = arg_sort(sPara)
	arg = create_linkstring_urlencode(sort_para)	'����������Ԫ�أ����ա�����=����ֵ����ģʽ�á�&���ַ�ƴ�ӳ��ַ���
	url = url & arg & "sign=" &mysign & "&sign_type=" & sign_type
	create_url = url
end function

'********************************************************************************

'����Post���ύHTML��POST��ʽ����
'��� ���ύHTML�ı�
function build_postform()
	sHtml = "<form id='alipaysubmit' name='alipaysubmit' action='"& gateway &"_input_charset="&input_charset&"' method='post'>"

	nCount = ubound(sPara)
	for i = 0 to nCount
		'��sArray���������Ԫ�ظ�ʽ��������=ֵ���ָ��
		pos = Instr(sPara(i),"=")			'���=�ַ���λ��
		nLen = Len(sPara(i))				'����ַ�������
		itemName = left(sPara(i),pos-1)		'��ñ�����
		itemValue = right(sPara(i),nLen-pos)'��ñ�����ֵ
		
		sHtml = sHtml & "<input type='hidden' name='"& itemName &"' value='"& itemValue &"'/>"
	next

	sHtml = sHtml & "<input type='hidden' name='sign' value='"& mysign &"'/>"
	sHtml = sHtml & "<input type='hidden' name='sign_type' value='"& sign_type &"'/></form>"

	sHtml = sHtml & "<input type=""button"" name=""v_action"" value=""֧����ȷ�ϸ���"" onClick=""document.forms['alipaysubmit'].submit();"">"
	build_postform = sHtml
end function

%>