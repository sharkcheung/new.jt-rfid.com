<!--#include file = "../admin_conn.asp" -->
<!-- #include file="../config1.asp" -->
<%
	'�汾��2.0
	'���ڣ�2009-07-30
	'
	'˵����
	'���´���ֻ�Ƿ����̻����ԣ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
	'�ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���

	
	seller_email	= alipay_uid	 '����дǩԼ֧�����˺ţ�
	partner			= alipay_id	 '��дǩԼ֧�����˺Ŷ�Ӧ��partnerID��
	key			    = alipay_key	 '��дǩԼ�˺Ŷ�Ӧ�İ�ȫУ����

    notify_url		= ""&yuming&"/alipay/Alipay_Notify.asp"	        '���׹����з�����֪ͨ��ҳ�� Ҫ�� http://��ʽ������·��������http://www.alipay.com/alipay/Alipay_Notify.asp  ע���ļ�λ������д��ȷ��
	return_url		= ""&yuming&"/alipay/return_Alipay_Notify.asp"	'��������ת��ҳ�� Ҫ�� http://��ʽ������·��, ����http://www.alipay.com/alipay/return_Alipay_Notify.asp  ע���ļ�λ������д��ȷ��
	'���ʹ����Alipay_Notify.asp����return_Alipay_Notify.asp�������������ļ��������Ӧ�ĺ��������ID�Ͱ�ȫУ����
	logistics_fee	   = "0.00"			'�������ͷ���
	logistics_payment  = "SELLER_PAY"	'�������ͷ��ø��ʽ��SELLER_PAY(����֧��)��BUYER_PAY(���֧��)��BUYER_PAY_AFTER_RECEIVE(��������)
	logistics_type	   = "EXPRESS"		'�������ͷ�ʽ��POST(ƽ��)��EMS(EMS)��EXPRESS(�������)

	 	 
'��½ www.alipay.com ��, ���̼ҷ���,���Կ���֧������ȫУ����ͺ���id,������������ 
%>