<!--#include file = "../admin/admin_conn.asp" -->
<!-- #include file="../config1.asp" -->
<%
	'�汾��2.0
	'���ڣ�2009-07-30
	'
	'˵����
	'���´���ֻ�Ƿ����̻����ԣ��ṩ���������룬�̻����Ը����Լ���վ����Ҫ�����ռ����ĵ���д,����һ��Ҫʹ�øô��롣
	'�ô������ѧϰ���о�֧�����ӿ�ʹ�ã�ֻ���ṩһ���ο���
%>

<!--#include file="alipayto/alipay_payto.asp"-->
<%
    '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
'����Ĳ���	
    service         =   "create_partner_trade_by_buyer"   'trade_create_by_buyer ��ʾ��׼˫�ӿڣ� create_partner_trade_by_buyer ��ʾ�������׽ӿ�
	subject			=	pay_config(session("jh_pro_id"),1)	'��Ʒ���ƣ�����ͻ��߹��ﳵ���̿�����Ϊ  "�����ţ�"&request("�ͻ���վ����")
	body			=	pay_config(session("jh_pro_id"),1)		'��Ʒ����
	out_trade_no    =   session("jh_pro_id")  '��ʱ���ȡ�Ķ����ţ������޸ĳ��Լ���վ�Ķ����ţ���֤ÿ���ύ�Ķ���Ψһ����
	price		    =	"0.01"'session("jh_pro_price")			'��Ʒ����			0.01��100000000.00  ��ע����Ҫ����3,000.00���۸�֧��","��
    quantity        =   pay_config(session("jh_pro_id"),2)             '��Ʒ����,����߹��ﳵĬ��Ϊ1
    seller_email    =   alipay_uid   '���ҵ�֧�����ʺţ�c2c�ͻ������Ը��Ĵ˲�����

 '�����ǿ�ѡ�Ĳ��� ���û�п���Ϊ�ա�ע�⣺��������ϵ��ַ���������� ������Ҫô��Ϊ�գ�Ҫô������Ϊ�ա�
  	show_url        = ""  '�̻���չʾ��ַ�����Ӻ��治���Զ������
	receive_name    = ""  '�ջ�������
    receive_address = ""  '�ջ��˵�ַ
	receive_zip     = ""  '�ʱ�5 λ��6 λ�������
	receive_phone   = ""  '�ջ��˵绰
	receive_mobile  = ""  '�ջ����ֻ� ������11 λ����
	buyer_email     = ""  '��ҵ�֧�����˺�
    discount        = ""  '��Ʒ�ۿ�

 '�����Ҫ����Ӽ���������ʽ���������ӵڶ�����������,�������Ҫ������Ϊ��
   	logistics_fee_1	   = ""			'�������ͷ���  0.00
	logistics_payment_1  = ""	'�������ͷ��ø��ʽ��SELLER_PAY(����֧��)��BUYER_PAY(���֧��)��BUYER_PAY_AFTER_RECEIVE(��������)
	logistics_type_1	   = ""		'�������ͷ�ʽ��POST(ƽ��)��EMS(EMS)��EXPRESS(�������)


	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(service,subject,body,out_trade_no,price,quantity,seller_email,show_url,receive_name,receive_address,receive_zip,receive_phone,receive_mobile,buyer_email,discount,logistics_fee_1,logistics_payment_1,logistics_type_1)
	response.Redirect(itemUrl)
%>