<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="./classes/MediPayResponseHandler.asp"-->
<%
'---------------------------------------------------------
'�Ƹ�ͨ�н鵣��֧��Ӧ�𣨴���ص���ʾ�����̻����մ��ĵ����п�������
'---------------------------------------------------------

'ƽ̨����Կ
Dim key
key = "123456"

'����֧��Ӧ�����
Dim resHandler
Set resHandler = new MediPayResponseHandler
resHandler.setKey(key)

'�ж�ǩ��
If resHandler.isTenpaySign() = True Then
	
	Dim cft_tid
	Dim total_fee
	Dim retcode
	Dim status

	'�Ƹ�ͨ���׵���
	cft_tid = resHandler.getParameter("cft_tid")
	
	'�����,�Է�Ϊ��λ
	total_fee = resHandler.getParameter("total_fee")

	'������
	retcode = resHandler.getParameter("retcode")

	'״̬
	status = resHandler.getParameter("status")
	
	'------------------------------
	'����ҵ��ʼ
	'------------------------------ 

	'ע�⽻�׵���Ҫ�ظ�����
	'ע���жϷ��ؽ��
	
	'�������ж�
	If "0" = retcode Then
		Select Case status
			Case "1":	'���״���

			Case "2":	'�ջ��ַ��д���

			Case "3":	'��Ҹ���ɹ���ע���ж϶����Ƿ��ظ����߼�

			Case "4":	'���ҷ����ɹ�

			Case "5":	'����ջ�ȷ�ϣ����׳ɹ�

			Case "6":	'���׹رգ�δ��ɳ�ʱ�ر�

			Case "7":	'�޸Ľ��׼۸�ɹ�

			Case "8":	'��ҷ����˿�

			Case "9":	'�˿�ɹ�

			Case "10":	'�˿�ر�

			Case else:	'error
				'nothing to do
		End Select
	Else
		'����ķ�����
		
	End If
	
	'------------------------------
	'����ҵ�����
	'------------------------------

	resHandler.doShow()
	

Else

	'ǩ��ʧ��
	Response.Write("ǩ��ǩ֤ʧ��")

	'Dim debugInfo
	'debugInfo = resHandler.getDebugInfo()
	'Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")

End If


%>