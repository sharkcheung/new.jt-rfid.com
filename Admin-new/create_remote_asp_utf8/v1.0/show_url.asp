<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="./classes/MediPayResponseHandler.asp"-->
<%
'---------------------------------------------------------
'�Ƹ�ͨ�н鵣��֧���ɹ���ʾҳ��ʾ�����̻����մ��ĵ����п�������
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
		
		'��Ҹ���ɹ�
		If "3" = status Then
			
			Response.Write("֧���ɹ�")
		End If
	Else
		Response.Write("֧��ʧ��")
	End If
	
	'------------------------------
	'����ҵ�����
	'------------------------------
	
Else

	'ǩ��ʧ��
	Response.Write("ǩ��ǩ֤ʧ��")

	'Dim debugInfo
	'debugInfo = resHandler.getDebugInfo()
	'Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")

End If


%>