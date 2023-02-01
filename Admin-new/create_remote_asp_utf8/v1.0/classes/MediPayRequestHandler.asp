<!--#include file="../util/md5.asp"-->
<!--#include file="../util/tenpay_util.asp"-->
<%
'
'�Ƹ�ͨ�н鵣��֧��������
'============================================================================
'api˵����
'init(),��ʼ��������Ĭ�ϸ�һЩ������ֵ����cmdno,date�ȡ�
'getGateURL()/setGateURL(),��ȡ/������ڵ�ַ,����������ֵ
'getKey()/setKey(),��ȡ/������Կ
'getParameter()/setParameter(),��ȡ/���ò���ֵ
'getAllParameters(),��ȡ���в���
'getRequestURL(),��ȡ������������URL
'doSend(),�ض��򵽲Ƹ�֧ͨ��
'getDebugInfo(),��ȡdebug��Ϣ
'
'============================================================================
'

Class MediPayRequestHandler
	
	'����url��ַ
	Private gateUrl
	
	'��Կ
	Private key
	
	'����Ĳ���
	Private parameters
	
	'debug��Ϣ
	Private debugInfo

	'��ʼ���캯��
	Private Sub class_initialize()
		gateUrl = "http://service.tenpay.com/cgi-bin/v3.0/payservice.cgi"
		key = ""
		Set parameters = Server.CreateObject("Scripting.Dictionary")
		debugInfo = ""
	End Sub

	'��ʼ��������Ĭ�ϸ�һЩ������ֵ����cmdno,date�ȡ�
	Public Function init()
		parameters.RemoveAll

		'�Զ�������ԭ������
		parameters.Add "attach", ""
		
		'ƽ̨���ʺ�
		parameters.Add "chnid", ""
		
		'�������
		parameters.Add "cmdno", "12"
		
		'�������� 1:gbk 2:utf-8
		parameters.Add "encode_type", "1"
		
		'����˵�������ܰ���<>����%�����ַ�
		parameters.Add "mch_desc", ""
		
		'��Ʒ���ƣ����ܰ���<>����%�����ַ�
		parameters.Add "mch_name", ""
		
		'��Ʒ�ܼۣ���λΪ�֡�
		parameters.Add "mch_price", ""
		
		'�ص�֪ͨURL
		parameters.Add "mch_returl", ""
		
		'�������ͣ�1��ʵ�ｻ�ף�2�����⽻��
		parameters.Add "mch_type", ""
		
		'�̼ҵĶ�����
		parameters.Add "mch_vno", ""
		
		'�Ƿ���Ҫ�ڲƸ�ͨ�������Ϣ��1����Ҫ��2������Ҫ��
		parameters.Add "need_buyerinfo", ""
		
		'���ҲƸ�ͨ�ʺ�
		parameters.Add "seller", ""
		
		'֧������̻�֧�����չʾҳ��
		parameters.Add "show_url", ""
		
		'������˾��������ʽ˵��
		parameters.Add "transport_desc", ""
		
		'������֧������������
		parameters.Add "transport_fee", ""
		
		'�汾��
		parameters.Add "version", "2"
		
		'ժҪ
		parameters.Add "sign", ""

	End Function

	'��ȡ��ڵ�ַ,����������ֵ
	Public Function getGateURL()
		getGateURL = gateUrl
	End Function
	
	'������ڵ�ַ,����������ֵ
	Public Function setGateURL(gateUrl_)
		gateUrl = gateUrl_
	End Function

	'��ȡ��Կ
	Public Function getKey()
		getKey = key
	End Function
	
	'������Կ
	Public Function setKey(key_)
		key = key_
	End Function
	
	'��ȡ����ֵ
	Public Function getParameter(parameter)
		getParameter = parameters.Item(parameter)
	End Function
	
	'���ò���ֵ
	Public Function setParameter(parameter, parameterValue)
		If parameters.Exists(parameter) = True Then
			parameters.Remove(parameter)
		End If
		parameters.Add parameter, parameterValue	
	End Function

	'��ȡ��������Ĳ���,����Scripting.Dictionary
	Public Function getAllParameters()
			
		'��������
		SortDictionary parameters, dictKey

		getAllParameters = parameters
	End Function

	'��ȡ������������URL
	Public Function getRequestURL()

		Call createSign()
		
		Dim reqPars
		Dim k
		For Each k In parameters
			reqPars = reqPars & k & "=" & Server.URLEncode(parameters(k)) & "&" 
		Next
		
		'ȥ�����һ��&
		reqPars = Left(reqPars, Len(reqPars)-1)

		getRequestURL = getGateURL & "?" & reqPars

	End Function
	
	'��ȡdebug��Ϣ
	Public Function getDebugInfo()
		getDebugInfo = debugInfo
	End Function
	
	'�ض��򵽲Ƹ�֧ͨ��
	Public Function doSend()
		Response.Redirect(getRequestURL())
		Response.End
	End Function

	'����ǩ��
	Private Sub createSign()
		'��������
		SortDictionary parameters, dictKey

		Dim signPars
		Dim k
		For Each k In parameters
			Dim v
			v = parameters(k)
			'���ַ������μ�ǩ��
			If v <> "" And k <> "sign" Then
				signPars = signPars & k & "=" & v & "&"
			End If
		Next
		
		'��Կ�������
		signPars = signPars & "key=" & key

		Dim sign
		sign= LCase(ASP_MD5(signPars))

		setParameter "sign", sign

		'debuginfo
		debugInfo = signPars & " => sign:" & sign

	End Sub
	
End Class


%>