function adk(o){
    var s=encodeURIComponent(o.innerHTML);
	
    var xml = new ActiveXObject("MSXML2.XMLHTTP");
    xml.open("get","addkw.asp?s="+s+"",false);
	xml.send();
	
	//document.getElementById('tishi').innerHTML=xml.responseText;
	//ymPrompt.alert(xml.responseText+'��Ϣ����');

     	  ymPrompt.alert({message:xml.responseText,title:'�����ʾ       (����ʾ����Զ��ر�)',width:400});
	  setTimeout(function(){ymPrompt.doHandler('ok')},1000);

}