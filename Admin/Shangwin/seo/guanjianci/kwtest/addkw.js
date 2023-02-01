function adk(o){
    var s=encodeURIComponent(o.innerHTML);
	
    var xml = new ActiveXObject("MSXML2.XMLHTTP");
    xml.open("get","addkw.asp?s="+s+"",false);
	xml.send();
	
	//document.getElementById('tishi').innerHTML=xml.responseText;
	//ymPrompt.alert(xml.responseText+'消息内容');

     	  ymPrompt.alert({message:xml.responseText,title:'结果提示       (本提示框会自动关闭)',width:400});
	  setTimeout(function(){ymPrompt.doHandler('ok')},1000);

}