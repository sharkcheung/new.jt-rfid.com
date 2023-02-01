<%
'********************************************** 
' vbs Cache类 
' 
' 属性valid，是否可用，取值前判断 
' 属性name，cache名，新建对象后赋值 
' 方法add(值,到期时间)，设置cache内容 
' 属性value，返回cache内容 
' 属性blempty，是否未设置值 
' 方法makeEmpty，释放内存，测试用 
' 方法equal(变量1)，判断cache值是否和变量1相同 
' 方法expires(time)，修改过期时间为time 
' 木鸟 2002.12.24 
' http://www.aspsky.net/ 
'********************************************** 
class Cache
private obj 'cache内容 
private expireTime '过期时间 
private expireTimeName '过期时间application名 
private cacheName 'cache内容application名 
private path 'uri 

private sub class_initialize() 
	path=request.servervariables("url") 
	path=left(path,instrRev(path,"/")) 
end sub 

private sub class_terminate() 
end sub 

public property get blEmpty 
	'是否为空 
	if isempty(obj) then 
	blEmpty=true 
	else 
	blEmpty=false 
	end if 
end property 

public property get valid 
	'是否可用(过期) 
	if isempty(obj) or not isDate(expireTime) then 
	valid=false 
	elseif CDate(expireTime)<now then 
	valid=false 
	else 
	valid=true 
	end if 
end property 

public property let name(str) 
	'设置cache名 
	cacheName=str & path 
	obj=application(cacheName) 
	expireTimeName=str & "expires" & path 
	expireTime=application(expireTimeName) 
end property 

public property let expires(tm) 
	'重设置过期时间 
	expireTime=tm 
	application.lock 
	application(expireTimeName)=expireTime 
	application.unlock 
end property 

public sub add(var,expire) 
	'赋值 
	if isempty(var) or not isDate(expire) then 
	exit sub 
	end if 
	obj=var 
	expireTime=expire 
	application.lock 
	application(cacheName)=obj 
	application(expireTimeName)=expireTime 
	application.unlock 
end sub 

public property get value 
	'取值 
	if isempty(obj) or not isDate(expireTime) then 
	value=null 
	elseif CDate(expireTime)<now then 
	value=null 
	else 
	value=obj 
	end if
end property 
	
' 删除某以缓存 
Public Sub Clear(Key) 
	Application.Contents.Remove(Key) 
End Sub 	

public sub makeEmpty()
	'释放application
	application.lock
	application(cacheName)=empty
	application(expireTimeName)=empty
	application.unlock
	obj=empty
	expireTime=empty
end sub

	public function equal(var2) 
		'比较 
		if typename(obj)<>typename(var2) then 
		equal=false 
		elseif typename(obj)="Object" then 
		if obj is var2 then 
		equal=true 
		else 
		equal=false 
		end if 
		elseif typename(obj)="Variant()" then 
		if join(obj,"^")=join(var2,"^") then 
		equal=true 
		else 
		equal=false 
		end if 
		else 
		if obj=var2 then 
		equal=true 
		else 
		equal=false 
		end if 
		end if 
	end function 
end class
%>