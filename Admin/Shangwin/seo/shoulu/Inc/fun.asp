<%
host=lcase(request.servervariables("HTTP_HOST"))

Function getHTTPPage(Path,charset)
        t = GetBody(Path)
        getHTTPPage=BytesToBstr(t,charset)
End function

Function GetBody(url) 
        on error resume next
        'Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
        Set Retrieval = CreateObject("MSXML2.XMLHTTP") 
        With Retrieval 
        .Open "Get", url, False, "", "" 
        .Send 
        if Retrieval.readystate<>4 then 
			GetBody="0"
			exit function
        end if
        GetBody = .ResponseBody
        End With 
        Set Retrieval = Nothing 
End Function

Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText 
        objstream.Close
        set objstream = nothing
End Function

Function Newstring(wstr,strng)
        Newstring=Instr(lcase(wstr),lcase(strng))
        if Newstring<=0 then Newstring=Len(wstr)
End Function

Function Newstring(S_Code,strng)
 Newstring=Instr(lcase(S_Code),lcase(strng))
 if Newstring<=0 then Newstring=Len(S_Code)
End Function

Function RemoveHTML(strHTML) 
Dim objRegExp, Match, Matches 
Set objRegExp = New Regexp 
objRegExp.IgnoreCase = True 
objRegExp.Global = True '取闭合的<> 
objRegExp.Pattern = "<.+?>" '进行匹配 
Set Matches = objRegExp.Execute(strHTML) 
For Each Match in Matches 
strHtml=Replace(strHTML,Match.Value,"") 
Next 
RemoveHTML=strHTML 
Set objRegExp = Nothing 
End Function
%>