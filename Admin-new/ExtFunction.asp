<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Field.asp
'文件用途：自定义字段管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If


'---- DataTypeEnum Values ----
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

'---- FieldAttributeEnum Values ----
Const adFldMayDefer = &H00000002
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldFixed = &H00000010
Const adFldIsNullable = &H00000020
Const adFldMayBeNull = &H00000040
Const adFldLong = &H00000080
Const adFldRowID = &H00000100
Const adFldRowVersion = &H00000200
Const adFldCacheDeferred = &H00001000

'---- SchemaEnum Values ----
'---- SchemaEnum Values ----
Const adSchemaProviderSpecific = -1
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCollations = 3
Const adSchemaColumns = 4
Const adSchemaCheckConstraints = 5
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaKeyColumnUsage = 8
Const adSchemaReferentialConstraints = 9
Const adSchemaTableConstraints = 10
Const adSchemaColumnsDomainUsage = 11
Const adSchemaIndexes = 12
Const adSchemaColumnPrivileges = 13
Const adSchemaTablePrivileges = 14
Const adSchemaUsagePrivileges = 15
Const adSchemaProcedures = 16
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaProviderTypes = 22
Const adSchemaViews = 23
Const adSchemaViewColumnUsage = 24
Const adSchemaViewTableUsage = 25
Const adSchemaProcedureParameters = 26
Const adSchemaForeignKeys = 27
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29
Const adSchemaDBInfoKeywords = 30
Const adSchemaDBInfoLiterals = 31
Const adSchemaCubes = 32
Const adSchemaDimensions = 33
Const adSchemaHierarchies = 34
Const adSchemaLevels = 35
Const adSchemaMeasures = 36
Const adSchemaProperties = 37
Const adSchemaMembers = 38
Const adSchemaTrustees = 39
Const adSchemaFunctions = 40
Const adSchemaActions = 41
Const adSchemaCommands = 42
Const adSchemaSets = 43

'定义页面变量
Dim Fk_Field_Name,Fk_Field_Tag,Fk_Field_Type,Fk_Field_Type1,Fk_Field_Type2

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call ExtFuncList() '自定义字段列表
	Case 2
		Call ExtFuncAddForm() '添加自定义字段表单
	Case 3
		Call ExtFuncAddDo() '执行添加自定义字段
	Case 4
		Call ExtFuncEditForm() '修改自定义字段表单
	Case 5
		Call ExtFuncEditDo() '执行修改自定义字段
	Case 6
		Call ExtFuncDelDo() '执行删除自定义字段
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：FieldList()
'作    用：自定义字段列表
'参    数：
'==========================================
Sub ExtFuncList()
%>


<div id="ListContent">
	<div class="gnsztopbtn">
    	<h3>表字段管理</h3><select name="fieldselect" id="fieldselect"><option value="Fk_Product">Fk_Product</option></select>
        <a style="width:90px; padding-left:30px;" class="sixjis" href="javascript:void(0);" onclick="ShowBox('ExtFunction.asp?Type=2','添加新字段','450px');">添加表字段</a>
        <a class="shuax" href="javascript:void(0);" onclick="SetRContent('MainRight','ExtFunction.asp?Type=1');return false">刷新</a>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
            <th align="center" class="ListTdTop" width="240px">字段名称</th>
            <th align="center" class="ListTdTop">字段类型</th>
            <th align="center" class="ListTdTop">字段大小</th>
            <th align="center" class="ListTdTop">字段说明</th>
            <th align="center" class="ListTdTop">是否允许空</th>
            <th align="center" class="ListTdTop">自动编号</th>
            <th align="center" class="ListTdTop">主键</th>
            <th align="center" class="ListTdTop" width="150">操作</th>
        </tr>
<%
	dim primary,primarykey
	Set primary = Conn.OpenSchema(adSchemaPrimaryKeys,Array(empty,empty,"Fk_Product"))
	if primary("COLUMN_NAME") <> "" then
		primarykey = primary("COLUMN_NAME")
	end if
	
	primary.Close
	set primary = nothing




        Dim OutField
        Dim MyDB,MyTable
        Dim Key,Key1
        Set OutField = Server.CreateObject( "Scripting.Dictionary" )
        OutField.CompareMode = 1   

        Set MyDB    = Server.CreateObject("ADOX.Catalog")
        Set MyTable = Server.CreateObject("ADOX.Table")

        MyDB.ActiveConnection = Conn         '数据库连接,自己写哈~
        Set MyTable = MyDB.Tables("Fk_Product")
       
        For Each Key In MyTable.Columns			
%>
        <tr>
            <td height="20" style="padding-left:20px;"><%=key.name%></td>
            <td align="center"><%=typ(key.type)%></td>
            <td align="center"><%=key.definedsize%></td>
            <td align="center"><%=Key.Properties("Description")%></td>
		  	<td align="center"><%=IIf(LCase(Key.Properties("Nullable") & "") = "true","是","否")%></td>
		  	<td align="center"><%=iif(Key.Properties("Autoincrement") = True,"是","否")%></td>
		  	<td align="center"><%=iif(Key.name = primarykey,"是","否")%></td>
            <td align="center"><a style="width:auto; line-height:21px; margin-right:10px; background:none;" href="javascript:void(0);" onclick="ShowBox('ExtFunction.asp?Type=4&tbfield=<%=Key.name%>&tb='+fieldselect.value,'修改表字段','450px');">修改</a> <a  style="width:auto; line-height:21px; margin-right:0; background:none" href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Key.name%>”，此操作不可逆！','ExtFunction.asp?Type=6&tbfield=<%=Key.name%>&tb='+fieldselect.value,'MainRight','ExtFunction.asp?Type=1');">删除</a></td>
        </tr>
<%
		next 
        Set OutField = Nothing
        Set MyTable = nothing
        Set MyDB = Nothing
'	Else
%>
        <tr>
            <td height="30" colspan="8">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：FieldFielddForm()
'作    用：添加自定义字段表单
'参    数：
'==========================================
Sub ExtFuncAddForm()

%>
<form id="FieldFieldd" name="FieldFieldd" method="post" action="ExtFunction.asp?Type=3" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
	
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td width="100" height="25" align="right">表名称：</td>
	        <td><select style="width:130px;" id="tb" name="tb"><option value="Fk_Product">Fk_Product</select></td>
	        </tr>
        <tr>
	    <tr>
	        <td height="25" align="right">字段名称：</td>
	        <td><input class="Input" name="fldname" type="text" size="30" maxlength="50"></td>
	        </tr>
        <tr>
            <td height="30" align="right">字段类型：</td>
            <td><%=fieldtypelist(0)%></td>
        </tr>
        <tr>
            <td height="30" align="right">字段大小：</td>
            <td><input class="Input" name="fldsize" type="text" size="30" maxlength="50"></td>
        </tr>
        <tr>
            <td height="30" align="right">字段说明：</td>
            <td><input class="Input" name="desc" type="text" value="" size="30" maxlength="50"></td>
        </tr>
        <!--tr>
            <td height="30" align="right">是否允许为空：</td>
            <td><input name="null" type="checkbox" value="ON" checked></td>
        </tr-->
        <tr>
            <td height="30" align="right">自动编号：</td>
            <td><input style="margin-left:10px;" type="checkbox" name="autoincrement" value="ON"></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" class="tcbtm" style="width:93%; margin: 0 auto; text-align:left;">
        <input style="margin-left:113px;" type="submit" onclick="Sends('FieldFieldd','ExtFunction.asp?Type=3',0,'',0,1,'MainRight','ExtFunction.asp?Type=1');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：FieldFielddDo
'作    用：执行添加自定义字段
'参    数：
'==============================
Sub ExtFuncAddDo()
	dim fldname,fldtype,fldsize,fldnull,fldautoincrement,table_name,sql,desc
	on error resume next
	fldname = request("fldname")
	fldtype = lcase(request("field_type"))
	fldsize = request("fldsize")
	'fldnull = request("null")
	fldautoincrement = request("autoincrement")
	table_name = request("tb")
	desc=Trim(Request.Form("desc"))
'	if fldname <> "" and fldtype <> "" then
'	  sql = "alter table [" & table_name & "] add ["&fldname&"] " & fldtype
	  
'	  if fldsize <> "" then
'		sql = sql & "(" & fldsize & ")"
'	  end if 
'	  
'	  
'	  if fldautoincrement = "ON" then
'		sql = sql & " identity"
'	  end if
'	  conn.execute(sql)
	dim oCat,tblnam,oColumn
	'打开表
	Set oCat = Server.CreateObject("ADOX.Catalog")
    Set oCat.ActiveConnection = conn
     '创建列
    Set oColumn = Server.CreateObject("ADOX.Column")
    With oColumn
         Set .ParentCatalog = oCat	'Must set before setting properties
         .Name = fldname
		 
	Select Case fldtype
		case "varchar"
         	.Type = 202
		case "text"
         	.Type = 203
		case "bit"
         	.Type = 11
		case "integer"
         	.Type = 3
		case "smallint"
         	.Type = 2
		case "single"
         	.Type = 4
		case "double"
         	.Type = 5
		case "dateTime"
         	.Type = 7
		case else
         	.Type = 200
	end select
	
		 if fldsize <> "" then
         .DefinedSize = fldsize
		 end if
		 ' if fldnull = "ON" then
         .Properties("Nullable") = True
		 ' end if
		 if fldautoincrement = "ON" then
         .Properties("AutoIncrement") = true
		 end if
		 if desc<>"" then
			.Properties("Description") = desc
		 end if
         '.Properties("Jet OLEDB:Allow Zero Length") = True
    End With
    oCat.Tables(table_name).Columns.Append oColumn
    ' 完成

    Set oColumn = Nothing
    Set oCat = Nothing
	
'	set clx=server.CreateObject("ADOX.Column")
'	Set cat.ActiveConnection = conn
'	tblnam.Name = table_name
'	clx.ParentCatalog = cat
'	clx.Type = 3
'	clx.Name = fldname
'	if fldnull <> "ON" then
'		'clx.Properties("Nullable") = true
'	end if
'	if fldsize <> "" then
'		'clx.MaxLength="100"
'	end if
'	if fldautoincrement = "ON" then
'		'clx.Properties("AutoIncrement") = true
'	end if
'	if desc<>"" then
'		'clx.Properties("Description") = desc
'	end if
'	tblnam.Columns.Append clx
'	'cat.Tables.Append tblnam
'	Set clx		= Nothing
'	Set cat		= Nothing
'	Set tblnam 	= Nothing

	'else
	 ' response.write "输入数据错误！"
	 ' response.end
	'end if
	if err <> 0 then
		response.write err.description
	else
	  	response.write "字段【"&fldname&"】添加成功！"
	end if
End Sub

'==========================================
'函 数 名：FieldEditForm()
'作    用：修改自定义字段表单
'参    数：
'==========================================
Sub ExtFuncEditForm()
	dim tbfield,tb
	tb=trim(Request.QueryString("tb"))
	tbfield=trim(Request.QueryString("tbfield"))
	
	
	Sqlstr="Select * From ["&tb&"]"
	Rs.Open Sqlstr,Conn,1,1
		for i = 0 to rs.fields.count - 1
			if rs(i).name = tbfield then
%>
<script LANGUAGE="JavaScript">
	function validate(theForm) {
		if (theForm.type.value == "")
		{
			alert("请选择数据类型");
			theForm.type.focus();
			return (false);
		}
			return (true);
	}
</script>
<form id="FieldEdit" name="FieldEdit" method="post" action="ExtFunction.asp?Type=5" onsubmit="return false;">
<!--<div id="BoxTop" style="width:500px;">修改表(<%=tb%>)-字段(<%=tbfield&rs(i).type%>)[按ESC关闭窗口]</div>-->
<div id="BoxContents" style="width:93%; padding-top:20px;">
	<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
	    <tr>
	        <td height="25" align="right" width="100">字段名称：</td>
	        <td><%=rs(i).name%></td>
	        </tr>
        <tr>
            <td height="30" align="right">字段类型：</td>
            <td><%=fieldtypelist(typ(rs(i).type))%></td>
        </tr>
        <tr>
            <td height="30" align="right">字段大小：</td>
            <td><input class="Input" name="size" type="text" value="<%=iif(rs(i).type=203,"",rs(i).definedsize)%>" size="30" maxlength="50"></td>
        </tr>
        <tr>
            <td height="30" align="right">字段说明：</td>
            <td><input class="Input" name="desc" type="text" value="<%=getDesc(tb,tbfield)%>" size="30" maxlength="50"></td>
        </tr>
        <!--tr>
            <td height="30" align="right">是否允许为空：</td>
            <td><input name="nulls" type="checkbox" id="nulls" value="null"<%=iif((rs(i).Attributes and adFldIsNullable)=false,""," checked")%> /></td>
        </tr-->
        <tr>
            <td height="30" align="right">自动编号：</td>
            <td style="padding-left:10px;"><input type="checkbox" name="autoincrement" value="y"<%=iif(rs(i).Properties("ISAUTOINCREMENT") = True," checked","")%> /></td>
        </tr>
	    </table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto; text-align:left" class="tcbtm">
		<input type="hidden" name="tb" value="<%=tb%>" />
		<input type="hidden" name="tbfield" value="<%=tbfield%>" />
        <input style="margin-left:113px;" type="submit" onclick="Sends('FieldEdit','ExtFunction.asp?Type=5',0,'',0,1,'MainRight','ExtFunction.asp?Type=1');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
		exit for
		end if
	next
	rs.close
End Sub

'==============================
'函 数 名：FieldEditDo
'作    用：执行修改自定义字段
'参    数：
'==============================
Sub ExtFuncEditDo()
	dim tb,tbfield,field_type,sizes,nulls,autoincrement,sql,desc
	tb=Trim(Request.Form("tb"))
	tbfield=Trim(Request.Form("tbfield"))
	field_type=Trim(Request.Form("field_type"))
	sizes=Trim(Request.Form("size"))
	' nulls=Trim(Request.Form("nulls")&" ")
	' response.write nulls
	' response.end
	autoincrement=Trim(Request.Form("autoincrement"))
	desc=Trim(Request.Form("desc"))
'	on error resume next
'	Dim cat,oTable,item
'	Set cat = server.CreateObject( "ADOX.Table") 
'	cat.ActiveConnection = conn
'	set oTable = cat.Tables.Item(tb)
'			set item = oTable.Fields.Item(Request.Form("tbfield").Item)
'			item.Name = Request.Form("tbfield").Item
'			item.FieldType = Request.Form("field_type").Item
'			item.MaxLength = Request.Form("size").Item
'			'item.DefaultValue = Request.Form("default").Item
'			item.Description = Request.Form("desc").Item
'			'item.AllowZeroLength = Request.Form("zero_length").Item
'			item.IsNullable =Request.Form("nulls").Item
'			item.UpdateBatch
'	Set cat = Nothing 
'	
'	if err then response.write err.description
	'on error resume next
	sql = "ALTER TABLE [" & tb & "] "
	sql = sql&"ALTER COLUMN [" & tbfield & "] "
	if field_type <> "" then
		sql = sql & field_type
	end if
	if sizes <> "" and field_type<>"Text" then
		sql = sql & "(" & sizes & ")"
	end if
	' if nulls = "" then
		' sql = sql & " NOT NULL"
	' end if
	if autoincrement = "y" then
		sql = sql & " identity"
	end if
	sql = trim(sql)
	' response.write sql
	' response.end
	conn.execute(sql)
	
	dim xCat,fieldd
	set xCat = Server.CreateObject("ADOX.Catalog")
			if not IsEmpty(xCat) and not xCat Is Nothing Then
				set xCat.ActiveConnection = conn
				set fieldd = xCat.Tables(tb).Columns(tbfield)
				with fieldd
					.Properties("Description").Value = desc
				end with
				set fieldd = Nothing
				set xCat = Nothing
			End If
	
	Response.Write("字段修改成功！")
End Sub

'==============================
'函 数 名：FieldDelDo
'作    用：执行删除自定义字段
'参    数：
'==============================
Sub ExtFuncDelDo()
	dim tb,tbfield
	on error resume next
	tb=Trim(Request.QueryString("tb"))
	tbfield=Trim(Request.QueryString("tbfield"))
	conn.execute("alter table ["&tb&"] drop ["&tbfield&"]")
	if err then
		response.write "删除字段失败，请重试！"
	else
		response.write "删除字段成功！"
	end if
End Sub



'==================================================================返回字段类型函数
Function typ(field_type)
	'field_type = 字段类型值
	Select Case field_type
		case adEmpty:typ = "Empty"
		case adTinyInt:typ = "TinyInt"
		case adSmallInt:typ = "SmallInt"
		case adInteger:typ = "Integer"
		case adBigInt:typ = "BigInt"
		case adUnsignedTinyInt:typ = "TinyInt" 'UnsignedTinyInt
		case adUnsignedSmallInt:typ = "UnsignedSmallInt"
		case adUnsignedInt:typ = "UnsignedInt"
		case adUnsignedBigInt:typ = "UnsignedBigInt"
		case adSingle:typ = "Single" 'Single
		case adDouble:typ = "Double" 'Double
		case adCurrency:typ = "Money" 'Currency
		case adDecimal:typ = "Decimal"
		case adNumeric:typ = "Numeric" 'Numeric
		case adBoolean:typ = "Bit" 'Boolean
		case adError:typ = "Error"
		case adUserDefined:typ = "UserDefined"
		case adVariant:typ = "Variant"
		case adIDispatch:typ = "IDispatch"
		case adIUnknown:typ = "IUnknown"
		case adGUID:typ = "GUID" 'GUID
		case adDATE:typ = "DateTime" 'Date
		case adDBDate:typ = "DBDate"
		case adDBTime:typ = "DBTime"
		case adDBTimeStamp:typ = "DateTime" 'DBTimeStamp
		case adBSTR:typ = "BSTR"
		case adChar:typ = "Char"
		case adVarChar:typ = "VarChar"
		case adLongVarChar:typ = "LongVarChar"
		case adWChar:typ = "Text" 'WChar类型 SQL中为Text
		case adVarWChar:typ = "VarChar" 'VarWChar
		case adLongVarWChar:typ = "Text" 'LongVarWChar
		case adBinary:typ = "Binary"
		case adVarBinary:typ = "VarBinary"
		case adLongVarBinary:typ = "LongBinary"'LongVarBinary
		case adChapter:typ = "Chapter"
		case adPropVariant:typ = "PropVariant"
		case else:typ = "Unknown"
	end select
End Function

'==================================================================返回字段类型函数
Function intTotyp(intT)
	'field_type = 字段类型值
	Select Case intT
		case "VarChar":intTotyp = 200
		case "Text":intTotyp = 203
		case "Bit":intTotyp = 11
		case "Integer":intTotyp = 3
		case "SmallInt":intTotyp = 2
		case "Single":intTotyp = 4
		case "Double":intTotyp = 5
		case "DateTime":intTotyp = 7
		case else:intTotyp = 200
	end select
End Function

'==================================================================返回字段类型函数
Function charToEnum(field_type)
	'field_type = 字段类型值
	Select Case field_type
		case "VarChar":charToEnum = "adVarChar"
		case "Text":charToEnum = "adLongVarWChar"
		case "Bit":charToEnum = "adBoolean"
		case "Integer":charToEnum = "adInteger"
		case "SmallInt":charToEnum = "adSmallInt"
		case "Single":charToEnum = "adSingle"
		case "Double":charToEnum = "adDouble"
		case "DateTime":charToEnum = "adDATE"
		case else:charToEnum = "adVarChar"
	end select
End Function

'==================================================================返回字段类型列表
Function fieldtypelist(n)
	dim strlist,str1,str2
	strlist = "<select name=""field_type"">"
		strlist = strlist & "<option value=""VarChar"">文本</option>"
		strlist = strlist & "<option value=""Text"">备注</option>"
		strlist = strlist & "<option value=""Bit"">(是/否)</option>"
		strlist = strlist & "<option value=""TinyInt"">数字(字节)</option>"
		strlist = strlist & "<option value=""SmallInt"">数字(整型)</option>"
		strlist = strlist & "<option value=""Integer"">数字(长整型)</option>"
		strlist = strlist & "<option value=""Single"">数字(单精度)</option>"
		strlist = strlist & "<option value=""Double"">数字(双精度)</option>"
		strlist = strlist & "<option value=""Numeric"">数字(小数)</option>"
		strlist = strlist & "<option value=""GUID"">数字(同步ID)</option>"
		strlist = strlist & "<option value=""DateTime"">时间/日期</option>"
		strlist = strlist & "<option value=""Money"">货币</option>"
		strlist = strlist & "<option value=""Binary"">二进制</option>"
		strlist = strlist & "<option value=""LongBinary"">长二进制</option>"
		strlist = strlist & "<option value=""LongBinary"">OLE 对象</option>"
		
	str1 = """" & n & """"
	str2 = """" & n & """" & " selected"
	strlist = replace(strlist,str1,str2)
	strlist = strlist & "</select>"
	response.write  strlist
End Function
Function IIf(var, val1, val2)
	If var = True Then
		IIf = val1
	 Else
		IIf = val2
	End If
End Function
'===========================================
'读取Access表的字段属性 TableName: 表名
'深山老熊 81090
'===========================================
Sub GetFieldsInfo(TableName)
        Dim oRs,o,Key
        Dim AccessFieldType

        Set AccessFieldType = Server.CreateObject( "Scripting.Dictionary" )
        With AccessFieldType
                .CompareMode = 1   
                .Add "2","整型"
                .Add "3","长整型"
                .Add "4","单精浮点"
                .Add "5","双精浮点"
                .Add "6","货币"
                .Add "7","日期/时间"
                .Add "11","是/否"
                .Add "17","字节"
                .Add "72","同步复制ID"
                .Add "131","小数"
                .Add "135","日期/时间"
                .Add "202","文本"
                .Add "203","备注"
                .Add "205","OLE对象"
        End With

        Set o   = GetDescriptionInfo( TableName )
        Set oRs = conn.Execute("Select * From [" & TableName & "] where 1=0")

        For Each Key In oRs.Fields
                '循环取各字段属性,值见如下注释:
                '--------------------------------
                '表    名: TableName
                '字段名称: Key.Name
                '类    型: Key.Type
                '类型名称: AccessFieldType(CStr(Key.Type))
                '长    度: Key.DefinedSize
                '默 认 值: o(Key.Name)(0)
                '字段描述: o(Key.Name)(1)
                '允 许 空: o(Key.Name)(2)  'True:1 or False:0
                '标    识: o(Key.Name)(3)  'True:1 or False:0
                '--------------------------------
                '保存到表或输出,略了,自己写
                '--------------------------------
				response.write o(Key.Name)(1)&"<br>"
        Next
        oRs.Close
        Set oRs = Nothing
        Set o   = Nothing
        Set AccessFieldType = Nothing
End Sub

'获取所有字段如下属性: Default(默认值) , Description(描述) , Nullable (允许空) , Autoincrement ( 标识 )
'通过 Dictionary 返回 [ 提高效率,一个表只读取一次 ]
Function GetDescriptionInfo( TableName )
        Dim OutField
        Dim MyDB,MyTable
        Dim Key,Key1
        Set OutField = Server.CreateObject( "Scripting.Dictionary" )
        OutField.CompareMode = 1   

        Set MyDB    = Server.CreateObject("ADOX.Catalog")
        Set MyTable = Server.CreateObject("ADOX.Table")

        MyDB.ActiveConnection = Conn         '数据库连接,自己写哈~
        Set MyTable = MyDB.Tables(TableName)
       
        For Each Key In MyTable.Columns
                OutField.Add Key.Name,Array(Key.Properties("Default") & "",Key.Properties("Description") & "",IIf(LCase(Key.Properties("Nullable") & "") = "true",1,0),IIf(LCase(Key.Properties("Autoincrement") & "") = "true",1,0))
        Next

        Set GetDescriptionInfo = OutField
        Set OutField = Nothing
        Set MyTable = nothing
        Set MyDB = Nothing
End Function


Function getDesc(strTableName, strColName) 
Dim cat 
Set cat = server.CreateObject( "ADOX.Catalog") 
cat.ActiveConnection = conn
getDesc = cat.Tables(strTableName).Columns(strColName).Properties( "Description").Value 
Set cat = Nothing 
End Function 

%><!--#Include File="../Code.asp"-->