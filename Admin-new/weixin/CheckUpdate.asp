<%
if isobject(conn) then
	on error resume next
	rs.open "weixin_config",conn
	if not err.number=0 then 
		Err.Clear
		conn.execute("create table weixin_config (id integer identity(1,1) primary key,wx_url varchar(100) null,wx_token varchar(100) null,wx_raw_id varchar(50) null,wx_AppId varchar(50) null,wx_AppSecret varchar(50) null,wx_Random integer default 0,wx_Subscribe varchar(50) null,wx_Repeat integer default 0)" )
		
		conn.execute("create table weixin_menu (id integer identity(1,1) primary key,menuName varchar(50) null,menuType varchar(50) null,menuOnEvent varchar(50) null,menuPx integer default 0,menuStatus integer default 0,menuParent integer default 0)" )
		
		conn.execute("create table weixin_imageText (id integer identity(1,1) primary key,imgText_Title varchar(50) null,imgText_Pic varchar(255) null,imgText_Id_List varchar(50) null,imgText_url varchar(255) null,imgText_px integer default 0,imgText_status integer default 0,imgText_Summary varchar(255) null,imgText_addtime date default now(),imgText_Content text)" )
		
		conn.execute("create table Weixin_CustReply (id integer identity(1,1) primary key,reply_qtitle text,reply_qanswerText text,reply_qanswerNews varchar(50) null,reply_qanswerResource varchar(50) null,reply_type integer default 0,px integer default 0,add_time date default now(),status integer default 0)" )
		
		conn.execute("create table weixin_Sucai (id integer identity(1,1) primary key,Sucai_title varchar(200) null,Sucai_type integer default 0,Sucai_source integer default 0,Sucai_file varchar(255) null,Sucai_desc varchar(255) null,Sucai_px integer default 0,Sucai_status integer default 0,Sucai_fileSize varchar(50) null,Sucai_time date default now())" )
		
		conn.execute("create table weixin_subscribeList (id integer identity(1,1) primary key,openID varchar(50) null,subscribe_time date default now())" )
	
	end if
	rs.close
end if
%>