function chkform(){
	var a=document.loginfrm;
	    if(a.login_name.value==""){
		   alert("会员账号不能为空!");
		   a.login_name.focus();
		   return false;
		}
		if(a.login_name.value.length<4 || a.login_name.value.length>10){
		   alert("请正确输入会员账号!(2-12位)");
		   a.login_name.focus();
		   return false;
		}
	    if(a.login_pwd.value==""){
		   alert("会员密码不能为空!");
		   a.login_pwd.focus();
		   return false;
		}
		if(a.login_pwd.value.length<6 || a.login_pwd.value.length>16){
		   alert("请正确输入会员密码!(6-16位)");
		   a.login_pwd.focus();
		   return false;
		}
	    if(a.login_code.value==""){
		   alert("验证码不能为空!");
		   a.login_code.focus();
		   return false;
		}
		if(a.login_code.value.length!=4){
		   alert("请正确输入验证码!(4位)");
		   a.login_code.focus();
		   return false;
		}
	}



//创建ajax对象
var name_re = false; 
   function name_xml()
   {
   try { 
     name_re = new XMLHttpRequest(); 
   } catch (trymicrosoft) { 
     try { 
       name_re = new ActiveXObject("Msxml2.XMLHTTP"); 
     } catch (othermicrosoft) { 
       try { 
         name_re = new ActiveXObject("Microsoft.XMLHTTP"); 
       } catch (failed) { 
         name_re = false; 
       }   
     } 
   } 
   if (!name_re) 
     alert("Error initializing XMLHttpRequest!"); 
}

var name_use;
var mail_use;

//ajax密码强度验证
function allNumber(v)
 {
  var reg = /^[0-9]*$/;
  if(reg.test(v))
  { 
   return true;
  }
  return false;
 }
 
 function CharMode(iN){
  if(iN>=48 && iN<=57)//数字
   return 1;
  if(iN>=65 && iN<=90)//大写字母
   return 2;
  if(iN>=97 && iN<=122)//小写
   return 4;
  else
   return 8;//特殊字符
 }

 //计算出当前密码当中一共有多少种模式
 function bitTotal(num){
  var modes=0;
  for(i=0;i<4;i++){
   if(num&1)
    modes++;
   num >>=1;
  }
  return modes;
 }
  
 //返回密码的强度级别
 function checkStrong(sPW){
  if(sPW.length<6)
   return 0;//密码太短 
  var Modes=0;
  for(i=0;i<sPW.length;i++){
  //测试每一个字符的类别并统计一共有多少种模式.
  Modes|=CharMode(sPW.charCodeAt(i));
  }
 // alert(bitTotal(Modes));
  return bitTotal(Modes);
 }


 function showStrongPic()
 {
  var v = document.getElementById('u_pass').value;
  if(v.length>=6){
  var m = checkStrong(v);
  if(m < 2)
  {
   document.getElementById('lowPic').style.display="";
   document.getElementById('midPic').style.display="none";
   document.getElementById('highPic').style.display="none";
  }
  else if(m==2)
  {
   document.getElementById('lowPic').style.display="none";
   document.getElementById('midPic').style.display="";
   document.getElementById('highPic').style.display="none";
  }
  else 
  {
   document.getElementById('lowPic').style.display="none";
   document.getElementById('midPic').style.display="none";
   document.getElementById('highPic').style.display="";
  }
  }
  else{
   document.getElementById('lowPic').style.display="none";
   document.getElementById('midPic').style.display="none";
   document.getElementById('highPic').style.display="none";
	  }
 }
 


//ajax验证码验证
function isCheckCode(){
var CheckCode = document.getElementById('CheckCode').value;
if (CheckCode==""){
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg>验证码不能为空！</span>';
return false;
}else{
Code_ajax(CheckCode)
}
}

function Code_ajax(CheckCode){
var Code=CheckCode;
var url="ajax.asp?act=code&Code="+ escape(Code); 
     name_xml();
     name_re.open("GET", url, true); 
     name_re.onreadystatechange = Code_requst; 
     name_re.send(null); 
}

function Code_requst(){
if(name_re.readyState==4 && name_re.status==200)//返回完成
{
var msg=name_re.responseText;
if (msg=="yes"){
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg>验证码错误！</span>';
Code_msg(0);
return false;
}
else{
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg2>输入正确！</span>';
Code_msg(1);
return true;
}
}
}
function Code_msg(n){
var n=n;
if(n==0){
Code_use=true;
}
else{
Code_use=false;
}
}

function chk_RegEx(reg,str){
   return reg.test(str);
	}
//ajax用户名验证
function isName(){
var reg=/^[a-zA-Z][a-zA-Z0-9_]{3,9}$/;
var u_name = document.getElementById('u_name').value;
if (u_name==""){
document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('name_re_m').innerHTML='<span class=msg>用户名不能为空，且只能字母开头,英文、数字、下划线、4-10个字符</span>';
return false;
}else {
	
	if (!chk_RegEx(reg,u_name)){
         document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
         document.getElementById('name_re_m').innerHTML='<span class=msg>用户名且只能字母开头,英文、数字、下划线、4-10个字符</span>';
		 return false;
		}
		else{
           user_ajax(u_name);
		}
}
}
function showinfos_name(){
document.getElementById('name_re').innerHTML='';
document.getElementById('name_re_m').innerHTML='* 字母开头,英文、数字、下划线(4～10个字符)';
	}

function user_ajax(u_name){
var name=u_name;
var url="ajax.asp?act=names&name="+ escape(name); 
     name_xml();
     name_re.open("GET", url, true); 
     name_re.onreadystatechange = name_requst; 
     name_re.send(null); 
}

function name_requst(){
if(name_re.readyState==4 && name_re.status==200)//返回完成
{
var msg=name_re.responseText;
if (msg=="yes"){
document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('name_re_m').innerHTML='<span class=msg>该用户名已经存在！</span>';
name_msg(0);
return false;
}
else{
document.getElementById('name_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('name_re_m').innerHTML='<span class=msg2>可以注册！</span>';
name_msg(1);
return true;
}
}
}
function name_msg(n){
var n=n;
if(n==0){
name_use=true;
}
else{
name_use=false;
}
}


//性别是否选择检测
function sex(){
if(document.regform.u_sex[0].checked==false && document.regform.u_sex[1].checked==false && document.regform.u_sex[2].checked==false) {
document.getElementById('sex_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('sex_re_m').innerHTML='<span class=msg>没有选择性别</span>';
return false;
} 
else{
document.getElementById('sex_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('sex_re_m').innerHTML='<span class=msg2>已选择</span>';
return true;
}
}

//密码是否为空检测
function password(){
var u_pass = document.getElementById('u_pass').value;
if(u_pass=="" || u_pass.length<6){
document.getElementById('pass_re').innerHTML='<img src=images/check_error.gif> <span class=msg>请正确输入密码6～16位</span>';
return false;
}
else{
document.getElementById('pass_re').innerHTML='';
document.getElementById('pass_re').innerHTML='<img src=images/check_right.gif>';
return true;
}
}
function showinfos_pass(){
	if(document.getElementById('u_pass').value==""){
       document.getElementById('pass_re').innerHTML='* 6～16位';
	   }
	}


//确认密码检测
function pass_re(){
var u_pass=document.getElementById('u_pass').value;
var pass_re=document.getElementById('u_pass_re').value;
if(u_pass != pass_re){
document.getElementById('pass_re_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('pass_re_re_m').innerHTML='<span class=msg>两次密码不一致，请重新输入</span>';
return false;
}
else{
	if(u_pass.length>=6){
document.getElementById('pass_re_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('pass_re_re_m').innerHTML='<span class=msg2>填写正确</span>';
return true;
	}
}
}

function showinfos_pass_re(){
	   	if(document.getElementById('u_pass_re').value==""){
       document.getElementById('pass_re_re').innerHTML='';
       document.getElementById('pass_re_re_m').innerHTML='';
	   }
	}

//密码保护问题检测
function answer(){
var u_answer=document.getElementById('u_answer').value; 
if(u_answer==""){
document.getElementById('answer_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('answer_re_m').innerHTML='<span class=msg>请填写问题答案</span>';
return false;
}
else{
document.getElementById('answer_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('answer_re_m').innerHTML='<span class=msg2>填写正确</span>';
return true;
}
}

function showinfos_answer(){
	   if(document.getElementById('u_answer').value==""){
              document.getElementById('answer_re').innerHTML='';
		      document.getElementById('answer_re_m').innerHTML='* 上面问题的答案，找回密码时用';
		   }
	}

//邮箱格式验证
function isEmail() {
var u_mail=document.getElementById('u_mail').value;
if (u_mail.search(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/) != -1){
email_ajax(u_mail);
}
else{
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>请输入正确的邮箱地址，格式为：123456@qq.com</span>';
return false;
}
}

function showinfos_email(){
	   	if(document.getElementById('u_mail').value==""){
           document.getElementById('mail_re').innerHTML='* 格式：123456@qq.com';
           document.getElementById('mail_re_m').innerHTML='';
	   }
	}

function email_ajax(u_mail){
var email=u_mail;
var url="ajax.asp?act=email&mail="+ escape(email); 
     name_xml();
     name_re.open("GET", url, true); 
     name_re.onreadystatechange = mail_requst; 
     name_re.send(null); 

}

function mail_requst(){
if(name_re.readyState==4 && name_re.status==200)//返回完成
{
var msg=name_re.responseText;
if (msg=="yes"){
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>该邮箱已被使用，请换一个</span>';
mail_msg(0);
return false;
}
else{
document.getElementById('mail_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg2>可以使用</span>';
mail_msg(1);
return true;
}
}
}
function mail_msg(n){
var n=n;
if(n==0){
mail_use=true;
}
else{
mail_use=false; 
}
}

//真实姓名检测
function name_zs(){
var name_zs=document.getElementById('u_name_zs').value; 
if (name_zs != name_zs.replace(/[^\u4E00-\u9FA5]/g,'') || name_zs=="" || name_zs.length<2 || name_zs.replace(/[^\u4E00-\u9FA5]/g,'').length>4){
document.getElementById('name_zs_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('name_zs_re_m').innerHTML='<span class="msg">请输入真实的中文名字(2-4个中文字符)</span>';
return false;
}
else{
document.getElementById('name_zs_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('name_zs_re_m').innerHTML='<span class="msg2">填写正确</span>';
return true;
}
}

function showinfosname_zs(){
	   	if(document.getElementById('u_name_zs').value==""){
       document.getElementById('name_zs_re').innerHTML='';
       document.getElementById('name_zs_re_m').innerHTML='* 真实的中文名字(2-4个中文汉字)';
	   }
	}
	
function tel(){
	   var member_tel=document.getElementById('member_tel').value; 
	   var reg=/\d{3}-\d{8}|\d{4}-\d{7}/;
	   if(member_tel!=""){
	   if(!chk_RegEx(reg,member_tel)){
          document.getElementById('tel_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('tel_re_m').innerHTML='<span class=msg>请正确填写电话号码(格式:0755-88888888或020-88888888)</span>';
		  return false;
		   }
		   else{
              document.getElementById('tel_re').innerHTML='<img src=images/check_right.gif>';
              document.getElementById('tel_re_m').innerHTML='<span class=msg2>填写正确</span>';
			  return true;
			   }
	   }
	   else{
              document.getElementById('tel_re').innerHTML='';
              document.getElementById('tel_re_m').innerHTML='';
		   }
	}
	
function mobile(){
	   var member_mobile=document.getElementById('member_mobile').value; 
	   var reg= /^1[3,5,8]\d{9}$/;
	   if(member_mobile!=""){
	   if(!chk_RegEx(reg,member_mobile)){
          document.getElementById('mobile_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('mobile_re_m').innerHTML='<span class=msg>请正确填写手机号码</span>';
		  return false;
		   }
		   else{
              document.getElementById('mobile_re').innerHTML='<img src=images/check_right.gif>';
              document.getElementById('mobile_re_m').innerHTML='<span class=msg2>填写正确</span>';
			  return true;
			   }
	   }
	   else{
              document.getElementById('mobile_re').innerHTML='';
              document.getElementById('mobile_re_m').innerHTML='';
		   }
	}
	
//QQ号码检测
function zip(){
var zip=document.getElementById('member_zip').value;
var reg=/[1-9]\d{5}(?!\d)/;
if (zip!=""){
   if(!chk_RegEx(reg,zip) || zip.length>10){
      document.getElementById('zip_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('zip_re_m').innerHTML='<span class="msg">请正确输入邮编(6位)</span>';
      return false;
   }
   else{
      document.getElementById('zip_re').innerHTML='<img src=images/check_right.gif>';
      document.getElementById('zip_re_m').innerHTML='<span class="msg2">填写正确</span>';
      return true;
   }
}
else{
   document.getElementById('zip_re').innerHTML='';
   document.getElementById('zip_re_m').innerHTML='';
   }
}
//QQ号码检测
function qq(){
var qq=document.getElementById('u_qq').value;
var reg=/[1-9][0-9]{4,}/;
if (qq!=""){
   if(!chk_RegEx(reg,qq) || qq.length>10){
      document.getElementById('qq_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('qq_re_m').innerHTML='<span class="msg">正确的QQ是5-10位哦</span>';
      return false;
   }
   else{
      document.getElementById('qq_re').innerHTML='<img src=images/check_right.gif>';
      document.getElementById('qq_re_m').innerHTML='<span class="msg2">填写正确</span>';
      return true;
   }
}
else{
   document.getElementById('qq_re').innerHTML='';
   document.getElementById('qq_re_m').innerHTML='';
   }
}

//检测支付宝帐号
//function alipay(){
//var alipay=document.getElementById('u_alipay').value;
//if (alipay.search(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/) != -1){
//document.getElementById('alipay_re').innerHTML='<img src=images/check_right.gif>';
//document.getElementById('alipay_re_m').innerHTML='<span class=msg2>填写正确</span>';
//return true;
//}
//else{
//document.getElementById('alipay_re').innerHTML='<img src=images/check_error.gif>';
//document.getElementById('alipay_re_m').innerHTML='<span class=msg>错误的支付宝帐号</span>';
//return false;
//}
//}

//身份证号码检测
//function nunber(){
//var idcard=document.getElementById('u_nunber').value;
//var Errors=new Array("验证通过!","身份证号码位数不对!","出生日期超出范围或含有非法字符!","身份证号码校验错误!","身份证地区非法!");
//var area={11:"北京",12:"天津",13:"河北",14:"山西",15:"内蒙古",21:"辽宁",22:"吉林",23:"黑龙江",31:"上海",32:"江苏",33:"浙江",34:"安徽",35:"福建",36:"江西",37:"山东",41:"河南",42:"湖北",43:"湖南",44:"广东",45:"广西",46:"海南",50:"重庆",51:"四川",52:"贵州",53:"云南",54:"西藏",61:"陕西",62:"甘肃",63:"青海",64:"宁夏",65:"新疆",71:"台湾",81:"香港",82:"澳门",91:"国外"}
//
//var idcard,Y,JYM;
//var S,M;
//var idcard_array = new Array();
//idcard_array = idcard.split("");
//if(area[parseInt(idcard.substr(0,2))]==null) 
//{
//document.getElementById('nunber_re').innerHTML='<img src=check_error.gif>';
//     document.getElementById('nunber_re_m').innerHTML="<span class=msg>"+Errors[4]+"</span>";
//return false;
//}
//
//switch(idcard.length){
//   case 15:
//   if ( (parseInt(idcard.substr(6,2))+1900) % 4 == 0 || ((parseInt(idcard.substr(6,2))+1900) % 100 == 0 && (parseInt(idcard.substr(6,2))+1900) % 4 == 0 )){
//    ereg=/^[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}$/;
//   } else {
//    ereg=/^[1-9][0-9]{5}[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}$/;
//   }
//   if(ereg.test(idcard)){
//document.getElementById('nunber_re').innerHTML='<img src=check_right.gif>';
//     document.getElementById('nunber_re_m').innerHTML="<span class=msg2>"+Errors[0]+"</span>";
//return true;
//    }
//   else {
//document.getElementById('nunber_re').innerHTML='<img src=check_error.gif>';
//     document.getElementById('nunber_re_m').innerHTML="<span class=msg>"+Errors[2]+"</span>";
//return false;
//     }
//   break; 
//   case 18:
//   //18位身份号码检测
//   if ( parseInt(idcard.substr(6,4)) % 4 == 0 || (parseInt(idcard.substr(6,4)) % 100 == 0 && parseInt(idcard.substr(6,4))%4 == 0 )){
//   ereg=/^[1-9][0-9]{5}19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))[0-9]{3}[0-9Xx]$/;
//   } else {
//   ereg=/^[1-9][0-9]{5}19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|1[0-9]|2[0-8]))[0-9]{3}[0-9Xx]$/;
//   }
//   if(ereg.test(idcard)){
//    S = (parseInt(idcard_array[0]) + parseInt(idcard_array[10])) * 7
//    + (parseInt(idcard_array[1]) + parseInt(idcard_array[11])) * 9
//    + (parseInt(idcard_array[2]) + parseInt(idcard_array[12])) * 10
//    + (parseInt(idcard_array[3]) + parseInt(idcard_array[13])) * 5
//    + (parseInt(idcard_array[4]) + parseInt(idcard_array[14])) * 8
//    + (parseInt(idcard_array[5]) + parseInt(idcard_array[15])) * 4
//    + (parseInt(idcard_array[6]) + parseInt(idcard_array[16])) * 2
//    + parseInt(idcard_array[7]) * 1 
//    + parseInt(idcard_array[8]) * 6
//    + parseInt(idcard_array[9]) * 3 ;
//    Y = S % 11;
//    M = "F";
//    JYM = "10X98765432";
//    M = JYM.substr(Y,1);
//    if(M == idcard_array[17]){
//document.getElementById('nunber_re').innerHTML='<img src=images/check_right.gif>';
//     document.getElementById('nunber_re_m').innerHTML="<span class=msg2>"+Errors[0]+"</span>";
//     return true;
//    }
//    else {
//document.getElementById('nunber_re').innerHTML='<img src=images/check_error.gif>';
//     document.getElementById('nunber_re_m').innerHTML="<span class=msg>"+Errors[3]+"</span>";
//     return false;
//    }
//   }
//   else {
//document.getElementById('nunber_re').innerHTML='<img src=images/check_error.gif>';
//    document.getElementById('nunber_re_m').innerHTML="<span class=msg>"+Errors[2]+"</span>";
//    return false;
//   }
//   break;
//   default: 
//document.getElementById('nunber_re').innerHTML='<img src=images/check_error.gif>';
//    document.getElementById('nunber_re_m').innerHTML="<span class=msg>"+Errors[1]+"</span>";
//    return false;
//}
//}

//全表单提交验证
function tijiao(){
if (isName()==false){
	document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
    document.getElementById('name_re_m').innerHTML='<span class=msg>用户名不能为空，且只能字母开头,英文、数字、下划线、4-10个字符</span>';
    //alert("用户名填写不正确");
    return false;
}
if (name_use==true){
	document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
    document.getElementById('name_re_m').innerHTML='<span class=msg>该用户名已经存在！</span>';
    //alert("用户名已存在，重新输入");
return false;
}

if(name_zs()==false){
    document.getElementById('name_zs_re').innerHTML='<img src=images/check_error.gif>';
    document.getElementById('name_zs_re_m').innerHTML='<span class="msg">姓名填写错误(2-4个中文字符)</span>';
    //alert("姓名填写错误");
    return false;
    }
if (sex()==false){
alert("请选择你的性别");
return false;
}
if (password()==false){
   document.getElementById('pass_re').innerHTML='<img src=images/check_error.gif> <span class=msg>请正确输入密码6～16位</span>';
//alert("密码必须填写");
return false;
}
if (pass_re()==false){
   document.getElementById('u_pass_re').innerHTML='<img src=images/check_error.gif> <span class=msg>确认密码错误</span>';
   //alert("确认密码错误");
return false;
}
if (answer()==false){
document.getElementById('answer_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('answer_re_m').innerHTML='<span class=msg>请填写问题答案</span>';
//alert("安全问题答案必须填写");
return false;
}
if (isEmail()==false){
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>邮箱地址为空或者错误，格式为：123456@qq.com</span>';
//alert("邮箱地址为空或者错误");
return false;
}
if (mail_use==true){
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>邮箱已经存在，重新输入一个</span>';
//alert("邮箱已经存在，重新输入一个");
return false;
}
if(document.getElementById("member_tel").value!=""){
	   if(tel()==false){
          document.getElementById('tel_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('tel_re_m').innerHTML='<span class=msg>请正确填写电话号码(格式:0755-88888888或020-88888888)</span>';
		      //alert("电话填写错误!");
			  return false;
		   }
	}
if(document.getElementById("member_mobile").value!=""){
	   if(mobile()==false){
          document.getElementById('mobile_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('mobile_re_m').innerHTML='<span class=msg>请正确填写手机号码</span>';
		      //alert("手机填写错误!");
			  return false;
		   }
	}
if(document.getElementById("member_zip").value!=""){
	   if(zip()==false){
      document.getElementById('zip_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('zip_re_m').innerHTML='<span class="msg">请正确输入邮编(6位)</span>';
		      //alert("邮编填写错误!");
			  return false;
		   }
	}
if(document.getElementById('u_qq').value !=""){
if(qq()==false){
      document.getElementById('qq_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('qq_re_m').innerHTML='<span class="msg">正确的QQ是5-10位哦</span>';
   //alert("qq号码填写错误");
   return false;
   }
}
if(isCheckCode()==false){
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg>验证码为空或错误！</span>';
    //alert("验证码为空或错误");
    return false;
}

//if (document.getElementById('u_alipay').value !=""){
//if(alipay()==false){
//alert("支付宝帐号填写错误");
//return false;
//}
//}
//if (document.getElementById('u_nunber').value !=""){
//if(nunber()==false){
//alert("身份证号码填写错误");
//return false;
//}
//}
document.regform.submit();
return true;
}