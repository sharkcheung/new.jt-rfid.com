function chkform(){
	var a=document.loginfrm;
	    if(a.login_name.value==""){
		   alert("��Ա�˺Ų���Ϊ��!");
		   a.login_name.focus();
		   return false;
		}
		if(a.login_name.value.length<4 || a.login_name.value.length>10){
		   alert("����ȷ�����Ա�˺�!(2-12λ)");
		   a.login_name.focus();
		   return false;
		}
	    if(a.login_pwd.value==""){
		   alert("��Ա���벻��Ϊ��!");
		   a.login_pwd.focus();
		   return false;
		}
		if(a.login_pwd.value.length<6 || a.login_pwd.value.length>16){
		   alert("����ȷ�����Ա����!(6-16λ)");
		   a.login_pwd.focus();
		   return false;
		}
	    if(a.login_code.value==""){
		   alert("��֤�벻��Ϊ��!");
		   a.login_code.focus();
		   return false;
		}
		if(a.login_code.value.length!=4){
		   alert("����ȷ������֤��!(4λ)");
		   a.login_code.focus();
		   return false;
		}
	}



//����ajax����
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

//ajax����ǿ����֤
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
  if(iN>=48 && iN<=57)//����
   return 1;
  if(iN>=65 && iN<=90)//��д��ĸ
   return 2;
  if(iN>=97 && iN<=122)//Сд
   return 4;
  else
   return 8;//�����ַ�
 }

 //�������ǰ���뵱��һ���ж�����ģʽ
 function bitTotal(num){
  var modes=0;
  for(i=0;i<4;i++){
   if(num&1)
    modes++;
   num >>=1;
  }
  return modes;
 }
  
 //���������ǿ�ȼ���
 function checkStrong(sPW){
  if(sPW.length<6)
   return 0;//����̫�� 
  var Modes=0;
  for(i=0;i<sPW.length;i++){
  //����ÿһ���ַ������ͳ��һ���ж�����ģʽ.
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
 


//ajax��֤����֤
function isCheckCode(){
var CheckCode = document.getElementById('CheckCode').value;
if (CheckCode==""){
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg>��֤�벻��Ϊ�գ�</span>';
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
if(name_re.readyState==4 && name_re.status==200)//�������
{
var msg=name_re.responseText;
if (msg=="yes"){
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg>��֤�����</span>';
Code_msg(0);
return false;
}
else{
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg2>������ȷ��</span>';
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
//ajax�û�����֤
function isName(){
var reg=/^[a-zA-Z][a-zA-Z0-9_]{3,9}$/;
var u_name = document.getElementById('u_name').value;
if (u_name==""){
document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('name_re_m').innerHTML='<span class=msg>�û�������Ϊ�գ���ֻ����ĸ��ͷ,Ӣ�ġ����֡��»��ߡ�4-10���ַ�</span>';
return false;
}else {
	
	if (!chk_RegEx(reg,u_name)){
         document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
         document.getElementById('name_re_m').innerHTML='<span class=msg>�û�����ֻ����ĸ��ͷ,Ӣ�ġ����֡��»��ߡ�4-10���ַ�</span>';
		 return false;
		}
		else{
           user_ajax(u_name);
		}
}
}
function showinfos_name(){
document.getElementById('name_re').innerHTML='';
document.getElementById('name_re_m').innerHTML='* ��ĸ��ͷ,Ӣ�ġ����֡��»���(4��10���ַ�)';
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
if(name_re.readyState==4 && name_re.status==200)//�������
{
var msg=name_re.responseText;
if (msg=="yes"){
document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('name_re_m').innerHTML='<span class=msg>���û����Ѿ����ڣ�</span>';
name_msg(0);
return false;
}
else{
document.getElementById('name_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('name_re_m').innerHTML='<span class=msg2>����ע�ᣡ</span>';
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


//�Ա��Ƿ�ѡ����
function sex(){
if(document.regform.u_sex[0].checked==false && document.regform.u_sex[1].checked==false && document.regform.u_sex[2].checked==false) {
document.getElementById('sex_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('sex_re_m').innerHTML='<span class=msg>û��ѡ���Ա�</span>';
return false;
} 
else{
document.getElementById('sex_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('sex_re_m').innerHTML='<span class=msg2>��ѡ��</span>';
return true;
}
}

//�����Ƿ�Ϊ�ռ��
function password(){
var u_pass = document.getElementById('u_pass').value;
if(u_pass=="" || u_pass.length<6){
document.getElementById('pass_re').innerHTML='<img src=images/check_error.gif> <span class=msg>����ȷ��������6��16λ</span>';
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
       document.getElementById('pass_re').innerHTML='* 6��16λ';
	   }
	}


//ȷ��������
function pass_re(){
var u_pass=document.getElementById('u_pass').value;
var pass_re=document.getElementById('u_pass_re').value;
if(u_pass != pass_re){
document.getElementById('pass_re_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('pass_re_re_m').innerHTML='<span class=msg>�������벻һ�£�����������</span>';
return false;
}
else{
	if(u_pass.length>=6){
document.getElementById('pass_re_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('pass_re_re_m').innerHTML='<span class=msg2>��д��ȷ</span>';
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

//���뱣��������
function answer(){
var u_answer=document.getElementById('u_answer').value; 
if(u_answer==""){
document.getElementById('answer_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('answer_re_m').innerHTML='<span class=msg>����д�����</span>';
return false;
}
else{
document.getElementById('answer_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('answer_re_m').innerHTML='<span class=msg2>��д��ȷ</span>';
return true;
}
}

function showinfos_answer(){
	   if(document.getElementById('u_answer').value==""){
              document.getElementById('answer_re').innerHTML='';
		      document.getElementById('answer_re_m').innerHTML='* ��������Ĵ𰸣��һ�����ʱ��';
		   }
	}

//�����ʽ��֤
function isEmail() {
var u_mail=document.getElementById('u_mail').value;
if (u_mail.search(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/) != -1){
email_ajax(u_mail);
}
else{
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>��������ȷ�������ַ����ʽΪ��123456@qq.com</span>';
return false;
}
}

function showinfos_email(){
	   	if(document.getElementById('u_mail').value==""){
           document.getElementById('mail_re').innerHTML='* ��ʽ��123456@qq.com';
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
if(name_re.readyState==4 && name_re.status==200)//�������
{
var msg=name_re.responseText;
if (msg=="yes"){
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>�������ѱ�ʹ�ã��뻻һ��</span>';
mail_msg(0);
return false;
}
else{
document.getElementById('mail_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg2>����ʹ��</span>';
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

//��ʵ�������
function name_zs(){
var name_zs=document.getElementById('u_name_zs').value; 
if (name_zs != name_zs.replace(/[^\u4E00-\u9FA5]/g,'') || name_zs=="" || name_zs.length<2 || name_zs.replace(/[^\u4E00-\u9FA5]/g,'').length>4){
document.getElementById('name_zs_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('name_zs_re_m').innerHTML='<span class="msg">��������ʵ����������(2-4�������ַ�)</span>';
return false;
}
else{
document.getElementById('name_zs_re').innerHTML='<img src=images/check_right.gif>';
document.getElementById('name_zs_re_m').innerHTML='<span class="msg2">��д��ȷ</span>';
return true;
}
}

function showinfosname_zs(){
	   	if(document.getElementById('u_name_zs').value==""){
       document.getElementById('name_zs_re').innerHTML='';
       document.getElementById('name_zs_re_m').innerHTML='* ��ʵ����������(2-4�����ĺ���)';
	   }
	}
	
function tel(){
	   var member_tel=document.getElementById('member_tel').value; 
	   var reg=/\d{3}-\d{8}|\d{4}-\d{7}/;
	   if(member_tel!=""){
	   if(!chk_RegEx(reg,member_tel)){
          document.getElementById('tel_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('tel_re_m').innerHTML='<span class=msg>����ȷ��д�绰����(��ʽ:0755-88888888��020-88888888)</span>';
		  return false;
		   }
		   else{
              document.getElementById('tel_re').innerHTML='<img src=images/check_right.gif>';
              document.getElementById('tel_re_m').innerHTML='<span class=msg2>��д��ȷ</span>';
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
          document.getElementById('mobile_re_m').innerHTML='<span class=msg>����ȷ��д�ֻ�����</span>';
		  return false;
		   }
		   else{
              document.getElementById('mobile_re').innerHTML='<img src=images/check_right.gif>';
              document.getElementById('mobile_re_m').innerHTML='<span class=msg2>��д��ȷ</span>';
			  return true;
			   }
	   }
	   else{
              document.getElementById('mobile_re').innerHTML='';
              document.getElementById('mobile_re_m').innerHTML='';
		   }
	}
	
//QQ������
function zip(){
var zip=document.getElementById('member_zip').value;
var reg=/[1-9]\d{5}(?!\d)/;
if (zip!=""){
   if(!chk_RegEx(reg,zip) || zip.length>10){
      document.getElementById('zip_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('zip_re_m').innerHTML='<span class="msg">����ȷ�����ʱ�(6λ)</span>';
      return false;
   }
   else{
      document.getElementById('zip_re').innerHTML='<img src=images/check_right.gif>';
      document.getElementById('zip_re_m').innerHTML='<span class="msg2">��д��ȷ</span>';
      return true;
   }
}
else{
   document.getElementById('zip_re').innerHTML='';
   document.getElementById('zip_re_m').innerHTML='';
   }
}
//QQ������
function qq(){
var qq=document.getElementById('u_qq').value;
var reg=/[1-9][0-9]{4,}/;
if (qq!=""){
   if(!chk_RegEx(reg,qq) || qq.length>10){
      document.getElementById('qq_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('qq_re_m').innerHTML='<span class="msg">��ȷ��QQ��5-10λŶ</span>';
      return false;
   }
   else{
      document.getElementById('qq_re').innerHTML='<img src=images/check_right.gif>';
      document.getElementById('qq_re_m').innerHTML='<span class="msg2">��д��ȷ</span>';
      return true;
   }
}
else{
   document.getElementById('qq_re').innerHTML='';
   document.getElementById('qq_re_m').innerHTML='';
   }
}

//���֧�����ʺ�
//function alipay(){
//var alipay=document.getElementById('u_alipay').value;
//if (alipay.search(/^\w+((-\w+)|(\.\w+))*\@[A-Za-z0-9]+((\.|-)[A-Za-z0-9]+)*\.[A-Za-z0-9]+$/) != -1){
//document.getElementById('alipay_re').innerHTML='<img src=images/check_right.gif>';
//document.getElementById('alipay_re_m').innerHTML='<span class=msg2>��д��ȷ</span>';
//return true;
//}
//else{
//document.getElementById('alipay_re').innerHTML='<img src=images/check_error.gif>';
//document.getElementById('alipay_re_m').innerHTML='<span class=msg>�����֧�����ʺ�</span>';
//return false;
//}
//}

//���֤������
//function nunber(){
//var idcard=document.getElementById('u_nunber').value;
//var Errors=new Array("��֤ͨ��!","���֤����λ������!","�������ڳ�����Χ���зǷ��ַ�!","���֤����У�����!","���֤�����Ƿ�!");
//var area={11:"����",12:"���",13:"�ӱ�",14:"ɽ��",15:"���ɹ�",21:"����",22:"����",23:"������",31:"�Ϻ�",32:"����",33:"�㽭",34:"����",35:"����",36:"����",37:"ɽ��",41:"����",42:"����",43:"����",44:"�㶫",45:"����",46:"����",50:"����",51:"�Ĵ�",52:"����",53:"����",54:"����",61:"����",62:"����",63:"�ຣ",64:"����",65:"�½�",71:"̨��",81:"���",82:"����",91:"����"}
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
//   //18λ��ݺ�����
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

//ȫ���ύ��֤
function tijiao(){
if (isName()==false){
	document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
    document.getElementById('name_re_m').innerHTML='<span class=msg>�û�������Ϊ�գ���ֻ����ĸ��ͷ,Ӣ�ġ����֡��»��ߡ�4-10���ַ�</span>';
    //alert("�û�����д����ȷ");
    return false;
}
if (name_use==true){
	document.getElementById('name_re').innerHTML='<img src=images/check_error.gif>';
    document.getElementById('name_re_m').innerHTML='<span class=msg>���û����Ѿ����ڣ�</span>';
    //alert("�û����Ѵ��ڣ���������");
return false;
}

if(name_zs()==false){
    document.getElementById('name_zs_re').innerHTML='<img src=images/check_error.gif>';
    document.getElementById('name_zs_re_m').innerHTML='<span class="msg">������д����(2-4�������ַ�)</span>';
    //alert("������д����");
    return false;
    }
if (sex()==false){
alert("��ѡ������Ա�");
return false;
}
if (password()==false){
   document.getElementById('pass_re').innerHTML='<img src=images/check_error.gif> <span class=msg>����ȷ��������6��16λ</span>';
//alert("���������д");
return false;
}
if (pass_re()==false){
   document.getElementById('u_pass_re').innerHTML='<img src=images/check_error.gif> <span class=msg>ȷ���������</span>';
   //alert("ȷ���������");
return false;
}
if (answer()==false){
document.getElementById('answer_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('answer_re_m').innerHTML='<span class=msg>����д�����</span>';
//alert("��ȫ����𰸱�����д");
return false;
}
if (isEmail()==false){
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>�����ַΪ�ջ��ߴ��󣬸�ʽΪ��123456@qq.com</span>';
//alert("�����ַΪ�ջ��ߴ���");
return false;
}
if (mail_use==true){
document.getElementById('mail_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('mail_re_m').innerHTML='<span class=msg>�����Ѿ����ڣ���������һ��</span>';
//alert("�����Ѿ����ڣ���������һ��");
return false;
}
if(document.getElementById("member_tel").value!=""){
	   if(tel()==false){
          document.getElementById('tel_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('tel_re_m').innerHTML='<span class=msg>����ȷ��д�绰����(��ʽ:0755-88888888��020-88888888)</span>';
		      //alert("�绰��д����!");
			  return false;
		   }
	}
if(document.getElementById("member_mobile").value!=""){
	   if(mobile()==false){
          document.getElementById('mobile_re').innerHTML='<img src=images/check_error.gif>';
          document.getElementById('mobile_re_m').innerHTML='<span class=msg>����ȷ��д�ֻ�����</span>';
		      //alert("�ֻ���д����!");
			  return false;
		   }
	}
if(document.getElementById("member_zip").value!=""){
	   if(zip()==false){
      document.getElementById('zip_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('zip_re_m').innerHTML='<span class="msg">����ȷ�����ʱ�(6λ)</span>';
		      //alert("�ʱ���д����!");
			  return false;
		   }
	}
if(document.getElementById('u_qq').value !=""){
if(qq()==false){
      document.getElementById('qq_re').innerHTML='<img src=images/check_error.gif>';
      document.getElementById('qq_re_m').innerHTML='<span class="msg">��ȷ��QQ��5-10λŶ</span>';
   //alert("qq������д����");
   return false;
   }
}
if(isCheckCode()==false){
document.getElementById('CheckCode_re').innerHTML='<img src=images/check_error.gif>';
document.getElementById('CheckCode_re_m').innerHTML='<span class=msg>��֤��Ϊ�ջ����</span>';
    //alert("��֤��Ϊ�ջ����");
    return false;
}

//if (document.getElementById('u_alipay').value !=""){
//if(alipay()==false){
//alert("֧�����ʺ���д����");
//return false;
//}
//}
//if (document.getElementById('u_nunber').value !=""){
//if(nunber()==false){
//alert("���֤������д����");
//return false;
//}
//}
document.regform.submit();
return true;
}