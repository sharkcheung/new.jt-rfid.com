function initt(num,pri){
			   var allp=num*pri;
			   document.getElementById("price_all").innerHTML=allp;
}
function pall(num,pri){
			   var allp=num*pri;
	           var p_p1=document.getElementById("lbl_fee").innerHTML;
			   allp1=(p_p1-0+allp);
			   document.getElementById("price_all").innerHTML=allp1;
}
function change_p(fee,pric){
	   var p_p1=document.getElementById("buy_num").value;
	   var pp_p=(0+p_p1*pric);
	   document.getElementById("lbl_fee").innerHTML=fee;
	   var lbl_fee=document.getElementById("lbl_fee").innerHTML
	   var p_p_p=(lbl_fee-0+pp_p);
	   document.getElementById("price_all").innerHTML=p_p_p;
	}
function checknum(num,kuncun){
	   if(num>kuncun){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="������д����Ʒ����������棡";
		      document.gobuy.buy_num.value="";
			  return false;
		   }
	   else if(!IsNum(num)){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="����д��ȷ����Ʒ������";
			  document.gobuy.buy_num.value="";
			  return false;
			   }
	   else if(num==0){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="����д��ȷ����Ʒ������";
			  document.gobuy.buy_num.value="";
			  return false;
			   }
		else{
			   document.getElementById("numerror").innerHTML="";
			   document.getElementById("numerror").style.display="none";
			   document.all.gobuy.submit();
			   return true;
			}
	}
function checknum2(num,kuncun){
	   if(num>kuncun){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="������д����Ʒ����������棡";
		      document.gobuy.buy_num.value="";
			  return false;
		   }
	   else if(!IsNum(num)){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="����д��ȷ����Ʒ������";
			  document.gobuy.buy_num.value="";
			  return false;
			   }
	   else if(num==0){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="����д��ȷ����Ʒ������";
			  document.gobuy.buy_num.value="";
			  return false;
			   }
		else{
			   document.getElementById("numerror").innerHTML="";
			   document.getElementById("numerror").style.display="none";
			   return true;
			}
	}
function checknum1(num,kuncun,pri){
	   if(num>kuncun){
			  document.getElementById("numerror").style.display="block";
		      document.gobuy.buy_num.value="";
			  document.getElementById("numerror").innerHTML="������д����Ʒ����������棡";
			   document.getElementById("price_all").innerHTML="";
			  return false;
		   }
	   else if(!IsNum(num)){
			  document.getElementById("numerror").style.display="block";
			  document.gobuy.buy_num.value="";
			  document.getElementById("numerror").innerHTML="����д��ȷ����Ʒ������";
			   document.getElementById("price_all").innerHTML="";
			  return false;
			   }
	   else if(num==0){
			  document.getElementById("numerror").style.display="block";
			  document.gobuy.buy_num.value="";
			  document.getElementById("numerror").innerHTML="����д��ȷ����Ʒ������";
			   document.getElementById("price_all").innerHTML="";
			  return false;
			   }
		else{
			   var allp=num*pri;
	           var p_p1=document.getElementById("lbl_fee").innerHTML;
			   allp1=(p_p1-0+allp);
			   document.getElementById("price_all").innerHTML=allp1;
			   document.getElementById("numerror").style.display="none";
			   return true;
			}
	}
function IsNum(num){
  var reNum=/^\d*$/;
  return(reNum.test(num));
}
function chk_order(){
	var a=document.gobuy;
	    if(a.hukouprovince.value==""){
		   alert("ʡ����Ϊ��!");
		   a.hukouprovince.focus();
		   return false;
		}
	    if(a.hukoucapital.value==""){
		   alert("�в���Ϊ��!");
		   a.hukoucapital.focus();
		   return false;
		}
	    if(a.hukoucity.value==""){
		   alert("�ز���Ϊ��!");
		   a.hukoucity.focus();
		   return false;
		}
	    if(a.zip.value==""){
		   alert("�ʱ಻��Ϊ��!");
		   a.zip.focus();
		   return false;
		}
	    if(!IsNum(a.zip.value) || a.zip.value.length>6){
		   alert("����ȷ�����ʱ�!");
		   a.zip.focus();
		   return false;
		}
		if(a.detail_address.value.length==""){
		   alert("����д��ϸ��ַ��");
		   a.detail_address.focus();
		   return false;
		}
	    if(a.shouhuo_name.value==""){
		   alert("�ջ��˲���Ϊ��!");
		   a.shouhuo_name.focus();
		   return false;
		}
		if(a.shouhuo_mobile.value.length=="" && a.shouhuo_tel.value.length==""){
		   alert("�ֻ��͵绰������һ��");
		   a.shouhuo_mobile.focus();
		   return false;
		}
flagt=false;
  for(i=0;i<a.mainRadio.length;i++)
     a.mainRadio[i].checked?flagt=true:'';
     if(!flagt){
        alert("��ѡ��֧����ʽ!");
        a.mainRadio[0].focus();
        return false;
}
	    if(a.buy_num.value==""){
		   alert("������������Ϊ��!");
		   a.buy_num.focus();
		   return false;
		}
	    if(a.songhuofangshi.value==""){
		   alert("�ͻ���ʽ����Ϊ��!");
		   a.songhuofangshi.focus();
		   return false;
		}
/*
if(a.mainRadio[0].checked){
flagt1=false;
  for(i=0;i<a.subRadio.length;i++)
     a.subRadio[i].checked?flagt1=true:'';
     if(!flagt1){
        alert("��ѡ��֧����ʽ!");
        a.mainRadio[0].focus();
        return false;
     }
}
a.action='buy_save.asp';
a.submit();*/
return true;
}
function goto_cart(){
   document.all.gobuy.action="member/add_cart.asp"
   document.all.gobuy.submit();
	}