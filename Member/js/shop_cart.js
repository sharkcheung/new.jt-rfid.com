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
			  document.getElementById("numerror").innerHTML="您所填写的商品数量超过库存！";
		      document.gobuy.buy_num.value="";
			  return false;
		   }
	   else if(!IsNum(num)){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="请填写正确的商品数量！";
			  document.gobuy.buy_num.value="";
			  return false;
			   }
	   else if(num==0){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="请填写正确的商品数量！";
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
			  document.getElementById("numerror").innerHTML="您所填写的商品数量超过库存！";
		      document.gobuy.buy_num.value="";
			  return false;
		   }
	   else if(!IsNum(num)){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="请填写正确的商品数量！";
			  document.gobuy.buy_num.value="";
			  return false;
			   }
	   else if(num==0){
			  document.getElementById("numerror").style.display="block";
			  document.getElementById("numerror").innerHTML="请填写正确的商品数量！";
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
			  document.getElementById("numerror").innerHTML="您所填写的商品数量超过库存！";
			   document.getElementById("price_all").innerHTML="";
			  return false;
		   }
	   else if(!IsNum(num)){
			  document.getElementById("numerror").style.display="block";
			  document.gobuy.buy_num.value="";
			  document.getElementById("numerror").innerHTML="请填写正确的商品数量！";
			   document.getElementById("price_all").innerHTML="";
			  return false;
			   }
	   else if(num==0){
			  document.getElementById("numerror").style.display="block";
			  document.gobuy.buy_num.value="";
			  document.getElementById("numerror").innerHTML="请填写正确的商品数量！";
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
		   alert("省不能为空!");
		   a.hukouprovince.focus();
		   return false;
		}
	    if(a.hukoucapital.value==""){
		   alert("市不能为空!");
		   a.hukoucapital.focus();
		   return false;
		}
	    if(a.hukoucity.value==""){
		   alert("县不能为空!");
		   a.hukoucity.focus();
		   return false;
		}
	    if(a.zip.value==""){
		   alert("邮编不能为空!");
		   a.zip.focus();
		   return false;
		}
	    if(!IsNum(a.zip.value) || a.zip.value.length>6){
		   alert("请正确输入邮编!");
		   a.zip.focus();
		   return false;
		}
		if(a.detail_address.value.length==""){
		   alert("请填写详细地址！");
		   a.detail_address.focus();
		   return false;
		}
	    if(a.shouhuo_name.value==""){
		   alert("收货人不能为空!");
		   a.shouhuo_name.focus();
		   return false;
		}
		if(a.shouhuo_mobile.value.length=="" && a.shouhuo_tel.value.length==""){
		   alert("手机和电话请任填一项");
		   a.shouhuo_mobile.focus();
		   return false;
		}
flagt=false;
  for(i=0;i<a.mainRadio.length;i++)
     a.mainRadio[i].checked?flagt=true:'';
     if(!flagt){
        alert("请选择支付方式!");
        a.mainRadio[0].focus();
        return false;
}
	    if(a.buy_num.value==""){
		   alert("购买数量不能为空!");
		   a.buy_num.focus();
		   return false;
		}
	    if(a.songhuofangshi.value==""){
		   alert("送货方式不能为空!");
		   a.songhuofangshi.focus();
		   return false;
		}
/*
if(a.mainRadio[0].checked){
flagt1=false;
  for(i=0;i<a.subRadio.length;i++)
     a.subRadio[i].checked?flagt1=true:'';
     if(!flagt1){
        alert("请选择支付方式!");
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