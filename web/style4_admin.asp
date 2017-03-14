<%

option explicit


%>





<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="JavaScript">
function verifica(DForm){
var elem=DForm.elements;
for(var x=0; x < elem.length; x++){
if(elem[x].type.indexOf("text")==0){
if(elem[x].value < 1){
alert("Por favor, preencha o campo "+elem[x].name);
elem[x].focus();
return false;
}
}
}

//------------- Verifica se é numérico---------------------



var elem=DForm.elements;





for (nCount=0; nCount < elem.length; nCount++)
  
    
  
	if(elem[nCount].type.indexOf("text")==0 && elem[nCount].name != "txt_nome" && elem[nCount].name != "txt_endereco"  && elem[nCount].name != "txt_admin_id" && elem[nCount].name != "txt_admin_pass" )	{
	var strValidNumber12_1="1234567890,";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)==-1) {



alert("o campo "+elem[nCount].name+" deve ser numérico");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}
//----------------------------------------------------------


//---------------verifica se a vírgula está no lugar-------


var elem=DForm.elements;



for(var x=0; x < elem.length; x++)
if(elem[x].type.indexOf("text")==0 && elem[x].name != "txt_nome" && elem[x].name != "txt_endereco"  && elem[x].name != "txt_admin_id" && elem[x].name != "txt_admin_pass" && elem[x].name != "txt_telefone" && elem[x].name != "txt_ano") {
if (elem[x].value.substring((elem[x].value.length - 3), (elem[x].value.length - 2)) != ","){

alert("o campo "+elem[x].name+"está com a vírgula fora do lugar");
        elem[x].focus();
		
		elem[x].select();
		
return false;
}
}
//------------------------------------------------------------


}
</script>

<script Language="JavaScript">
var trocaV = /,/g;
var trocaP = /\./g;

function verifica2()
{
var txt_1_2 = document.frm.txt_1_2.value.replace(trocaV,".");
var txt_2_2 = document.frm.txt_2_2.value.replace(trocaV,".");
var txt_3_2 = document.frm.txt_3_2.value.replace(trocaV,".");
var txt_4_2 = document.frm.txt_4_2.value.replace(trocaV,".");

var txt_5_2 = document.frm.txt_5_2.value.replace(trocaV,".");
var txt_6_2 = document.frm.txt_6_2.value.replace(trocaV,".");
var txt_7_2 = document.frm.txt_7_2.value.replace(trocaV,".");
var txt_8_2 = document.frm.txt_8_2.value.replace(trocaV,".");

var txt_9_2 = document.frm.txt_9_2.value.replace(trocaV,".");
var txt_10_2 = document.frm.txt_10_2.value.replace(trocaV,".");
var txt_11_2 = document.frm.txt_11_2.value.replace(trocaV,".");
var txt_12_2 = document.frm.txt_12_2.value.replace(trocaV,".");




//var total = eval(quant * preco);
var total = (eval(txt_1_2-0) + eval(txt_2_2-0) + eval(txt_3_2-0) + eval(txt_4_2-0) + eval(txt_5_2-0) + eval(txt_6_2-0) + eval(txt_7_2-0) + eval(txt_8_2-0) + eval(txt_9_2-0) + eval(txt_10_2-0) + eval(txt_11_2-0) + eval(txt_12_2-0) );
total=Math.round(total*100)/100; 
if(total.toString().indexOf(".")==-1) {total+=".00";}
total=total.toString().replace(trocaP,",");
parte=total.split(",");
(parte[1].length == 1)? total+="0":total;

document.frm.txt_13_2.value=total;


//------------------------Soma da segunda Coluna-------------------------------------


var txt_1_3 = document.frm.txt_1_3.value.replace(trocaV,".");
var txt_2_3 = document.frm.txt_2_3.value.replace(trocaV,".");
var txt_3_3 = document.frm.txt_3_3.value.replace(trocaV,".");
var txt_4_3 = document.frm.txt_4_3.value.replace(trocaV,".");

var txt_5_3 = document.frm.txt_5_3.value.replace(trocaV,".");
var txt_6_3 = document.frm.txt_6_3.value.replace(trocaV,".");
var txt_7_3 = document.frm.txt_7_3.value.replace(trocaV,".");
var txt_8_3 = document.frm.txt_8_3.value.replace(trocaV,".");

var txt_9_3 = document.frm.txt_9_3.value.replace(trocaV,".");
var txt_10_3 = document.frm.txt_10_3.value.replace(trocaV,".");
var txt_11_3 = document.frm.txt_11_3.value.replace(trocaV,".");
var txt_12_3 = document.frm.txt_12_3.value.replace(trocaV,".");




//var total = eval(quant * preco);
var total2 = (eval(txt_1_3-0) + eval(txt_2_3-0) + eval(txt_3_3-0) + eval(txt_4_3-0) + eval(txt_5_3-0) + eval(txt_6_3-0) + eval(txt_7_3-0) + eval(txt_8_3-0) + eval(txt_9_3-0) + eval(txt_10_3-0) + eval(txt_11_3-0) + eval(txt_12_3-0) );
total2=Math.round(total2*100)/100; 
if(total2.toString().indexOf(".")==-1) {total2+=".00";}
total2=total2.toString().replace(trocaP,",");
parte2=total2.split(",");
(parte2[1].length == 1)? total2+="0":total2;

document.frm.txt_13_3.value=total2;

//------------------------Soma da Terceira Coluna-------------------------------------


var txt_1_4 = document.frm.txt_1_4.value.replace(trocaV,".");
var txt_2_4 = document.frm.txt_2_4.value.replace(trocaV,".");
var txt_3_4 = document.frm.txt_3_4.value.replace(trocaV,".");
var txt_4_4 = document.frm.txt_4_4.value.replace(trocaV,".");

var txt_5_4 = document.frm.txt_5_4.value.replace(trocaV,".");
var txt_6_4 = document.frm.txt_6_4.value.replace(trocaV,".");
var txt_7_4 = document.frm.txt_7_4.value.replace(trocaV,".");
var txt_8_4 = document.frm.txt_8_4.value.replace(trocaV,".");

var txt_9_4 = document.frm.txt_9_4.value.replace(trocaV,".");
var txt_10_4 = document.frm.txt_10_4.value.replace(trocaV,".");
var txt_11_4 = document.frm.txt_11_4.value.replace(trocaV,".");
var txt_12_4 = document.frm.txt_12_4.value.replace(trocaV,".");




//var total = eval(quant * preco);
var total3 = (eval(txt_1_4-0) + eval(txt_2_4-0) + eval(txt_3_4-0) + eval(txt_4_4-0) + eval(txt_5_4-0) + eval(txt_6_4-0) + eval(txt_7_4-0) + eval(txt_8_4-0) + eval(txt_9_4-0) + eval(txt_10_4-0) + eval(txt_11_4-0) + eval(txt_12_4-0) );
total3=Math.round(total3*100)/100; 
if(total3.toString().indexOf(".")==-1) {total3+=".00";}
total3=total3.toString().replace(trocaP,",");
parte3=total3.split(",");
(parte3[1].length == 1)? total3+="0":total3;

document.frm.txt_13_4.value=total3;


//------------------------Soma da quarta Coluna-------------------------------------


var txt_1_5 = document.frm.txt_1_5.value.replace(trocaV,".");
var txt_2_5 = document.frm.txt_2_5.value.replace(trocaV,".");
var txt_3_5 = document.frm.txt_3_5.value.replace(trocaV,".");
var txt_4_5 = document.frm.txt_4_5.value.replace(trocaV,".");

var txt_5_5 = document.frm.txt_5_5.value.replace(trocaV,".");
var txt_6_5 = document.frm.txt_6_5.value.replace(trocaV,".");
var txt_7_5 = document.frm.txt_7_5.value.replace(trocaV,".");
var txt_8_5 = document.frm.txt_8_5.value.replace(trocaV,".");

var txt_9_5 = document.frm.txt_9_5.value.replace(trocaV,".");
var txt_10_5 = document.frm.txt_10_5.value.replace(trocaV,".");
var txt_11_5 = document.frm.txt_11_5.value.replace(trocaV,".");
var txt_12_5 = document.frm.txt_12_5.value.replace(trocaV,".");




//var total = eval(quant * preco);
var total4 = (eval(txt_1_5-0) + eval(txt_2_5-0) + eval(txt_3_5-0) + eval(txt_4_5-0) + eval(txt_5_5-0) + eval(txt_6_5-0) + eval(txt_7_5-0) + eval(txt_8_5-0) + eval(txt_9_5-0) + eval(txt_10_5-0) + eval(txt_11_5-0) + eval(txt_12_5-0) );
total4=Math.round(total4*100)/100; 
if(total4.toString().indexOf(".")==-1) {total4+=".00";}
total4=total4.toString().replace(trocaP,",");
parte4=total4.split(",");
(parte4[1].length == 1)? total4+="0":total4;

document.frm.txt_13_5.value=total4;


//------------------------Soma da quinta Coluna-------------------------------------


var txt_1_6 = document.frm.txt_1_6.value.replace(trocaV,".");
var txt_2_6 = document.frm.txt_2_6.value.replace(trocaV,".");
var txt_3_6 = document.frm.txt_3_6.value.replace(trocaV,".");
var txt_4_6 = document.frm.txt_4_6.value.replace(trocaV,".");

var txt_5_6 = document.frm.txt_5_6.value.replace(trocaV,".");
var txt_6_6 = document.frm.txt_6_6.value.replace(trocaV,".");
var txt_7_6 = document.frm.txt_7_6.value.replace(trocaV,".");
var txt_8_6 = document.frm.txt_8_6.value.replace(trocaV,".");

var txt_9_6 = document.frm.txt_9_6.value.replace(trocaV,".");
var txt_10_6 = document.frm.txt_10_6.value.replace(trocaV,".");
var txt_11_6 = document.frm.txt_11_6.value.replace(trocaV,".");
var txt_12_6 = document.frm.txt_12_6.value.replace(trocaV,".");




//var total = eval(quant * preco);
var total5 = (eval(txt_1_6-0) + eval(txt_2_6-0) + eval(txt_3_6-0) + eval(txt_4_6-0) + eval(txt_5_6-0) + eval(txt_6_6-0) + eval(txt_7_6-0) + eval(txt_8_6-0) + eval(txt_9_6-0) + eval(txt_10_6-0) + eval(txt_11_6-0) + eval(txt_12_6-0) );
total5=Math.round(total5*100)/100; 
if(total5.toString().indexOf(".")==-1) {total5+=".00";}
total5=total5.toString().replace(trocaP,",");
parte5=total5.split(",");
(parte5[1].length == 1)? total5+="0":total5;

document.frm.txt_13_6.value=total5;

//------------------------------------------primeira linha---------------------

var total6 = (eval(txt_1_2-0) - eval(txt_1_3-0) + eval(txt_1_4-0) - eval(txt_1_5-0));
total6=Math.round(total6*100)/100; 
if(total6.toString().indexOf(".")==-1) {total6+=".00";}
total6=total6.toString().replace(trocaP,",");
parte6=total6.split(",");
(parte6[1].length == 1)? total6+="0":total6;

document.frm.txt_1_6.value=total6;

//------------------------------------------segunda  linha---------------------

var total7 = (eval(txt_2_2-0) - eval(txt_2_3-0) + eval(txt_2_4-0) - eval(txt_2_5-0));
total7=Math.round(total7*100)/100; 
if(total7.toString().indexOf(".")==-1) {total7+=".00";}
total7=total7.toString().replace(trocaP,",");
parte7=total7.split(",");
(parte7[1].length == 1)? total7+="0":total7;

document.frm.txt_2_6.value=total7;

//------------------------------------------terceira  linha---------------------

var total8 = (eval(txt_3_2-0) - eval(txt_3_3-0) + eval(txt_3_4-0) - eval(txt_3_5-0));
total8=Math.round(total8*100)/100; 
if(total8.toString().indexOf(".")==-1) {total8+=".00";}
total8=total8.toString().replace(trocaP,",");
parte8=total8.split(",");
(parte8[1].length == 1)? total8+="0":total8;

document.frm.txt_3_6.value=total8;

//------------------------------------------quarta  linha---------------------

var total9 = (eval(txt_4_2-0) - eval(txt_4_3-0) + eval(txt_4_4-0) - eval(txt_4_5-0));
total9=Math.round(total9*100)/100; 
if(total9.toString().indexOf(".")==-1) {total9+=".00";}
total9=total9.toString().replace(trocaP,",");
parte9=total9.split(",");
(parte9[1].length == 1)? total9+="0":total9;

document.frm.txt_4_6.value=total9;

//------------------------------------------quinta  linha---------------------

var total10 = (eval(txt_5_2-0) - eval(txt_5_3-0) + eval(txt_5_4-0) - eval(txt_5_5-0));
total10=Math.round(total10*100)/100; 
if(total10.toString().indexOf(".")==-1) {total10+=".00";}
total10=total10.toString().replace(trocaP,",");
parte10=total10.split(",");
(parte10[1].length == 1)? total10+="0":total10;

document.frm.txt_5_6.value=total10;

//------------------------------------------sexta   linha---------------------

var total_11 = (eval(txt_6_2-0) - eval(txt_6_3-0) + eval(txt_6_4-0) - eval(txt_6_5-0));
total_11=Math.round(total_11*100)/100; 
if(total_11.toString().indexOf(".")==-1) {total_11+=".00";}
total_11=total_11.toString().replace(trocaP,",");
parte_11=total_11.split(",");
(parte_11[1].length == 1)? total_11+="0":total_11;

document.frm.txt_6_6.value=total_11;

//------------------------------------------sétima   linha---------------------

var total_12 = (eval(txt_7_2-0) - eval(txt_7_3-0) + eval(txt_7_4-0) - eval(txt_7_5-0));
total_12=Math.round(total_12*100)/100; 
if(total_12.toString().indexOf(".")==-1) {total_12+=".00";}
total_12=total_12.toString().replace(trocaP,",");
parte_12=total_12.split(",");
(parte_12[1].length == 1)? total_12+="0":total_12;

document.frm.txt_7_6.value=total_12;

//------------------------------------------Oitava   linha---------------------

var total_13 = (eval(txt_8_2-0) - eval(txt_8_3-0) + eval(txt_8_4-0) - eval(txt_8_5-0));
total_13=Math.round(total_13*100)/100; 
if(total_13.toString().indexOf(".")==-1) {total_13+=".00";}
total_13=total_13.toString().replace(trocaP,",");
parte_13=total_13.split(",");
(parte_13[1].length == 1)? total_13+="0":total_13;

document.frm.txt_8_6.value=total_13;

//------------------------------------------Nona   linha---------------------

var total_14 = (eval(txt_9_2-0) - eval(txt_9_3-0) + eval(txt_9_4-0) - eval(txt_9_5-0));
total_14=Math.round(total_14*100)/100; 
if(total_14.toString().indexOf(".")==-1) {total_14+=".00";}
total_14=total_14.toString().replace(trocaP,",");
parte_14=total_14.split(",");
(parte_14[1].length == 1)? total_14+="0":total_14;

document.frm.txt_9_6.value=total_14;

//------------------------------------------Décima   linha---------------------

var total_15 = (eval(txt_10_2-0) - eval(txt_10_3-0) + eval(txt_10_4-0) - eval(txt_10_5-0));
total_15=Math.round(total_15*100)/100; 
if(total_15.toString().indexOf(".")==-1) {total_15+=".00";}
total_15=total_15.toString().replace(trocaP,",");
parte_15=total_15.split(",");
(parte_15[1].length == 1)? total_15+="0":total_15;

document.frm.txt_10_6.value=total_15;


//------------------------------------------Décima primeira   linha---------------------

var total_16 = (eval(txt_11_2-0) - eval(txt_11_3-0) + eval(txt_11_4-0) - eval(txt_11_5-0));
total_16=Math.round(total_16*100)/100; 
if(total_16.toString().indexOf(".")==-1) {total_16+=".00";}
total_16=total_16.toString().replace(trocaP,",");
parte_16=total_16.split(",");
(parte_16[1].length == 1)? total_16+="0":total_16;

document.frm.txt_11_6.value=total_16;

//------------------------------------------Décima segunda   linha---------------------

var total_17 = (eval(txt_12_2-0) - eval(txt_12_3-0) + eval(txt_12_4-0) - eval(txt_12_5-0));
total_17=Math.round(total_17*100)/100; 
if(total_17.toString().indexOf(".")==-1) {total_17+=".00";}
total_17=total_17.toString().replace(trocaP,",");
parte_17=total_17.split(",");
(parte_17[1].length == 1)? total_17+="0":total_17;

document.frm.txt_12_6.value=total_17;








}







</SCRIPT>

<!--#include file="style.asp"-->
<body onload=frm.txt_nome.focus();  bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form name="frm"  action="style4_incluir2.asp" method="post" onSubmit="return verifica(this)">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="50"></td>
  </tr>
  <tr> 
    <td width="590" height="24" bgcolor="#87A5B0"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="200" height="24"> <div align="center"> 
                <table width="200" border="0" cellspacing="0" cellpadding="0">
                  <tr bgcolor="7B9AB9"> 
                    <td width="5" height="24">&nbsp;</td>
                    <td width="45" height="24"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome:</strong></font></td>
                    <td width="150" height="24"> 
                      <input name="txt_nome" type="text"  class="inputBox" id="txt_nome" style="HEIGHT: 18px; WIDTH: 100px"  size="7" maxlength="50"></td>
                </tr>
              </table>
            </div></td>
          <td width="220" height="24"><div align="left"> 
                <table width="220" height="24" border="0" cellpadding="0" cellspacing="0">
                  <tr bgcolor="7B9AB9"> 
                    <td width="60" height="24"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o:</strong></font></td>
                    <td width="160" height="24"> 
                      <input name="txt_endereco" type="text"  class="inputBox" id="txt_endereco" style="HEIGHT: 18px; WIDTH: 130px"  size="7" maxlength="50"></td>
                </tr>
              </table>
              <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
          <td width="170" height="24"> <div align="center"> 
                <table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
                  <tr bgcolor="7B9AB9"> 
                    <td width="60" height="24"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone:</strong></font></td>
                    <td width="110" height="24"> 
                      <input name="txt_telefone" type="text"  class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 80px" size="7" maxlength="9"></td>
                </tr>
              </table>
              <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24" bgcolor="#87A5B0"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="200" height="24"><table width="200" height="24" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor="7B9AB9"> 
                  <td width="5" height="24">&nbsp;</td>
                  <td width="45" height="24"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>ID:</strong></font></td>
                  <td width="150" height="24"> 
                    <input name="txt_admin_id" type="text"  class="inputBox" id="txt_admin_id" style="HEIGHT: 18px; WIDTH: 100px"  size="7" maxlength="7"></td>
              </tr>
            </table></td>
          <td width="220" height="24"><table width="220" height="24" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor="7B9AB9"> 
                  <td width="60" height="24"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Senha:</strong></font></td>
                  <td width="160" height="24"> 
                    <input name="txt_admin_pass" type="text"  class="inputBox" id="txt_admin_pass" style="HEIGHT: 18px; WIDTH: 130px" size="7" maxlength="7"></td>
              </tr>
            </table></td>
          <td width="170" height="24"><table width="170" height="24" border="0" cellpadding="0" cellspacing="0">
                <tr bgcolor="7B9AB9"> 
                  <td width="60" height="24"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ano:</strong></font></td>
                  <td width="110" height="24"> 
                    <input name="txt_ano" type="text"  class="inputBox" id="txt_ano" style="HEIGHT: 18px; WIDTH: 80px"  size="7" maxlength="4"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
        <tr bgcolor="#000000"> 
          <td width="110" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vencimento</strong></font></div></td>
          <td width="74" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Aluguel</strong></font></div></td>
          <td width="74" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Desconto</strong></font></div></td>
          <td width="74" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Acr&eacute;scimo</strong></font></div></td>
          <td width="74" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Taxa 
              adm.</strong></font></div></td>
          <td width="74" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Repasse</strong></font></div></td>
          <td width="110" height="24" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
              repasse</strong></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <font color="#FFFFFF"> 
                <select name="txt_1_1" id="select145" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_1" id="select" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Janeiro</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_1_2" type="text" value="0,00" size="7" maxlength="9"  class="inputBox" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_1_3" type="text" value="0,00" size="7" maxlength="9" class="inputBox" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_1_4" type="text" value="0,00" size="7" maxlength="9"  class="inputBox" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_1_5" type="text" value="0,00" size="7" maxlength="9"  class="inputBox" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_1_6" type="text" value="0,00" size="7" maxlength="9"  class="inputBox" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_1_7" id="select195" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_1" id="select196" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Janeiro</option>
                <option>Fevereiro</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_2_1" id="select197" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_2" id="select2" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Fevereiro</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_2_2" type="text"  class="inputBox" id="txt_2_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"> 
                <input name="txt_2_3" type="text"  class="inputBox" id="txt_2_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_2_4" type="text"  class="inputBox" id="txt_2_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_2_5" type="text"  class="inputBox" id="txt_2_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_2_6" type="text"  class="inputBox" id="txt_2_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_2_7" id="select199" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_2" id="select200" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Fevereiro</option>
                <option>Março</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                </font><font color="#FFFFFF"> 
                <select name="txt_3_1" id="select201" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_3" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Março</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_3_2" type="text"  class="inputBox" id="txt_3_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_3_3" type="text"  class="inputBox" id="txt_3_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_3_4" type="text"  class="inputBox" id="txt_3_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_3_5" type="text"  class="inputBox" id="txt_3_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"> 
                <input name="txt_3_6" type="text"  class="inputBox" id="txt_3_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_3_7" id="select203" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_3" id="select204" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Março</option>
                <option>Abril</option>
              </select>
              </font></font> </div></td>
        </tr>
      </table>
        <table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_4_1" id="select205" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_4" id="select4" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Abril</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_4_2" type="text"  class="inputBox" id="txt_4_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_4_3" type="text"  class="inputBox" id="txt_4_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_4_4" type="text"  class="inputBox" id="txt_4_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_4_5" type="text"  class="inputBox" id="txt_4_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_4_6" type="text"  class="inputBox" id="txt_4_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
              <font color="#FFFFFF"> 
                <select name="txt_4_7" id="select207" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_4" id="select208" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Abril</option>
                <option>Maio</option>
              </select>
              </font> </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_5_1" id="select209" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_5" id="select5" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Maio</option>
                 
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_5_2" type="text"  class="inputBox" id="txt_5_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"> 
                <input name="txt_5_3" type="text"  class="inputBox" id="txt_5_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_5_4" type="text"  class="inputBox" id="txt_5_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_5_5" type="text"  class="inputBox" id="txt_5_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_5_6" type="text"  class="inputBox" id="txt_5_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_5_7" id="select211" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_5" id="select212" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Maio</option>
                <option>Junho</option>
              </select>
              </font></font> </div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_6_1" id="select213" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_6" id="select6" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Junho</option>
                 
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_6_2" type="text"  class="inputBox" id="txt_6_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_6_3" type="text"  class="inputBox" id="txt_6_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_6_4" type="text"  class="inputBox" id="txt_6_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_6_5" type="text"  class="inputBox" id="txt_6_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"> 
                <input name="txt_6_6" type="text"  class="inputBox" id="txt_6_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
              <font color="#FFFFFF"> 
                <select name="txt_6_7" id="select215" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_6" id="select216" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Junho</option>
                <option>Julho</option>
              </select>
              </font> </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_7_1" id="select217" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_7" id="select7" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Julho</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_7_2" type="text"  class="inputBox" id="txt_7_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_7_3" type="text"  class="inputBox" id="txt_7_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_7_4" type="text"  class="inputBox" id="txt_7_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_7_5" type="text"  class="inputBox" id="txt_7_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_7_6" type="text"  class="inputBox" id="txt_7_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_7_7" id="select219" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_7" id="select220" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Julho</option>
                <option>Agosto</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_8_1" id="select221" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_8" id="select8" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Agosto</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_8_2" type="text"  class="inputBox" id="txt_8_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_8_3" type="text"  class="inputBox" id="txt_8_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_8_4" type="text"  class="inputBox" id="txt_8_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_8_5" type="text"  class="inputBox" id="txt_8_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_8_6" type="text"  class="inputBox" id="txt_8_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_8_7" id="select223" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_8" id="select224" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Agosto</option>
                <option>Setembro</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_9_1" id="select225" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_9" id="select9" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Setembro</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_9_2" type="text"  class="inputBox" id="txt_9_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" bgcolor="7B9AB9" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_9_3" type="text"  class="inputBox" id="txt_9_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
              </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_9_4" type="text"  class="inputBox" id="txt_9_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"> 
                <input name="txt_9_5" type="text"  class="inputBox" id="txt_9_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_9_6" type="text"  class="inputBox" id="txt_9_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_9_7" id="select227" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_9" id="select228" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Setembro</option>
                <option>Outubro</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_10_1" id="select229" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_10" id="select10" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Outubro</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_10_2" type="text"  class="inputBox" id="txt_10_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_10_3" type="text"  class="inputBox" id="txt_10_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_10_4" type="text"  class="inputBox" id="txt_10_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_10_5" type="text"  class="inputBox" id="txt_10_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_10_6" type="text"  class="inputBox" id="txt_10_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_10_7" id="select231" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_10" id="select232" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Outubro</option>
                <option>Novembro</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_11_1" id="select233" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_11" id="select11" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Novembro</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_11_2" type="text"  class="inputBox" id="txt_11_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()" >
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_11_3" type="text"  class="inputBox" id="txt_11_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_11_4" type="text"  class="inputBox" id="txt_11_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_11_5" type="text"  class="inputBox" id="txt_11_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_11_6" type="text"  class="inputBox" id="txt_11_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_11_7" id="select235" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_11" id="select236" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Novembro</option>
                <option>Dezembro</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="7B9AB9"> 
            <td width="110" height="24" style="border:1px solid #1F607F;" > 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_12_1" id="select237" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                  <option>02</option>
                  <option>03</option>
                  <option>04</option>
                  <option>05</option>
                  <option>06</option>
                  <option>07</option>
                  <option>08</option>
                  <option>09</option>
                  <option>10</option>
                  <option>11</option>
                  <option>12</option>
                  <option>13</option>
                  <option>14</option>
                  <option>15</option>
                  <option>16</option>
                  <option>17</option>
                  <option>18</option>
                  <option>19</option>
                  <option>20</option>
                  <option>21</option>
                  <option>22</option>
                  <option>23</option>
                  <option>24</option>
                  <option>25</option>
                  <option>26</option>
                  <option>27</option>
                  <option>28</option>
                  <option>29</option>
                  <option>30</option>
                  <option>31</option>
                </select>
                <font color="#FFFFFF">
                <select name="txt_venc_12" id="select12" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Dezembro</option>
                  
                </select>
                </font> <font color="#FFFFFF"> </font></font></div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_12_2" type="text"  class="inputBox" id="txt_12_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_12_3" type="text"  class="inputBox" id="txt_12_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_12_4" type="text"  class="inputBox" id="txt_12_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_12_5" type="text"  class="inputBox" id="txt_12_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"> 
                <input name="txt_12_6" type="text"  class="inputBox" id="txt_12_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" style="border:1px solid #1F607F;"> 
              <div align="center"><font color="#FFFFFF"> 
                <select name="txt_12_7" id="select239" class="inputBox" style="HEIGHT: 18px; WIDTH: 35px">
                  <option>01</option>
                <option>02</option>
                <option>03</option>
                <option>04</option>
                <option>05</option>
                <option>06</option>
                <option>07</option>
                <option>08</option>
                <option>09</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              <font color="#FFFFFF"> 
                <select name="txt_rep_12" id="select240" class="inputBox" style="HEIGHT: 18px; WIDTH: 60px">
                  <option>Dezembro</option>
                <option>Janeiro</option>
              </select>
              </font></font></div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
          <tr bgcolor="#87A5B0"> 
            <td width="110" height="24" bgcolor="#406496">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total</strong></font></div></td>
            <td width="74" height="24" bgcolor="7B9AB9"> 
              <div align="center"> 
                <input name="txt_13_2" type="text"  class="inputBox" id="txt_13_2" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" bgcolor="7B9AB9"> 
              <div align="center"> 
                <input name="txt_13_3" type="text"  class="inputBox" id="txt_13_3" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" bgcolor="7B9AB9"> 
              <div align="center"> 
                <input name="txt_13_4" type="text"  class="inputBox" id="txt_13_4" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" bgcolor="7B9AB9"> 
              <div align="center"> 
                <input name="txt_13_5" type="text"  class="inputBox" id="txt_13_5" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="74" height="24" bgcolor="7B9AB9"> 
              <div align="center"> 
                <input name="txt_13_6" type="text"  class="inputBox" id="txt_13_6" value="0,00" size="7" maxlength="9" onBlur="return verifica2()">
            </div></td>
            <td width="110" height="24" bgcolor="#406496">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24" bgcolor="#87A5B0"><table width="590" height="24" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="150" height="24" bgcolor="7B9AB9">&nbsp;</td>
          <td width="290" height="24"><table width="290" height="24" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="145" height="24"><input type="image" src="bt_enviar2.jpg" width="145" height="18"></td>
                <td width="145" height="24"><a href="javascript:document.forms.frm.reset()"><img src="bt_apagar2.jpg" width="145" height="18" border="0"></a></td>
              </tr>
            </table></td>
            <td width="150" height="24" bgcolor="7B9AB9">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td width="590" height="24">&nbsp;</td>
  </tr>
</table>
</form>
</body>
</html>
