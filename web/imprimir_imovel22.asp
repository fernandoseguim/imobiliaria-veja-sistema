


<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->

<!--#include file="loggedin.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")




if varCod_imovel = "" then
varCod_imovel = "0"
end if

 Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou FROM imoveis where cod_imovel="&varCod_imovel


Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
 if not rs.eof then

%>		


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>



<html>
<title>Imprimir imóvel</title>

<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=medio%>;}
</STYLE>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>




<script>
function isValidDigitNumber (doublecombo)
{






if (doublecombo.txt_proprietario.value == "") {
        alert("Você precisa indicar o nome do comprador!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do comprador!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
    }
	
	
	
	
	
var strValidNumber1_7="1234567890,";
for (nCount=0; nCount < doublecombo.txt_telefone.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_telefone.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O telefone do comprador só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}





	var strValidNumber1_6="1234567890,.";
for (nCount=0; nCount < doublecombo.stage22.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.stage22.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.stage22.focus();
doublecombo.stage22.select();
return false;
}
}


var strValidNumber1_7="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_valor_vend.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_valor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor_vend.focus();
doublecombo.txt_valor_vend.select();
return false;
}
}




var strValidNumber1_8="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_condominio_vend.value.length; nCount++) 
		{
strTempChar1_8=doublecombo.txt_condominio_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_8.indexOf(strTempChar1_8,0)==-1) 
{
alert("O formulário Condomínio só pode conter números!");
doublecombo.txt_condominio_vend.focus();
doublecombo.txt_condominio_vend.select();
return false;
}
}





	
	if (doublecombo.stage22.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }


var strText2_4 = doublecombo.stage22.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.stage22.focus();
		
		doublecombo.stage22.select();
		
return false;

}





var elem=doublecombo.elements;

for (nCount=0; nCount < elem.length; nCount++)
  

	if(elem[nCount].type.indexOf("text")==0)	{
	var strValidNumber12_1="'";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)!=-1) {



alert("Este campo  não pode conter aspas");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}








}



</script>



</head>

<!--#include file="style_imprimir.asp"-->


<body onload=doublecombo.txt_atendimento.focus(); bgcolor="#FFFFFF" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >


<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="20"> <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="imprimir_imovel022.asp?varCod_imovel=<%=varCod_imovel%>" style="color:#000000">Ficha 
        do imóvel</a></strong></font></div></td>
  </tr>
  <tr>
  </tr>
  
  <tr>
      <td><table width="590" border="0" cellspacing="0" cellpadding="0">
            
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
             
			 <tr>
                  <td  style="border:1px solid #000000;">
<div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Código de referência do imóvel</font></div></td>
                <td  style="border:1px solid #000000;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("cod_imovel")%>" size="38" maxlength="50" align="left"></td>
              </tr> 
			 
			 
			  <tr>
                  <td  style="border:1px solid #000000;">
<div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de inclus&atilde;o</font></div></td>
                <td  style="border:1px solid #000000;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("data")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			   <tr>
                  <td  style="border:1px solid #000000;">
<div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      da &uacute;ltima atualiza&ccedil;&atilde;o</font></div></td>
                <td  style="border:1px solid #000000;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("data_atualizacao")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			 
			 
			 
			 
			  <tr>
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Capta&ccedil;&atilde;o</font></div></td>
                <td style="border:1px solid #000000;"><input name="txt_proprietario14" type="text" class="inputBox" id="txt_proprietario14" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("captacao")%>" size="38" maxlength="35" align="left"></td>
              </tr>
			 
			  <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presen&ccedil;a 
                      na Primeira P&aacute;gina</font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_proprietario15" type="text" class="inputBox" id="txt_proprietario15" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("presenca_primeira")%>" size="38" maxlength="35" align="left"></td>
                </tr>
			
			 
			 
              <tr>
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do propriet&aacute;rio</font></div></td>
                <td  style="border:1px solid #000000;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("proprietario")%>" size="38" maxlength="35" align="left"></td>
              </tr>
              <tr>
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      do propriet&aacute;rio</font></div></td>
                <td  style="border:1px solid #000000;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("telefone")%>" size="38" maxlength="20" align="left"></td>
              </tr>
              <tr>
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                      do propriet&aacute;rio</font></div></td>
                <td  style="border:1px solid #000000;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="HEIGHT: 18px; WIDTH: 290px ; " value="<%=rs("email")%>" size="38" maxlength="50" align="left"></td>
              </tr>
              
			  
			   <tr> 

                  <td width="290" height="18"  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                      do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #000000;"><input name="txt_titulo_anuncio_vend" type="text" class="inputBox" id="txt_titulo4" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("titulo_anuncio")%>" size="38" maxlength="40" align="left"></td>
              </tr>
			  <tr > 
                  <td width="290" height="18"  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto 
                    do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #000000;"><input name="txt_texto_anuncio_vend" type="text" class="inputBox" id="txt_texto_anuncio_vend" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("texto_anuncio")%>" size="38" maxlength="120" align="left"></td>
              </tr>
			  
			   
			
			 <tr> 
              <td width="290"  style="border:1px solid #000000; ">
<div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                      do im&oacute;vel</font></div></td>
                  <td width="290" style="border:1px solid #000000; "><input name="txt_endereco_vend" type="text" class="inputBox" id="txt_endereco_vend" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("endereco")%>" size="38" maxlength="35" align="left"></td>
            </tr>
			
			
			
			
			  <tr> 
                  <td  height="18" style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      do im&oacute;vel </font></div></td>
                  
                <td  height="18" style="border:1px solid #000000;"> 
                  <input name="txt_proprietario2" type="text" class="inputBox" id="txt_proprietario2" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs("cidade") <> "cqualquer" then response.write rs("cidade") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
                </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      do im&oacute;vel </font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_proprietario3" type="text" class="inputBox" id="txt_proprietario3" style="HEIGHT: 18px; WIDTH: 290px;" value="<%if rs("bairro") <> "bqualquer" then response.write rs("bairro") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
                </tr>
                 <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      do im&oacute;vel </font></div></td>
                  
                <td #000000 style="border:1px solid #000000;"><input name="txt_proprietario4" type="text" class="inputBox" id="txt_proprietario4" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs("vila") <> "vlqualquer" then response.write rs("vila") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
              </tr>
			  
			  
                <tr > 
                  <td style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel </font></div></td>
                  
                <td style="border:1px solid #000000;">
                  <input name="txt_proprietario5" type="text" class="inputBox" id="txt_proprietario5" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs("tipo") <> "tqualquer" then response.write rs("tipo") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
              </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Total do im&oacute;vel</font></div></td>
                  <td  style="border:1px solid #000000;"><font color="#FFFFFF"> 
                    <input name="txt_a_total_vend" type="text" class="inputBox" id="txt_a_total_vend" style="HEIGHT: 18px; WIDTH: 150px; " value="<%=rs("area_total")%>" size="12" maxlength="20">
                    <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font></td>
              </tr>
                <tr > 
                  <td style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      Constru&iacute;da do im&oacute;vel</font></div></td>
                  <td style="border:1px solid #000000;"> 
                    <input name="txt_a_constr_vend" type="text" class="inputBox" id="txt_a_constr_vend" style="HEIGHT: 18px; WIDTH: 150px; " value="<%=rs("area_construida")%>" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font> </td>
              </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos 
                      do im&oacute;vel </font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_proprietario6" type="text" class="inputBox" id="txt_proprietario6" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs("quartos") <> "qqualquer" then response.write rs("quartos") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
              </tr>
                <tr > 
                  <td style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros 
                      do im&oacute;vel </font></div></td>
                  
                <td style="border:1px solid #000000;"><input name="txt_proprietario7" type="text" class="inputBox" id="txt_proprietario7" style="HEIGHT: 18px; WIDTH: 290px;" value="<%=rs("banheiros")%>" size="38" maxlength="35" align="left"></td>
              </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na Garagem do im&oacute;vel</font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_proprietario8" type="text" class="inputBox" id="txt_proprietario8" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs("vagas") <> "vgqualquer" then response.write rs("vagas") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
              </tr>
                <tr > 
                  <td style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      </font></div></td>
                  
                <td style="border:1px solid #000000;"><input name="txt_proprietario9" type="text" class="inputBox" id="txt_proprietario9" style="HEIGHT: 18px; WIDTH: 290px; " value="<%if rs("negociacao") <> "nqualquer" then response.write rs("negociacao") else response.write "não informado" end if %>" size="38" maxlength="35" align="left"></td>
              </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do im&oacute;vel </font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; " value="<%=FormatNumber(rs("valor"),2)%>" size="12" maxlength="30">
                  </td>
              </tr>
			  
			   <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do condom&iacute;nio</font></div></td>
                  <td  style="border:1px solid #000000;"> 
                    <input name="txt_condominio_vend" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; " value="<%if rs("condominio") <> "" then response.write FormatNumber(rs("condominio"),2) else response.write "0,00" end if%>" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Placa</font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_proprietario10" type="text" class="inputBox" id="txt_proprietario10" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("placa")%>" size="38" maxlength="35" align="left"></td>
              </tr>
			  
			  
			    <tr> 
                  <td style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy</font></div></td>
                  
                <td style="border:1px solid #000000;"><input name="txt_proprietario11" type="text" class="inputBox" id="txt_proprietario11" style="HEIGHT: 18px; WIDTH: 290px;" value="<%=rs("standby")%>" size="38" maxlength="35" align="left"></td>
              </tr>
                <tr> 
                  <td  style="border:1px solid #000000;"> 
                    <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></div></td>
                  
                <td  style="border:1px solid #000000;"><input name="txt_proprietario12" type="text" class="inputBox" id="txt_proprietario12" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs("ocupacao")%>" size="38" maxlength="35" align="left"></td>
              </tr>
			  
			  
			  

			 
			 
			   
			  
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  
  
  <tr>
              
    <td width="570" height="80"><table width="570" height="80" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5">&nbsp;</td>
                <td width="480"> <p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Por 
                    meio desta autorizo a imobili&aacute;ria Veja Creci 11.676-J 
                    a colocar a sua placa de vende-se ou aluga-se, esclarecendo 
                    que essa n&atilde;o d&aacute; exclusividade a referida imobili&aacute;ria.</strong></font> 
                  <p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Aceita 
                    foto no site? </strong></font>
                  <p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    Sim/N&atilde;o: _________<br>
                    <br>
                    ass:______________________S.B.Campo___/___/_____</strong></font> 
                </td>
                <td width="90"><div align="center"><a href="" onclick="javascript:print();return false;" style="color:#000000"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong></a></div></td>
        </tr>
      </table> </td>
              
            <td>&nbsp; </td>
            </tr>
		
			
				
			
			
			
			
			
			
			
			
          </table></td>
  </tr>
  
  
</table>


<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui é criada uma variável "groups" que receberá os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a variável group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero até o número de elementos do array "groups" */

group2[i2]=new Array()
/* aqui é criado o array "group" que receberá valores conforme o número de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receberá valores de opções. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receberá valores de opções. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("200,00 até 500,00","0000000200 0000000500")
group2[2][4]=new Option("500,00 até 1000,00","0000000500 0000001000")
group2[2][5]=new Option("1000,00 até 2000,00","0000001000 0000002000")
group2[2][6]=new Option("Mais de 2000,00","0000002000 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Menos de 20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.000,00 até 50.000,00","0000020000 0000050000")
group2[3][4]=new Option("50.000,00 até 100.000,00","0000050000 0000100000")
group2[3][5]=new Option("100.000,00 até 200.000,00","0000100000 0000200000")
group2[3][6]=new Option("Mais de 200.000,00","0000200000 1000000000")









/* aqui temos um array bidimensional "group" que receberá valores de opções. */


var temp2=document.doublecombo.stage22
/* aqui a variável "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui é criada a função "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que dá um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */


for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que é escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" será o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a variável "location" recebe os valores de "stage2" que corresponde ao endereço de
link para o carregamento de página. */


//-->
</script>

<%else%>
imóvel não encontrado
<% end if%>


<%
           rs.Close
           'fecha a conexão
         
           Set rs = Nothing
		   
		   
		   conexao.close
		   
		   set conexao = nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>

