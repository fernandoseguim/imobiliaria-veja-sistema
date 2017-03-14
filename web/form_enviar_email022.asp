<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%


if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if

response.buffer=true
dim varCod_imovel

varCod_imovel = request.querystring("varCod_imovel")


dim varCodCompradores

varCodCompradores = request.querystring("varCodCompradores")




 dim varSucesso_email
 dim varExistente
   
   dim rs
   Set rs = Server.CreateObject("ADODB.RecordSet")
    
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	
	dim strSQL
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel  FROM imoveis where cod_imovel="&varCod_imovel
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
	
	
	
	dim rs444Indicacao,strSQL444Indicacao
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	'--------------------------Vamos pegar as informações do comprador
	
	

'------------------------Cidade---------------------------

stringIndex2 = " where cod_compradores<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "não informado" and rs("cidade") <> ""  then
stringCidade2 = " and (cidade='"&rs("cidade")&"' or cidade='"&"não informado"&"')"
else
stringCidade2 = ""
end if



 '--------------------------Bairro----------------------------

if rs("bairro") <> "qualquer um" and rs("bairro") <> "não informado" and rs("bairro") <> ""  then
stringBairro2 = " and (Bairro like '%"&rs("bairro")&"%' or Bairro like'%"&"não informado"&"%')"
else
stringBairro2 = ""
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

'if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" then
'stringTipo2 = " and Tipo='"&rs("Tipo")&"'"
'else
'stringTipo2 = ""
'end if


if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" and  rs("tipo") <> "" then
stringTipo2 = " and Tipo like '%"&rs("Tipo")&"%'"
else
stringTipo2 = ""
end if




 '------------------------------------------------------------- 







'-------------------Negociação---------------------------
vNegocio = "Compra"
if rs("negociacao") = "venda" then
vNegocio = "compra"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if

if  rs("negociacao") <> "qualquer um" and rs("negociacao") <> "" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> 0 and rs("quartos") <> "" then
stringQuartos2 = " and quartos<="&rs("quartos")&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> 0 and rs("vagas") <> "" then
stringVagas2 = " and vagas <="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------





'---------------------------------Valor-----------------------------------



 if rs("valor") <> "" and rs("valor") <> "0,00" and rs("valor") <> "0" then
'---------------------------------Valor-----------------------------------



 
   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)
  








stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

else

stringValor2 = ""

end if

dim stringCondominio101


Porcentual02101 = int(rs("condominio"))*10/100
   


   vCondominioMenor101 = int(rs("condominio")) - int(Porcentual02101)
   vCondominioMaior101 = int(rs("condominio")) + int(Porcentual02101)




if  int(rs("condominio")) <> 0 and rs("condominio") <> ""  then
stringCondominio101 = " and Condominio >="& int(rs("condominio")) &" "
else
stringCondominio101 = ""
end if


'---------------------------------------------------------------------------


'---------------------------------Área Total-----------------------------------



dim stringAreaTotal101


Porcentual03101 = int(rs("area_total"))*10/100
   


   vAreaTotalMenor101 = int(rs("area_total")) - int(Porcentual03101)
   vAreaTotalMaior101 = int(rs("area_total")) + int(Porcentual03101)



if  int(rs("area_total")) <> 0 and rs("area_total") <> "" then
stringAreaTotal101 = " and area_total >="& vAreaTotalMenor101 &" and area_total <="& vAreaTotalMaior101 &""
else
stringAreaTotal101 = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Suítes-----------------------------------------


dim stringSuites101
 
if  rs("suites") <> "suiqualquer" and rs("suites") <> "não" and rs("suites") <> "0" and rs("suites") <>"00" and rs("suites") <>"" then
stringSuites101 = "  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites101 = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina101
 
if  rs("piscina") <> "pisqualquer" and rs("piscina") <> "não" and rs("piscina") <> "0" and rs("piscina") <>"00" and rs("piscina") <>"" then
stringPiscina101 = "  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina101 = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria101
 
if  rs("portaria") <> "porqualquer" and rs("portaria") <> "não" and rs("portaria") <> "0" and rs("portaria") <>"00" and rs("portaria") <>"" then
stringPortaria101 = "  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria101 = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal101
 
if  rs("quintal") <> "quiqualquer" and rs("quintal") <> "não" and rs("quintal") <> "0" and rs("quintal") <>"00" and rs("quintal") <>"" then
stringQuintal101 = "  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal101 = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras101
 
if  rs("quadras") <> "quaqualquer" and rs("quadras") <> "não" and rs("quadras") <> "0" and rs("quadras") <>"00" and rs("quadras") <>"" then
stringQuadras101 = "  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras101 = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula101
 
if  rs("edicula") <> "ediqualquer" and rs("edicula") <> "não" and rs("edicula") <> "0" and rs("edicula") <>"00" and rs("edicula") <>"" then
stringEdicula101 = "  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula101 = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Ocupação--------------------------------------
dim stringOcupacao101
 
if  rs("ocupacao") <> "oqualquer" and rs("ocupacao") <> "não informado"  then
stringOcupacao101 = "  and ocupacao ='"&rs("ocupacao")&"'  and ocupacao IS NOT NULL"




else

stringOcupacao101 = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

'stringStandby = " and standby like '"&"suspenso"&"' and standby like '"&"comprador OK"&"'"

stringStandby = " and ( standby like 'comprador OK') and origem_franquia like '"&session("vOrigem_Franquia")&"' "










'---------------------------------------------------------------------------



	strSQL444 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	
	
	 rs444Indicacao.Open strSQL444, Conexao 	
%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Email</title>

<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (nform) 



{




{
if (nform.txtDe.value == "") {
		alert("Digite quem está mandando o email.");
		nform.txtDe.focus();
		nform.txtDe.select();
		return false;
}
}





{
if (nform.txtPara.value == "") {
		alert("Digite para quem deseja mandar o email.");
		nform.txtPara.focus();
		nform.txtPara.select();
		return false;
}
}







{
if (nform.txtAssunto.value == "") {
		alert("Digite o assunto.");
		nform.txtAssunto.focus();
		nform.txtAssunto.select();
		return false;
}
}

{
if (nform.txtMensagem.value == "") {
		alert("Digite sua mensagem.");
		nform.txtMensagem.focus();
		nform.txtMensagem.select();
		return false;
}
}



//-------------------------verifica se tem aspas no campo email------------------------------

var strValidNumber="ABCDEFGHIJKLMNOPQRSTUVXZWY";
for (nCount=0; nCount < nform.txtPara.value.length; nCount++) 
		{
strTempChar=nform.txtPara.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O este campo Para  não pode conter letras maiúsculas!")
nform.txtPara.focus();
nform.txtPara.select();
return false;
}
}


//-----------------------------------------------------------------------------



//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------






var elem=nform.elements;

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


<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>



</head>

<!--#include file="style2_sugestoes.asp"-->
<body onload=b2.txtNome.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" onSubmit="return isValidDigitNumber(this);"  action="atualizar_lastemail022.asp?varCod_imovel=<%=varCod_imovel%>"  name="nform">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
    </tr>
    <tr>
      <td width="590" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="590" height="54"><table width="590" height="54" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="54">&nbsp;</td>
            <td width="580" height="54"><table width="580" height="54" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" >
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">D&ecirc;:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtDe" type="text" id="txtDe" value="<%=LCase(session("nome_id"))%>@imobiliariaveja.com.br" size="38" maxlength="45" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>;">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtPara" type="text" id="txtPara" value="<%=rs("email")%>" size="38" maxlength="35" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" id="txtAssunto" value="Indicação de compradores para o seu imóvel" size="38" maxlength="30" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                    </div></td>
                </tr>
              </table></td>
            <td width="5" height="54">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="140">&nbsp;</td>
            <td width="580" height="140"><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="290" height="218"><table width="289" height="218" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Mensagem:</font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="200" bgcolor="<%=escuro%>"> 
                          <div align="center"></div></td>
                      </tr>
                    </table></td>
                  <td width="290" height="140"><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="290" height="200"><textarea name="txtMensagem" cols="32" rows="30" class="inputBox" id="txtMensagem" style="HEIGHT: 200px; WIDTH: 292px; background: <%=claro%>; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 800)">Olá sr(a) <%=rs("proprietario")%> , o sitema Veja detectou que estão cadastrados em nosso sitema interessados	em comprar ou alugar o seu imóvel. Clique nos links abaixo para conhecer os	interessados. Ligue já para 4123-72-44 e fale com o atendente do interessado para que este marque uma visita ao seu imóvel. Aproveitamos para informar que	desejando ver mais opções, o sr(a) poderá acessar a sua conta cadastro gratuita	pelo nosso site www.imobiliariaveja.com.br , Obrigado.
						
						<% if varCodCompradores <> "" then %>
						 <% = "http://www.imobiliariaveja.com.br/visualizar_comprador01.asp?varCodCompradores="&varCodCompradores&"&telefone="&rs("telefone")&"&nome="&rs("proprietario")&"&email="&rs("email")&"" %>
                    
						<%else%>
						
						 <% if not rs444Indicacao.eof then %>
                      <% While NOT rs444Indicacao.EoF %>
                     <% = "http://www.imobiliariaveja.com.br/visualizar_comprador01.asp?varCodCompradores="&rs444Indicacao("cod_compradores")&"&telefone="&rs("telefone")&"&nome="&rs("proprietario")&"&email="&rs("email")&"" %><br>
                    
                     
                      <% rs444Indicacao.MoveNext %>
                      <% Wend %>
                      <%else%>
					  
                      Não há indicações
                      <%end if%>
					  <%end if%>
					
					Se você já vendeu o seu imóvel clique no link abaixo:<br>
					
					<%="http://www.imobiliariaveja.com.br/form_enviar_email.asp?varJaVendeu="&"sim"&"&varTelefone="&rs("telefone")&"&varProprietario="&rs("proprietario")&"&varEmail="&rs("email")&""%>
					 
					
					
					
					
Obrigado pela atenção.	
						
						
						
						
						
						
						</textarea></td>
                      </tr>
                      <tr>
                        <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="145" height="18"><input name="image" type="image" src="bt_enviar0011.jpg" width="145" height="18"></td>
                              <td width="145" height="18"><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar0011.jpg" width="145" height="18" border="0"></a></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
            <td width="5" height="140">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
 <%
  rs.Close
  rs444Indicacao.Close
           'fecha a conexão
           
           Set rs = Nothing
		   Set rs444Indicacao = Nothing
		   
		   conexao.close
		   set conexao = nothing
		   
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>

