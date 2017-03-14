<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%


response.buffer=true
dim varCodPermuta

varCodPermuta = request.querystring("varCodPermuta")


 dim varSucesso_email
 dim varExistente
   
   dim rs
   Set rs = Server.CreateObject("ADODB.RecordSet")
    
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	
	dim strSQL
	strSQL = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais   FROM permuta where cod_permuta="&varCodPermuta
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
	
	
	
	dim rs444Indicacao,strSQL444Indicacao
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	'--------------------------Vamos pegar as informações do permutante
	
	

'------------------------Sua Cidade--------------------------

stringIndex = " where cod_permuta<>"&"0"&""
 
 
 
  if   rs("cidade_vend") = "não informado" or rs("cidade_vend") = "" or rs("cidade_vend") = "cqualquer" or  rs("cidade_vend") = "qualquer um" then
	stringCidadeVend = ""
 else

stringCidadeVend = " and (Cidade_comp='"&rs("cidade_vend")&"' or Cidade_comp='"&"não informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend

 if   rs("bairro_vend") = "não informado" or rs("bairro_vend") = "" or rs("bairro_vend") = "bqualquer" or  rs("bairro_vend") = "qualquer um" then
	stringBairroVend = ""
 else
'stringBairroVend = ""
stringBairroVend = " and (Bairro_comp like'%"&rs("bairro_vend")&"%' or Bairro_comp like'%"&"não informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend

'" and (Vila_comp='"&rs("vila_vend")&"' or Vila_comp='"&"não informado"&"' or Vila_comp='"&"vlqualquer"&"' or Vila_comp='"&"qualquer um"&"' )"

 if   rs("vila_vend") = "não informado" or rs("vila_vend") = "" or rs("vila_vend") = "vlqualquer" or rs("vila_vend") = "qualquer um" then
	stringVilaVend =  ""
 else

stringVilaVend = ""

end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend
 
 
 if rs("tipo_vend") = "não informado" or rs("tipo_vend") = "" or rs("tipo_vend") = "tqualquer" or rs("tipo_vend") = "qualquer um"  then

stringTipoVend = ""

else
stringTipoVend = " and Tipo_comp like '%"&rs("tipo_vend")&"%'"
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend
 
 
 

stringQuartosVend = " and Quartos_comp <="&int(rs("quartos_vend"))&""

 


 '-----------------------Número de Vagas ???_?r??E????????????????????•??????????????????????????????????????????????????????????????†??????????????????????????????????????????????????????????????????????????????????????3?b??E???????????????????????????? do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend
 
 
 



stringVagasVend = " and vagas_comp <="&int(rs("vagas_vend"))&""

 




 
 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
dim PorcentualVend

dim vValorMenorVend
dim vValorMaiorVend

PorcentualVend = int(rs("valor_vend"))*20/100

   


   vValorMenorVend = int(rs("valor_vend")) - int(PorcentualVend)
   vValorMaiorVend = int(rs("valor_vend")) + int(PorcentualVend)

 
 
 
 
 
	 dim stringValorVend
  
	
	
	
	stringValorVend = " and Valor_comp >="&  vValorMenorVend &" and Valor_comp <="& vValorMaiorVend&""
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp
  if rs("cidade_comp")="não informado" or rs("cidade_comp")="" or rs("cidade_comp")="cqualquer" or rs("cidade_comp") = "qualquer um" then
	stringCidadeComp = ""
	else
	
	stringCidadeComp = " and Cidade_vend ='"& rs("cidade_comp") &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp

	if rs("bairro_comp") = "não informado" or  rs("bairro_comp") = "" or  rs("bairro_comp") = "bqualquer" or rs("bairro_comp") = "qualquer um" then
	
	
	
	
	
	stringBairroComp = ""
	
	
	
	
	else
	
	
	
	'stringBairroComp = " and Bairro_vend ='"& rs("bairro_comp") &"'"
	
	
	
	
 
dim Numero_Indicacoes
dim Numero_Indicacoes02




Numero_Indicacoes = 0
Numero_Indicacoes02 = 0


dim soma02
dim soma

soma = 0
soma02 = 0

dim Variavel
dim Retorno
dim contar
Variavel = rs("bairro_comp")
Retorno = Split(Variavel,", ")

contar=0

dim stringBairro3
dim stringBairro4
dim stringBairro5

for contar=0 to UBound(Retorno)

stringBairro3 = "and ( "
stringBairro4 = " Bairro_vend='"&Retorno(contar)&"'or  " &stringBairro4

stringBairro5 = " cod_permuta=0)"

stringBairroComp = stringBairro3&stringBairro4&stringBairro5

next




	
	
	
	
	
	
	end if
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 'and Vila_vend ='"& rs("vila_comp") &"'
	 dim stringVilaComp

	if rs("vila_comp") <> "não informado" and rs("vila_comp") <> "" and rs("vila_comp") <> "vlqualquer" and rs("vila_comp") <> "qualquer um" then
	stringVilaComp = ""
	else
	
	stringVilaComp = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp
  'if rs("tipo_comp")="não informado" or rs("tipo_comp")="" or rs("tipo_comp")="tqualquer" or rs("tipo_comp") = "qualquer um" then
	'stringTipoComp = ""
	'else
	
	
	'stringTipoComp = " and Tipo_vend ='"& rs("tipo_comp")&"'"
	'end if
	
	
 
      '--------------------------Tipo----------------------------

if rs("tipo_comp") <> "qualquer um" and rs("tipo_comp") <> "não informado" then




 
dim Numero_IndicacoesTipoComp
dim Numero_Indicacoes02TipoComp




Numero_IndicacoesTipoComp = 0
Numero_Indicacoes02TipoComp = 0


dim soma02TipoComp
dim somaTipoComp

somaTipoComp = 0
soma02TipoComp = 0

dim VariavelTipoComp
dim RetornoTipoComp
dim contarTipoComp
VariavelTipoComp =  rs("tipo_comp")
RetornoTipoComp = Split(rs("tipo_comp"),", ")

contarTipoComp=0

dim stringTipo3Comp
dim stringTipo4Comp
dim stringTipo5Comp

for contarTipoComp=0 to UBound(RetornoTipoComp)

stringTipo3Comp = "and ( "
stringTipo4Comp = " tipo_vend='"&RetornoTipoComp(contarTipoComp)&"'or  " &stringTipo4Comp

stringTipo5Comp = " cod_permuta=0)"


stringTipo2Comp = stringTipo3Comp&stringTipo4Comp&stringTipo5Comp







next

stringTipo3Comp = ""
stringTipo4Comp = ""
stringTipo5Comp = ""


else
stringTipo2Comp = ""
end if

	
	
	
	
	
 
 
 
 
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp
  
	
	stringQuartosComp = " and Quartos_vend >="& int(rs("quartos_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp
 
	
	stringVagasComp = " and vagas_vend >="& int(rs("vagas_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------

dim PorcentualComp

dim vValorMenorComp
dim vValorMaiorComp

PorcentualComp = int(rs("valor_comp"))*20/100

   


   vValorMenorComp = int(rs("valor_comp")) - int(PorcentualComp)
   vValorMaiorComp = int(rs("valor_comp")) + int(PorcentualComp)


	 dim stringValorComp
  
	
	
	stringValorComp = " and Valor_vend >="& vValorMenorComp &" and Valor_vend <="& vValorMaiorComp &""
	
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	'strSQL444 = "SELECT * FROM permuta"
	'&stringVilaVend
	'&stringVilaComp
	varIndicacaoCodigo=rs("cod_permuta")
	
	strSQL444 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais   FROM permuta"&stringIndex&stringCidadeVend&stringBairroVend&stringVilaVend&stringTipoVend&stringQuartosVend&stringVagasVend&stringValorVend&stringCidadeComp&stringBairroComp&stringVilaComp&stringTipo2Comp&stringQuartosComp&stringVagasComp&stringValorComp&" and cod_permuta not like "&varIndicacaoCodigo
	
	
	
	
	 rs444Indicacao.Open strSQL444, Conexao 	
%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Enviar Email</title>
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
alert("O campo Para não pode conter letras maiúsculas!")
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
<form method="post" onSubmit="return isValidDigitNumber(this);"  action="atualizar_lastemail33.asp?varCodPermuta=<%=varCodPermuta%>"  name="nform">
 

  
  
  
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
                      <input name="txtDe" type="text" id="txtDe" value="veja@imobiliariaveja.com.br" size="38" maxlength="45" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>;">
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
                      <input name="txtAssunto" type="text" id="txtAssunto" value="Indica&ccedil;&otilde;es de im&oacute;veis" size="38" maxlength="30" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
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
                        <td width="290" height="200"><textarea name="txtMensagem" cols="32" rows="30" class="inputBox" id="txtMensagem" style="HEIGHT: 200px; WIDTH: 292px; background: <%=claro%>; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 800)">Olá sr(a) <%=rs("nome")%> , me chamo <%=rs("atendimento")%> e sou seu atendente aqui na imobiliária Veja,estou te enviando este email com indicações	de imóveis para o sr(a).Clique nos links abaixo para visualizar os imóveis.Querendo	visitar algum deles entre em contato comigo pelo telefone: 4123-72-44 ou em veja@imobiliariaveja.com.br , aproveito para informar que desejando ver mais opções, o sr(a) poderá acessar sua conta cadastro gratuita pelo nosso site www.imobiliariaveja.com.br , Obrigado.
						
						 <% if not rs444Indicacao.eof then %>
                      <% While NOT rs444Indicacao.EoF %>
                    							
					
					
					 <% = "http://www.imobiliariaveja.com.br/visualizar_permuta01.asp?varCodPermuta="&rs444Indicacao("cod_permuta")&"&nome="&rs("nome")&"&telefone="&rs("telefone")&"&email="&rs("email")&"                                                     " %>
                    
					
					
					
                     
                      <% rs444Indicacao.MoveNext %>
                      <% Wend %>
                      <%else%>
                      Não há indicações
                      <%end if%>
					
					
Obrigado pela atenção,<%=rs("atendimento")%>.	
						
						
						
						
						
						
						</textarea></td>
                      </tr>
                      <tr>
                        <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="145" height="18"><input name="image" type="image" src="bt_enviar001.jpg" width="145" height="18"></td>
                              <td width="145" height="18"><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
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
           Conexao.Close
           Set rs = Nothing
		   Set rs444Indicacao = Nothing
		   
		   set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
