<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%


response.buffer=true
dim varCodCompradores

varCodCompradores = request.querystring("varCodCompradores")


 dim varSucesso_email
 dim varExistente
   
   dim rs
   Set rs = Server.CreateObject("ADODB.RecordSet")
    
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	
	dim strSQL
	strSQL = "SELECT  compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where cod_compradores="&varCodCompradores
	 
   Conexao.Open dsn
   
rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao

        rs.Open strSQL, Conexao 
	
	
	
	dim rs444Indicacao,strSQL444Indicacao
   
    Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	'--------------------------Vamos pegar as informações do comprador
	
	

'------------------------Cidade---------------------------

stringIndex2 = " where cod_imovel<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "não informado"  then
stringCidade2 = " and cidade='"&rs("cidade")&"'"
else
stringCidade2 = ""
end if

 '--------------------------Bairro----------------------------








if rs("bairro") <> "qualquer um" and rs("bairro") <> "não informado" then


 
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
Variavel = rs("bairro")
Retorno = Split(Variavel,", ")

contar=0

dim stringBairro3
dim stringBairro4
dim stringBairro5

for contar=0 to UBound(Retorno)

stringBairro3 = "and ( "
stringBairro4 = " Bairro='"&Retorno(contar)&"'or  " &stringBairro4

stringBairro5 = " cod_imovel=0)"


stringBairro2 = stringBairro3&stringBairro4&stringBairro5

next




else
stringBairro2 = ""
end if








 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if rs("tipo") <> "qualquer um" then
dim Numero_IndicacoesTipo
dim Numero_Indicacoes02Tipo




Numero_IndicacoesTipo = 0
Numero_Indicacoes02Tipo = 0


dim soma02Tipo
dim somaTipo

somaTipo = 0
soma02Tipo = 0

dim VariavelTipo
dim RetornoTipo
dim contarTipo
VariavelTipo =  rs("tipo")
RetornoTipo = Split(rs("tipo"),", ")

contarTipo=0

dim stringTipo3
dim stringTipo4
dim stringTipo5

for contarTipo=0 to UBound(RetornoTipo)

stringTipo3 = "and ( "
stringTipo4 = " tipo='"&RetornoTipo(contarTipo)&"'or  " &stringTipo4

stringTipo5 = " cod_imovel=0)"


stringTipo2 = stringTipo3&stringTipo4&stringTipo5







next

stringTipo3 = ""
stringTipo4 = ""
stringTipo5 = ""


else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negociação---------------------------
if rs("negociacao") = "Compra" then
vNegocio = "venda"
end if

if rs("negociacao") = "compra" then
vNegocio = "venda"
end if

if rs("negociacao") = "Aluguel" then
vNegocio = "aluguel"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if


if  rs("negociacao") <> "qualquer um" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> int(0) then
stringQuartos2 = " and quartos >="&rs("quartos")&""
else
stringQuartos2 = ""
end if
'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> int(0) then
stringVagas2 = " and vagas >="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------------Valor-----------------------------------


   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)






stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""








dim stringStandby

stringStandby = "  and imovel_em_negociacao <>  '"&"Vendido pela Veja"&"' and imovel_em_negociacao <>  '"&"Vendido por outros"&"' and imovel_em_negociacao <>  '"&"Suspenso"&"' and imovel_em_negociacao <>  '"&"Com proposta"&"' and (imovel_em_negociacao <>  '"&"incluido"&"' or imovel_em_negociacao IS NULL)"





'---------------------------------------------------------------------------


'se no cliente ou no servidor.


	strSQL444 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	
	
	rs444Indicacao.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Indicacao.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Indicacao.ActiveConnection = Conexao
	
	
	
	 rs444Indicacao.Open strSQL444, Conexao 	
%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Enviar email</title>
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
<form method="post" onSubmit="return isValidDigitNumber(this);"  action="atualizar_lastemail.asp?varCodCompradores=<%=varCodCompradores%>"  name="nform">
 

  
  
  
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
                      <input name="txtDe" type="text" id="txtDe" value="<%=LCase(session("nome_id"))%>@imobiliariaveja.com.br" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>;">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtPara" type="text" id="txtPara" value="<%=rs("email")%>" size="50" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" id="txtAssunto" value="Indica&ccedil;&otilde;es de im&oacute;veis" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
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
                        <td width="290" height="200"><textarea name="txtMensagem" cols="32" rows="30" class="inputBox" id="txtMensagem" style="HEIGHT: 200px; WIDTH: 292px; background: <%=claro%>; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 8000)">Olá sr(a) <%=rs("nome")%> , me chamo <%=rs("atendimento")%> e sou seu atendente aqui na imobiliária Veja,estou te enviando este email com indicações	de imóveis para o sr(a).Clique nos links abaixo para visualizar os imóveis.Querendo	visitar algum deles entre em contato comigo pelo telefone: 4123-72-44 ou em veja@imobiliariaveja.com.br , aproveito para informar que desejando ver mais opções, o sr(a) poderá acessar sua conta cadastro gratuita pelo nosso site www.imobiliariaveja.com.br , Obrigado.
						
						 <% if not rs444Indicacao.eof then %>
                      <% While NOT rs444Indicacao.EoF %>
                    							
					
					
					 <% = "http://www.imobiliariaveja.com.br/mostrar_imovel2.asp?varCodimovel="&rs444Indicacao("cod_imovel")&"&nome="&rs("nome")&"&telefone="&rs("telefone")&"&email="&rs("email")&"                                                     " %>
                    
					
					
					
                     
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
           Conexao.Close
           Set rs = Nothing
		   Set rs444Indicacao = Nothing
		   Set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
