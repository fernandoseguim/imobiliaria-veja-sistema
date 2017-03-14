<!--#include file="dsn.asp"-->
<%


response.buffer=true


dim varCod_compradores
	varCod_compradores = request.querystring("varCod_compradores")




if varCod_compradores = "" then
varCod_compradores = "0"
end if

dim rs
Set rs = Server.CreateObject("ADODB.RecordSet")

	dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	dim strSQL
	
	strSQL = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.historico_atual01,compradores.historico_atual02,compradores.historico_atual03,compradores.historico_atual04,compradores.historico_atual05,compradores.historico_atual06,compradores.historico_quem01,compradores.historico_quem02,compradores.historico_quem03,compradores.historico_quem04,compradores.historico_quem05,compradores.historico_quem06,compradores.ocupacao_hist,compradores.endereco_hist,compradores.valor_hist,compradores.quartos_hist,compradores.vagas_hist,compradores.suites_hist,compradores.piscina_hist,compradores.area_total_hist,compradores.area_construida_hist,compradores.edicula_hist,compradores.condominio_hist  FROM compradores where cod_compradores="&varCod_compradores
   
   
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 








dim varSucesso

varSucesso = request.querystring("varSucesso")

%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>


<title>Histórico de atualizações</title>
<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (nform) 



{




{
if (nform.txtNome.value == "") {
		alert("Digite seu nome.");
		nform.txtNome.focus();
		nform.txtNome.select();
		return false;
}
}

{
if (nform.txtEmail.value == "") {
		
}
}






{
if (nform.txtAssunto.value == "") {
		alert("Digite seu assunto.");
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


//-------------------------verifica se tem aspas no campo nome------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtNome.value.length; nCount++) 
		{
strTempChar=nform.txtNome.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Nome não pode conter aspas simples!")
nform.txtNome.focus();
nform.txtNome.select();
return false;
}
}


//-----------------------------------------------------------------------------

//-------------------------verifica se tem aspas no campo email------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtEmail.value.length; nCount++) 
		{
strTempChar=nform.txtEmail.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Email não pode conter aspas simples!")
nform.txtEmail.focus();
nform.txtEmail.select();
return false;
}
}


//-----------------------------------------------------------------------------

//-------------------------verifica se tem aspas no campo Sugestao------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtMensagem.value.length; nCount++) 
		{
strTempChar=nform.txtMensagem.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Mensagem não pode conter aspas simples!")
nform.txtMensagem.focus();
nform.txtMensagem.select();
return false;
}
}


//-----------------------------------------------------------------------------

//-------------------------verifica se tem aspas no campo Telefone------------------------------

var strValidNumber="'";
for (nCount=0; nCount < nform.txtAssunto.value.length; nCount++) 
		{
strTempChar=nform.txtAssunto.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Assunto não pode conter aspas simples!")
nform.txtAssunto.focus();
nform.txtAssunto.select();
return false;
}
}


//-----------------------------------------------------------------------------





	
}












</script>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>

<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR:<%=medio%>;}
</STYLE>

</head>

<!--#include file="style_imoveis.asp"-->
<!--#include file="cores.asp"-->
<body onload=b2.txtNome.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" action="incluir_email.asp" onSubmit="return isValidDigitNumber(this);" name="b2">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
    </tr>
    <tr>
      <td width="590" height="126"><table width="590" height="126" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="126">&nbsp;</td>
            <td width="580" height="126"><table width="580" height="126" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="580" height="18"><div align="center"><% if varSucesso = "" then %><% else %><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=varSucesso%></font> <% end if%></div></td>
                </tr>
                <tr>
                  <td width="580" height="90"><table width="580" height="90" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="290" height="25" bgcolor="#000000" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;ltimas 
                            atualiza&ccedil;&otilde;es </strong></font></div></td>
                        <td width="290" height="25" bgcolor="#000000" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quem 
                            atualizou </strong></font></div></td>
                      </tr>
					 
					 
					  <tr> 
                        <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_atual01") <> "" then response.Write rs("historico_atual01") else response.write "não informado" end if %></font></div></td>
                        <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_quem01") <> "" then response.Write rs("historico_quem01") else response.write "não informado" end if %></font></div></td>
                      </tr>
                      <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_atual02") <> "" then response.Write rs("historico_atual02") else response.write "não informado" end if %></font></div></td>
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_quem02") <> "" then response.Write rs("historico_quem02") else response.write "não informado" end if %></font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_atual03") <> "" then response.Write rs("historico_atual03") else response.write "não informado" end if %></font></div></td>
                        <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_quem03") <> "" then response.Write rs("historico_quem03") else response.write "não informado" end if %></font></div></td>
                      </tr>
                      <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_atual04") <> "" then response.Write rs("historico_atual04") else response.write "não informado" end if %></font></div></td>
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_quem04") <> "" then response.Write rs("historico_quem04") else response.write "não informado" end if %></font></div></td>
                      </tr>
                      <tr> 
                        <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_atual05") <> "" then response.Write rs("historico_atual05") else response.write "não informado" end if %></font></div></td>
                        <td width="290" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_quem05") <> "" then response.Write rs("historico_quem05") else response.write "não informado" end if %></font></div></td>
                      </tr>
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_atual06") <> "" then response.Write rs("historico_atual06") else response.write "não informado" end if %></font></div></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs("historico_quem06") <> "" then response.Write rs("historico_quem06") else response.write "não informado" end if %></font></div></td>
                      </tr>
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="40" bgcolor="<%=escuro%>" style="border:1px solid #FFFFFF;" > 
                          <div align="center"></div></td>
                        <td width="290" height="20" bgcolor="<%=escuro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"></div></td>
                      </tr>
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" bgcolor="#000000" style="border:1px solid #FFFFFF;" > 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Situa&ccedil;&atilde;o 
                            anterior </strong></font></div></td>
                        <td width="290" height="20" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Situa&ccedil;&atilde;o 
                            atual </strong></font></div></td>
                      </tr>
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <div align="center">
                            <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                              <tr>
                                <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("quartos_hist") <> "" then response.write rs("quartos_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("quartos") <> "" then response.write rs("quartos") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("vagas_hist") <> "" then response.write rs("vagas_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("vagas") <> "" then response.write rs("vagas") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Suites</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("suites_hist") <> "" then response.write rs("suites_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Suites</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("suites") <> "" then response.write rs("suites") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Piscina</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("piscina_hist") <> "" then response.write rs("piscina_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Piscina</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("piscina") <> "" then response.write rs("piscina") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ed&iacute;cula</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("edicula_hist") <> "" then response.write rs("edicula_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ed&iacute;cula</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("edicula") <> "" then response.write rs("edicula") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
                                total </strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("area_total_hist") <> "" then response.write rs("area_total_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
                                total </strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("area_total") <> "" then response.write rs("area_total") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
                                constru&iacute;da </strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("area_construida_hist") <> "" then response.write rs("area_construida_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Aacute;rea 
                                constru&iacute;da </strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("area_construida") <> "" then response.write rs("area_construida") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ocupa&ccedil;&atilde;o</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("ocupacao_hist") <> "" then response.write rs("ocupacao_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ocupa&ccedil;&atilde;o</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("ocupacao") <> "" then response.write rs("ocupacao") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  <tr bgcolor="<%=claro%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="80"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("endereco_hist") <> "" then response.write rs("endereco_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="80"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("endereco") <> "" then response.write rs("endereco") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("valor_hist") <> "" then response.write FormatNumber(rs("valor_hist"),2) else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                         <table width="290" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("valor") <> "" then response.write FormatNumber(rs("valor"),2) else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  
					  <tr bgcolor="<%=medio%>"> 
                        <td width="290" height="20" style="border:1px solid #FFFFFF;" > 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Condom&iacute;nio</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("condominio_hist") <> "" then response.write FormatNumber(rs("condominio_hist"),2) else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                         <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                
                              <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Condom&iacute;nio</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("condominio") <> "" then response.write FormatNumber(rs("condominio"),2) else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
                      </tr>
					  
					  
					  
					  
					  
					  
					  
					  
					  
                    </table></td>
                </tr>
                <tr>
                  <td width="580" height="18"><div align="center"></div></td>
                </tr>
              </table></td>
            <td width="5" height="126">&nbsp;</td>
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
 <% response.flush%>
  <%response.clear%>

</body>
</html>
