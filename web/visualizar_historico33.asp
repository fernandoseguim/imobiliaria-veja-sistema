<!--#include file="dsn.asp"-->
<%


response.buffer=true


dim varCod_imovel
	varCod_imovel = request.querystring("varCod_imovel")




if varCod_imovel = "" then
varCod_imovel = "0"
end if

dim rs
Set rs = Server.CreateObject("ADODB.RecordSet")

	dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	
	dim strSQL
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.historico_atual01,imoveis.historico_atual02,imoveis.historico_atual03,imoveis.historico_atual04,imoveis.historico_atual05,imoveis.historico_atual06,imoveis.historico_quem01,imoveis.historico_quem02,imoveis.historico_quem03,imoveis.historico_quem04,imoveis.historico_quem05,imoveis.historico_quem06,imoveis.ocupacao_hist,endereco_hist,valor_hist,quartos_hist,vagas_hist,suites_hist,piscina_hist,area_total_hist,area_construida_hist,edicula_hist,imoveis.condominio_hist,imoveis.captacao_hist   FROM imoveis where cod_imovel="&varCod_imovel

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
                                <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("captacao_hist") <> "" then response.write rs("captacao_hist") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="290" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                          <table width="290" border="0" cellpadding="0" cellspacing="0" bgcolor="<%=claro%>">
                            <tr>
                                <td width="145" height="20"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação</strong></font></td>
                                <td width="145" height="20"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if rs("captacao") <> "" then response.write rs("captacao") else response.write "não informado" end if%></strong></font></div></td>
                              </tr>
                            </table></td>
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
