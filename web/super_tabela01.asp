<!--#include file="cores.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#003399">
<table width="675" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
  <tr bgcolor="<%=claro%>"> 
    <td width="135" height="20"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis.asp" target="_blank">Im&oacute;veis</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores.asp" target="_blank">Compradores</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta.asp" target="_blank">Permuta</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta.asp" target="_blank">Proposta</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email.asp" target="_blank">Email</a></strong></font></div></td>
  </tr>
  <tr bgcolor="<%=claro%>"> 
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow7777('procurar_avaliacao_corretor.asp')" style="color:#FFFFFF">Avaliação </a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ligar_urgente_comprador.asp" target="_blank" style="color:#FFFFFF">Ligar 
                urgente</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imovel_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Imóveis 
                clicados</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_contas_procuradas_corretor.asp" target="_blank" style="color:#FFFFFF">Contas 
                acessadas</a></strong></font></div></td>
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_imovel.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                imóvel</a></strong></font></div></td>
  </tr>
  <tr bgcolor="<%=claro%>"> 
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_comprador.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                compradores</a></strong></font></div></td>
    <td width="135" height="20"> 
      <% if session("permissao") = "6" then%>
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo02.asp" target="_blank" style="color:#FFFFFF">Captação 
                bloco</a></strong></font></div>
				<%else%>
				
				<%end if%></td>
    <td width="135" height="20"> 
      <% if session("permissao") = "6" then%>
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo01.asp" target="_blank" style="color:#FFFFFF">Atendente 
                bloco</a></strong></font></div>
			<%else%>
			<%end if%></td>
    <td width="135" height="20"> 
      <% if session("permissao") = "6" or session("permissao") = "2" then%>
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_financiamentos.asp" target="_blank" style="color:#FFFFFF">Financiamentos</a></strong></font></div>
			<%else%>
			<%end if%></td>
    <td width="135" height="20"> 
      <% if session("permissao") = "6" then%>
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp" target="_blank" style="color:#FFFFFF">Cidade</a></strong></font></div>
			  <%else%>
              <%end if%></td>
  </tr>
  <tr bgcolor="<%=claro%>"> 
    <td width="135" height="20"> 
      <% if session("permissao") = "6" then%>
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro.asp" target="_blank" style="color:#FFFFFF">Bairro</a></strong></font></div>
			  <%else%>
              <%end if%></td>
    <td width="135" height="20"> 
      <% if session("permissao") = "6" then%>
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_vila.asp" target="_blank" style="color:#FFFFFF">Vila</a></strong></font></div>
			  <%else%>
              <%end if%></td>
    <td width="135" height="20"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_comprador_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Compradores 
                Clicados</a></strong></font></div></td>
    <td width="135" height="20" style="color:#FFFFFF"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_procurados.asp">Im&oacute;veis 
          procurados</a></strong></font></div></td>
    <td width="135" height="20" style="color:#FFFFFF"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_referencia_procurados.asp">Refer&ecirc;ncias 
          procuradas</a></strong></font></div></td>
  </tr>
  <tr bgcolor="<%=claro%>"> 
    <td width="135" height="20" style="color:#FFFFFF"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_procurados.asp">Permutantes 
          procurados</a></strong></font></div></td>
    <td width="135" height="20" style="color:#FFFFFF"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_origem.asp">Origem</a></strong></font></div></td>
      
    <td width="135"></td>
    <td width="135" height="20" style="color:#FFFFFF"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_tipo.asp">Tipos de imóveis</a></strong></font></div></td>
    <td width="135" height="20">&nbsp;</td>
    
  </tr>
</table>
</body>
</html>
