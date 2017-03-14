
<%

option explicit 
response.buffer=true



%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->
<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_email,varSucesso_email
varCod_email = request.QueryString("varCod_email")
varSucesso_email = request.QueryString("varSucesso_email")
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT email.cod_email,email.nome,email.email,email.assunto,email.mensagem,email.data,email.cod_imovel,email.telefone,email.atendimento,email.origem  FROM email where cod_email="&varCod_email 
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
		
	'----------------------verifica comprador--------------------------	
	
	dim rsVerifica
	dim strSQLVerifica
	
 Set rsVerifica = Server.CreateObject("ADODB.RecordSet")
    
	strSQLVerifica = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone='"&rs("Telefone")&"' or telefone02 like '" & rs("Telefone") & "' or telefone03 like '" & rs("Telefone") & "'"
	 
   
   
rsVerifica.CursorLocation = 3
rsVerifica.CursorType = 3

        rsVerifica.Open strSQLVerifica, Conexao 	
		
		
	'----------------------------verifica imóvel---------------------------
	
	
	dim rsVerifica2
	dim strSQLVerifica2
	
 Set rsVerifica2 = Server.CreateObject("ADODB.RecordSet")
    
	
	dim varCod_imovel
	
	if rs("cod_imovel") <> "" and rs("cod_imovel") <> "0" then
	
	varCod_imovel = rs("cod_imovel")
	else
	varCod_imovel = "0"
	end if
	
	
	
	strSQLVerifica2 = "SELECT imoveis.cod_imovel FROM imoveis where cod_imovel="&varCod_imovel
	 
   
   
rsVerifica2.CursorLocation = 3
rsVerifica2.CursorType = 3

        rsVerifica2.Open strSQLVerifica2, Conexao 	
		
		
	
	
	
	
	
	if not rsVerifica.eof then
	'-----------------------------------------------------------------	
		if Ucase(rsVerifica("atendimento")) = UCase(Session("nome_id"))  then
		 Conexao.execute"update email set clique='"&"sim"&"' where cod_email="&rs("cod_email")
	    end if
		
		end if
		
		
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
if (nform.txtNome.value == "") {
		alert("Digite seu nome.");
		nform.txtNome.focus();
		nform.txtNome.select();
		return false;
}
}

{
if (nform.txtEmail.value == "") {
		alert("Digite seu email.");
		nform.txtEmail.focus();
		nform.txtEmail.select();
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
		alert("Digite sua Mensagem.");
		nform.txtMensagem.focus();
		nform.txtMensagem.select();
		return false;
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


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=605,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>


</head>
<!--#include file="style2_sugestoes.asp"-->
<body bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post"  onSubmit="return isValidDigitNumber(this);" name="b2">
<table width="590" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=escuro%>">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  <tr>
      <td width="590" height="18" bgcolor="<%=escuro%>"> 
        <div align="center"></div></td>
  </tr>
  <tr>
    <td width="590" height="54"><table width="590" height="54" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="5" height="54" bgcolor="<%=escuro%>">&nbsp;</td>
          <td width="580" height="54"><table width="580" height="54" border="0" cellpadding="0" cellspacing="0">
             
			 
			   <tr> 
                <td width="290" height="30"> 
                  </td>
                <td width="290" height="30">
				<% if rs("origem") = "Busca de imóvel" then  %>
				<%if not rsVerifica2.eof then %>
                  <div align="center"><a href="javascript:newWindow22('visualizar_imovel33.asp?varCod_imovel=<%=rsVerifica2("cod_imovel")%>')" style="color:#FFFFFF"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Clique 
                    aqui e veja a ficha do imóvel procurado</strong></font></a> 
                    <% else %>
                    <%end if%>
                  </div>
				  
				  <%else%>
				  
				  <% if rs("cod_imovel") <> "0" and rs("cod_imovel") <> "" and rs("origem") = "Busca de comprador" then%>
				  <div align="center"><a href="javascript:newWindow22('visualizar_compradores33.asp?varCodCompradores=<%=rs("cod_imovel")%>')" style="color:#FFFFFF"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Clique 
                    aqui e veja a ficha do comprador procurado</strong></font></a> 
				  <%end if%>
				  
				  
				  <%end if%>
				  
				  
				  </td>
              </tr>
			 
			 
			 
			   
			   
			    <tr> 
                <td width="290" height="30"> 
                  </td>
                <td width="290" height="30"><%if not rsVerifica.eof then %>
                  <div align="center"><a href="javascript:newWindow22('visualizar_compradores33.asp?varCodCompradores=<%=rsVerifica("cod_compradores")%>')" style="color:#FFFFFF"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Clique 
                    aqui e veja a ficha do comprador</strong></font></a> 
                    <% else %>
                    <%end if%>
                  </div></td>
              </tr>
			  
			   <tr> 
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atendimento</font></div></td>
                  <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
				  <%
		dim SqlAtendimento0101
		dim SqlAtendimento0202
		
		dim rsAtendimento0101
		dim rsAtendimento0202
		
		 
		 '-----------------------abrindo o banco de dados de compradores----------------
SqlAtendimento0101 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone='"&rs("Telefone")&"' or telefone02 like '" & rs("Telefone") & "' or telefone03 like '" & rs("Telefone") & "'" 

Set rsAtendimento0101 = Server.CreateObject("ADODB.RecordSet")

	rsAtendimento0101.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsAtendimento0101.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsAtendimento0101.ActiveConnection = Conexao
	
	
	rsAtendimento0101.Open sqlAtendimento0101, Conexao
	
	'--------------------------------------------------------------
	
	
	'------------------------Abrindo o banco de dados de imóveis---------------
	
	'-----------------------abrindo o banco de dados de compradores----------------
SqlAtendimento0202 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento  FROM imoveis where telefone='"&rs("Telefone")&"' or telefone02 like '" & rs("Telefone") & "' or telefone03 like '" & rs("Telefone") & "' order by cod_imovel DESC"  

Set rsAtendimento0202 = Server.CreateObject("ADODB.RecordSet")

	rsAtendimento0202.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsAtendimento0202.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsAtendimento0202.ActiveConnection = Conexao
	
	
	rsAtendimento0202.Open sqlAtendimento0202, Conexao
	
	
	
	
	
	'----------------------------------------------------------------------------
	
	

if not rsAtendimento0101.eof then

response.write rsAtendimento0101("atendimento")

else

if not rsAtendimento0202.eof then

response.write rsAtendimento0202("captacao")

else

response.write "Não disponível"
end if

end if




rsAtendimento0101.close

set rsAtendimento0101 = nothing
		 
		 
		 
		 %> </font>
                  </td>
              </tr>
			    
				<tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">


<%if rs("origem") = "Busca de comprador" then  %>
	Código do comprador procurado
	<%else%>
	Código do imóvel procurado
	<%end if%>
	
	
	<%if rs("origem") = "Busca de imóvel" then  %>
	Código do imóvel procurado
	<%end if%>				
					
					
					</font></div></td>
                <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtNome" type="text" id="txtNome" value="<%if rs("cod_imovel") <> "" then response.write rs("cod_imovel") else response.write "0" end if%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
				
				
				
				
				<tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome</font></div></td>
                <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txtNome" type="text" id="txtNome" value="<%=rs("nome")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
              <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtEmail" type="text" id="txtTelefone" value="<%if rs("telefone") <> "" then response.write rs("telefone") else response.write "não informado" end if%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
			  
			   <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txtEmail" type="text" id="txtTelefone" value="<%=rs("email")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>"></td>
              </tr>
			  
			  
              <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto</font></div></td>
                <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtAssunto" type="text" id="txtEmail" value="<%=rs("assunto")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
            </table></td>
            <td width="5" height="54" bgcolor="<%=escuro%>">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="5" height="140" bgcolor="<%=escuro%>">&nbsp;</td>
          <td><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td><table width="288" height="142" border="0" cellpadding="0" cellspacing="0" align="right">
                      <tr> 
                        <td width="290" align="center" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Mensagem</font></td>
                    </tr>
                    <tr> 
                        <td width="290" height="122" bgcolor="<%=escuro%>"> 
                          <div align="center"></div></td>
                    </tr>
                  </table></td>
                <td><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="290" height="122" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txtSugestao" cols="32" rows="8" class="inputBox" id="txtMensagem" style="HEIGHT: 120px; WIDTH: 290px; background: <%=medio%>"  OnKeyPress="return limitfield(this, 500)"><%=rs("mensagem")%></textarea></td>
                    </tr>
                    <tr>
                      <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">

                          <tr>
                              <td width="145" height="18" bgcolor="<%=escuro%>">&nbsp;</td>
                              <td width="145" height="18" bgcolor="<%=escuro%>">&nbsp;</td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
            <td width="5" height="140" bgcolor="<%=escuro%>">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
<%
           rs.Close
           'fecha a conexão
		   
		   rsVerifica.close
		   
		   set rsVerifica = nothing
		   
		   
		    rsVerifica2.close
		   
		   set rsVerifica2 = nothing
		   
		   
		   
           Conexao.Close
           Set rs = Nothing
		   Set Conexao = Nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
