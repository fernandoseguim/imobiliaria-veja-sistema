






<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCodProposta, objFSO
varCodProposta = request.QueryString("varCodProposta")
   

   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT proposta.Cod_proposta,proposta.proposta_proposta,proposta.foto_proposta,proposta.foto_proposta,proposta.nome_proposta,proposta.telefone_proposta,proposta.email_proposta,proposta.data_proposta,proposta.horario_proposta,proposta.interesse_proposta,proposta.cod_imovel_proposta  FROM proposta where Cod_proposta="&varCodProposta
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQL, Conexao 
		
	
	
	dim rsVerifica
	dim strSQLVerifica
	
 Set rsVerifica = Server.CreateObject("ADODB.RecordSet")
    
	strSQLVerifica = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone='"&rs("Telefone_proposta")&"' or telefone02 like '%" & rs("Telefone_proposta") & "%' or telefone03 like '%" & rs("Telefone_proposta") & "%'"
	 
   
   
rsVerifica.CursorLocation = 3
rsVerifica.CursorType = 3

        rsVerifica.Open strSQLVerifica, Conexao 	
	




'-----------------------Sinal verde para visualização--------------------
	if Ucase(rsVerifica("atendimento")) = UCase(Session("nome_id")) then
		 Conexao.execute"update proposta set clique='"&"sim"&"' where cod_proposta="&rs("cod_proposta")
	    end if
		
%>		









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=605,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>


<title>Proposta</title>




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
		alert("Informe seu e-mail.");
		nform.txtEmail.focus();
		nform.txtEmail.select();
		return false;
	} else {
		prim = nform.txtEmail.value.indexOf("@")
		if(prim < 2) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("@",prim + 1) != -1) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(".") < 1) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(" ") != -1) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("zipmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("hotmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(".@") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("@.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(".com.br.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("/") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("[") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("]") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("(") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf(")") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		}
		if(nform.txtEmail.value.indexOf("..") > 0) {
			alert("O e-mail informado parece não estar correto.");
			nform.txtEmail.focus();
			nform.txtEmail.select();
			return false;
		
		
		
		}
		
		
	}
	//--------
	var strValidNumber="1234567890.,";
for (nCount=0; nCount < nform.txtTelefone.value.length; nCount++) 
		{
strTempChar=nform.txtTelefone.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)==-1) 
{
alert("O campo telefone deve ser numérico!")
nform.txtTelefone.focus();
nform.txtTelefone.select();
return false;
}
}
	
{
if (nform.txtTelefone.value == "") {
		alert("Digite seu telefone.");
		nform.txtTelefone.focus();
		nform.txtTelefone.select();
		return false;
}
}

{
if (nform.txtProposta.value == "") {
		alert("Digite sua proposta.");
		nform.txtProposta.focus();
		nform.txtProposta.select();
		return false;
}
}






	
}
	
	
	
	
		return true;
}












</script>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>








</head>
<!--#include file="style4_proposta.asp"-->

<body bgcolor="<%=escuro%>" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  <tr>
    <td width="590" height="170"><table width="590" height="170" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" height="170" border="0" cellpadding="0" cellspacing="0" style="border:1px solid #FFFFFF;">
              <tr>
                <td width="290" height="170" background="fundo_teste2.jpg">&nbsp;</td>
                <td width="290" height="170"><%If objFSO.FileExists(Server.MapPath(rs("Foto_proposta"))) = True Then%><a href="javascript:newWindow22('visualizar_imovel33.asp?varCod_imovel=<%=rs("cod_imovel_proposta")%>')" style="color:#FFFFFF"><img src="<%=rs("foto_proposta")%>" width="290" height="170" border="0"></img></a><% else %>
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow22('visualizar_imovel33.asp?varCod_imovel=<%=rs("cod_imovel_proposta")%>')" style="color:#FFFFFF">Foto 
                      não disponível</a></strong></font></div>
                    <% end if %></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
             
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
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atendimento</font></div></td>
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <input name="txtNome2" type="text" id="txtNome2" value="<%if not rsVerifica.eof then response.write rsVerifica("atendimento") else response.write "não informado" end if%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 16px; WIDTH: 289px; background: <%=claro%>"></td>
              </tr>
			  
			  
			  <tr> 
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                    do cliente</font></div></td>
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <input name="txtNome2" type="text" id="txtNome2" value="<%=rs("nome_proposta")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 16px; WIDTH: 289px; background: <%=medio%>"></td>
              </tr>
              <tr> 
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                    do cliente</font></div></td>
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <input name="txtTelefone2" type="text" id="txtTelefone2" value="<%=rs("telefone_proposta")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 16px; WIDTH: 289px; background: <%=claro%>"></td>
              </tr>
              <tr> 
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                    do cliente</font></div></td>
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <input name="txtEmail2" type="text" id="txtEmail2" value="<%=rs("email_proposta")%>" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 16px; WIDTH: 289px; background:<%=medio%>"></td>
              </tr>
              <tr> 
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                    do im&oacute;vel</font></div></td>
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow22('visualizar_imovel22.asp?varCod_imovel=<%=rs("cod_imovel_proposta")%>')" style="color:#FFFFFF"><%=rs("cod_imovel_proposta")%></a></font></td>
              </tr>
              <tr> 
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Hor&aacute;rio 
                    para contato</font></div></td>
                <td width="290" height="16" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <select name="txtHorario" id="select6" class="inputBox" style="HEIGHT: 16px; WIDTH: 289px; background: <%=medio%>">
                    <option value="<%=rs("horario_proposta")%>" selected><%=rs("horario_proposta")%></option>
                  </select></td>
              </tr>
              <tr> 
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                    interesse </font></div></td>
                <td width="290" height="16" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <select name="select" id="select" class="inputBox" style="HEIGHT: 16px; WIDTH: 289px; background: <%=claro%>">
                    <option value="<%=rs("interesse_proposta")%>" selected><%=rs("interesse_proposta")%></option>
                  </select></td>
              </tr>
              <tr> 
                <td width="290" height="140"><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="290" height="16" bgcolor="<%=medio%>" style="border-bottom: 1px solid #FFFFFF;border-left: 1px solid #FFFFFF;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                          do im&oacute;vel ou proposta</font></div></td>
                    </tr>
                    <tr> 
                      <td width="290" height="124"><div align="center"></div></td>
                    </tr>
                  </table></td>
                <td width="290" height="140"><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="290" height="122" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txtProposta" rows="8" cols="32"  OnKeyPress="return limitfield(this, 1024)" style="HEIGHT: 120px; WIDTH: 289px; background: <%=medio%>;" class="inputBox"><%=rs("proposta_proposta")%></textarea></td>
                    </tr>
                    <tr> 
                      <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                          <tr> 
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
 <%
           rs.Close
           'fecha a conexão
		   
		   
		   set objfso = nothing
		   
		   rsVerifica.close
		   
		   set rsVerifica = nothing
		   
		   
		   
           Conexao.Close
           Set rs = Nothing
		   set conexao = nothing
		   
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
