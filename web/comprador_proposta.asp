






<%
Option Explicit
%>
<!--#include file="dsn.asp"-->


<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,intRecordCount,varCodImovel,objFSO
varCodImovel = request.QueryString("varCodImovel")
   
   dim varCod_comprador02
   varCod_comprador02 = request.QueryString("varCod_comprador02")
   
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	'strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou FROM imoveis  Where cod_imovel = "&varCodImovel
	 
   Conexao.Open dsn
   
'RS.CursorLocation = 3
'RS.CursorType = 3

      '  rs.Open strSQL, Conexao 
		
		
%>		









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Proposta pelo imóvel</title>




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



if (nform.txtNegociacao.value == "Qualquer Negociação") {
		alert("Escolha o tipo de negociação que você deseja fazer.");
		nform.txtNegociacao.focus();
		
		return false;
}

if (nform.txtNome.value == "não informado") {
		alert("Digite seu nome.");
		nform.txtNome.focus();
		nform.txtNome.select();
		return false;
}

}









{
if (nform.txtEmail.value == "") {
		
		
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
		alert("Descreva sua proposta.");
		nform.txtProposta.focus();
		nform.txtProposta.select();
		return false;
}
}

}






{






//------------- Verifica se é numérico---------------------



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
//-----------------------------------------------

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
<!--#include file="cores02.asp"-->
<!--#include file="style_imoveis02.asp"-->
<body onload=document.forms.b2.txtNome.focus(); bgcolor="#FFFFFF" topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">

<form method="post" action="comprador_proposta_incluir.asp?varCodImovel=<%=varCodImovel%>&varCod_comprador02=<%=varCod_comprador02%>" onSubmit="return isValidDigitNumber(this);" name="b2">

<table width="590" border="0" cellspacing="0" cellpadding="0" bgcolor="#f7ecbf">
  <tr>
      <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></td>
  </tr>
  
  <tr>
      <td width="590" height="18" bgcolor="#f7ecbf">&nbsp;</td>
  </tr>
  <tr>
    <td width="590" height="90"><table width="590" height="90" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="5" height="90" bgcolor="#f7ecbf">&nbsp;</td>
          <td width="580" height="90"><table width="580" height="90" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      nome </font></div></td>
                <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;">
<div align="center">
                    <input name="txtNome" type="text" id="txtNome" value="<%=session("nome")%>" size="38" maxlength="50" align="left" class="inputBox" style="border-color:#f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background: #f7ecbf">
                  </div></td>
              </tr>
              <tr > 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      telefone </font></div></td>
                <td width="290" bgcolor="<%=claro%>" height="18" style="border:1px solid #FFFFFF;"> 
                  <div align="center">
                    <input name="txtTelefone" type="text" id="txtTelefone" value="<%=session("telefone")%>" size="38" maxlength="12" align="left" class="inputBox" style="border-color:<%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                  </div></td>
              </tr>
              <tr > 
                  <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Seu 
                      emaill</font></div></td>
                <td width="290" bgcolor="#f7ecbf" height="18" style="border:1px solid #FFFFFF;"> 
                  <div align="center">
                    <input name="txtEmail" type="text" id="txtEmail" value="<%=session("email")%>" size="38" maxlength="50" align="left" class="inputBox" style="border-color:#f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background: #f7ecbf">
                  </div></td>
              </tr>
              <tr > 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Hor&aacute;rio 
                    para contato</font></div></td>
                <td width="290" bgcolor="<%=claro%>" height="18"> 
                  <div align="center">
                    <select name="txtHorario" id="txtHorario" class="inputBox" style="border-color:<%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 292px; background: <%=claro%>">
                      <option value="Qualquer Horário">Qualquer horário</option>
                      <option value="Manhã">Manhã</option>
                      <option value="Tarde">Tarde</option>
                      <option value="Noite">Noite</option>
                    </select>
                  </div></td>
              </tr>
              <tr bgcolor="355A63"> 
                  <td width="290" height="18" bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      de negocia&ccedil;&atilde;o que deseja fazer</font></div></td>
                <td width="290" bgcolor="#f7ecbf" height="18"> 
                  <div align="center">
                      <select name="txtNegociacao" id="txtNegociacao" class="inputBox" style="border-color:#f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 292px; background:#f7ecbf">
                        <option value="Qualquer Negociação">Qualquer Negociação</option>
                      <option value="Encontrar um inquilino">Encontrar um inquilino</option>
                      <option value="Encontrar um comprador">Encontrar um comprador</option>
                       <option value="Permutar imóvel">Permutar imóvel</option>
                    </select>
                  </div></td>
              </tr>
            </table></td>
            <td width="5" height="90" >&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="590" height="140"><table width="590" height="140" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td width="5" height="140" >&nbsp;</td>
          <td width="580" height="140"><table width="580" height="140" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="290" height="140"><table width="289" height="140" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; border-right:2px solid #FFFFFF;"> 
                          <div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif">Informe 
                            a sua d&uacute;vida</font></div></td>
                    </tr>
                    <tr> 
                        <td width="292" height="122" >&nbsp;</td>
                    </tr>
                  </table></td>
                <td width="290" height="140"><table width="290" height="140" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td width="290" height="122" style="border:1px solid #FFFFFF;"><textarea name="txtProposta" rows="8" cols="32"  OnKeyPress="return limitfield(this, 500)" style="border-color:<%=claro%>;color:#9d9249;HEIGHT: 119px; WIDTH: 289px; background:<%=claro%>" class="inputBox"></textarea></td>
                    </tr>
                    <tr>
                      <td width="290" height="18"><table width="290" height="18" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="145" height="18"><input name="image" type="image" src="bt_enviar001.jpg" width="145" height="18"></td>
                              <td width="145" height="18"><a href="javascript:document.forms.b2.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
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
          ' rs.Close
           'fecha a conexão
           Conexao.Close
           'Set rs = Nothing
		   
		  ' set objfso = nothing
		   
		   set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>



</body>
</html>

