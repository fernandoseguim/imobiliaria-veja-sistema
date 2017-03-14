




<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->


<% response.buffer=True%>



<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")


 dim Conexao9,rs9
 Set Conexao9 = Server.CreateObject("ADODB.Connection")
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	Conexao9.Open dsn
	dim strSQL9
	dim varCodPermuta
	varCodPermuta=request.QueryString("varCodPermuta")
	
	 strSQL9 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais   FROM permuta where cod_permuta="&varCodPermuta
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	  
	  rs9.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs9.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs9.ActiveConnection = Conexao9
	  
	 rs9.Open strSQL9, Conexao9
	




	
		
%>		













<html>

<title>Imprimir permuta</title>
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



	
var strValidNumber1_77="1234567890,";
for (nCount=0; nCount < doublecombo.txt_cod_imovel.value.length; nCount++) 
		{
strTempChar1_77=doublecombo.txt_cod_imovel.value.substring(nCount,nCount+1);
if (strValidNumber1_77.indexOf(strTempChar1_77,0)==-1) 
{
alert("O formulário cod imovel só pode conter números!");
doublecombo.txt_cod_imovel.focus();
doublecombo.txt_cod_imovel.select();
return false;
}
}






if (doublecombo.txt_proprietario.value == "") {
        alert("Você precisa indicar o nome do proprietário!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do proprietário!");
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
alert("O formulário Telefone só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}


if (doublecombo.txt_endereco.value == "") {
        alert("Você precisa indicar o endereço do proprietário!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }




if (doublecombo.txt_valor_vend.value == "") {
        alert("O formulário valor do seu Imóvel está vazio!");
        doublecombo.txt_valor_vend.focus();
		doublecombo.txt_valor_vend.select();
        return false;
    }
	
	
	if (doublecombo.txt_valor_comp.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.txt_valor_comp.focus();
		doublecombo.txt_valor_comp.select();
        return false;
    }


var strText2_4 = doublecombo.txt_valor_vend.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.txt_valor_vend.focus();
		
		doublecombo.txt_valor_vend.select();
		
return false;

}



var strText2_5 = doublecombo.txt_valor_comp.value;
var s_strText2_5 = strText2_5.length
if (strText2_5.substring((s_strText2_5 - 3), (s_strText2_5 - 2)) != ","){

       alert("A vírgula do formulário Valor do imóvel pretendido está fora do lugar!");
       doublecombo.txt_valor_comp.focus();
		
		doublecombo.txt_valor_comp.select();
		
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


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="#FFFFFF" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  
  <div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="visualizar_permuta22.asp?varCodPermuta=<%=varCodPermuta%>" style="color:#000000">Voltar</a></font></strong> 
  </div>
  <tr>
    <td width="590" height="18">&nbsp;</td>
  </tr>
 
  
  
  <tr>
    <td>&nbsp;<div align="center">
          <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi atualizado  com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
  
  
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#FFFFFF"> 
                      <td height="20" style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atendimento</font></div></td>
                      <td height="18" style="border:1px solid #000000;"><font color="#000000"> 
                        <% if session("permissao") = "4" then%>
                        <input name="txt_atendimento" type="text" class="inputBox" id="txt_atendimento" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left">
                        <%else%>
                        <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs9("atendimento")%></font> 
                        <input name="txt_atendimento" type="hidden" class="inputBox" id="txt_atendimento" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("atendimento")%>" size="38" maxlength="50" align="left">
                        <%end if%>
                        </font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="18" style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                          de inclus&atilde;o</font></div></td>
                      <td height="18" style="border:1px solid #000000;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("data")%>" size="38" maxlength="50" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="18" style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                          da &uacute;ltima atualiza&ccedil;&atilde;o</font></div></td>
                      <td height="18" style="border:1px solid #000000;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("data_atualizacao")%>" size="38" maxlength="50" align="left">
                        </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="20" style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                          da permuta</font></div></td>
                      <td height="20" style="border:1px solid #000000;"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><%=rs9("cod_permuta")%></font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                          do im&oacute;vel do propriet&aacute;rio</font></div></td>
                      <td style="border:1px solid #000000;"><input name="txt_cod_imovel" type="text" class="inputBox" id="txt_cod_imovel" style="HEIGHT: 18px; WIDTH: 290px;" value="<% if rs9("cod_imovel") = "não informado" or rs9("cod_imovel") = "" then response.write "00" else response.write rs9("cod_imovel") end if%>" size="38" maxlength="20" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link 
                          de visualiza&ccedil;&atilde;o do im&oacute;vel do propriet&aacute;rio</font></div></td>
                      <td style="border:1px solid #000000;"><input name="txt_link" type="text" class="inputBox" id="txt_link" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("link_imovel")%>" size="38" maxlength="50" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                          do propriet&aacute;rio</font></div></td>
                      <td style="border:1px solid #000000;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("nome")%>" size="38" maxlength="35" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="20" style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                          do propriet&aacute;rio</font></div></td>
                      <td style="border:1px solid #000000;">      
                          <input name="txt_telefone" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("telefone")%>" size="38" maxlength="20" align="left">
</td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="20" style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">email 
                          do propriet&aacute;rio</font></div></td>
                      <td style="border:1px solid #000000;"> 
                        <div align="left"> <font color="#000000"> 
                          <input name="txt_email" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_email" style="HEIGHT: 18px; WIDTH: 290px ; " value="<%=rs9("email")%>" size="38" maxlength="50" align="left">
                          </font></div></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                          do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="HEIGHT: 18px; WIDTH: 290px ;" value="<%=rs9("endereco_vend")%>" size="38" maxlength="50" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                          do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone2" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone2" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("cidade_vend")%>" size="38" maxlength="20" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                          do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;"><input name="txt_telefone22" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone22" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("bairro_vend")%>" size="38" maxlength="20" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                          do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone23" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone23" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("vila_vend")%>" size="38" maxlength="20" align="left"></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                          do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone24" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone24" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("tipo_vend")%>" size="38" maxlength="20" align="left">
                        </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de dormit&oacute;rios im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone25" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone25" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("quartos_vend")%>" size="38" maxlength="20" align="left">
                        </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de vagas na garagem do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone26" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone26" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("vagas_vend")%>" size="38" maxlength="20" align="left">
                       </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                          do im&oacute;vel atual</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 150px; " value="<%=FormatNumber(rs9("valor_vend"),2)%>" size="12" maxlength="30">
                       </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <table width="290" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td width="290" height="18" style="border-bottom: 2px solid #FFFFFF;"> 
                                <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descrição 
                                  do imóvel atual</font></div></td>
                            </tr>
                            <tr> 
                              <td width="290" height="82" >&nbsp;</td>
                            </tr>
                          </table>
                          </font></div></td>
                      <td style="border:1px solid #000000;"> <font color="#000000"> 
                        <textarea name="txt_descricao_vend" class="inputBox" id="txt_descricao_vend" style="HEIGHT: 100px; WIDTH: 290px; " onKeyPress="return limitfield(this, 200)"><%=rs9("descricao_vend")%></textarea>
                        </font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td height="40"><font color="#000000">&nbsp;</font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                          pretendida </font></div></td>
                      <td style="border:1px solid #000000;"> <font color="#FFFFFF"> 
                        <font color="#000000"><font color="#000000">
                        <input name="txt_telefone27" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone27" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("cidade_comp")%>" size="38" maxlength="20" align="left">
                        </font></font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                        </font></font> </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                          pretendido </font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone28" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone28" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("bairro_comp")%>" size="38" maxlength="20" align="left">
                       </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                          pretendida</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone29" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone29" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("vila_comp")%>" size="38" maxlength="20" align="left">
                        </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                          de im&oacute;vel pretendido</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone210" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone210" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("tipo_comp")%>" size="38" maxlength="20" align="left">
                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de dormit&oacute;rios do im&oacute;vel pretendido</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone2102" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone2102" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("quartos_comp")%>" size="38" maxlength="20" align="left">
                         </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
                          de vagas do im&oacute;vel pretendido</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_telefone2103" type="<% if session("permissao") = "1" then response.write "Hidden" else response.write "text" end if %>" class="inputBox" id="txt_telefone2103" style="HEIGHT: 18px; WIDTH: 290px; " value="<%=rs9("vagas_comp")%>" size="38" maxlength="20" align="left">
                      </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td style="border:1px solid #000000;"> 
                        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                          do im&oacute;vel pretendido</font></div></td>
                      <td style="border:1px solid #000000;">
                        <input name="txt_valor_comp" type="text" class="inputBox" id="txt_valor_comp" style="HEIGHT: 18px; WIDTH: 150px; " value="<%=FormatNumber(rs9("valor_comp"),2)%>" size="12" maxlength="30">
                         </td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td width="290" height="100" style="border:1px solid #000000;" > 
                        <table width="290" border="0" cellspacing="0" cellpadding="0">
                          <tr> 
                            <td width="290" height="18" style="border-bottom: 2px solid #FFFFFF;"> 
                              <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o 
                                do im&oacute;vel pretendido</font></div></td>
                          </tr>
                          <tr> 
                            <td width="290" height="82"  >&nbsp;</td>
                          </tr>
                        </table></td>
                      <td width="290" height="100" style="border:1px solid #000000;" ><font color="#000000"> 
                        <textarea name="txt_descricao_comp" class="inputBox" id="txt_descricao_comp" style="HEIGHT: 100px; WIDTH: 290px; " onKeyPress="return limitfield(this, 200)"><%=rs9("descricao_comp")%></textarea>
                        </font></td>
                    </tr>
                    <tr bgcolor="#FFFFFF"> 
                      <td><font color="#000000">&nbsp;</font></td>
                      <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="145">&nbsp;</td>
                            <td width="145"><div align="center"><a href="" onclick="javascript:print();return false;" style="color:#000000"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Imprimir</strong></font></a></div></td>
                          </tr>
                        </table> </td>
                    </tr>
                  </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>

<%
           rs9.Close
           set rs9 = nothing
		   
           Conexao9.Close
           Set Conexao9 = nothing
           %>
 

<% response.flush%>
  <%response.clear%>
</body>
</html>
