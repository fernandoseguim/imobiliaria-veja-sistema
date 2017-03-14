<!--#include file="dsn.asp"-->



<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 





Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open Sql3, Conexao3





Sql33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 





Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs33.ActiveConnection = Conexao3
	
	
	rs33.Open Sql33, Conexao3








	strSQL44 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 




Set rs44 = Server.CreateObject("ADODB.RecordSet")

	rs44.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs44.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs44.ActiveConnection = Conexao3
	
	
	rs44.Open strSQL44, Conexao3




%> 






<%

'Criando conexão com o banco de dados! 
Set Conexao333 = Server.CreateObject("ADODB.Connection")
Conexao333.Open dsn

'

Sql333 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 


Set rs333 = Server.CreateObject("ADODB.RecordSet")

	rs333.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs333.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs333.ActiveConnection = Conexao3
	
	
	rs333.Open Sql333, Conexao3




%> 










<!--#include file="cores.asp"-->

<% response.buffer=True%>

<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
   Set rs = Server.CreateObject("ADODB.RecordSet")
   dim rs4,strSQL4
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3

RS.ActiveConnection = Conexao



RS4.CursorLocation = 3
RS4.CursorType = 3

RS4.ActiveConnection = Conexao





        rs.Open strSQL, Conexao 
		rs4.Open strSQL4, Conexao
		
	
	
	dim rs444Placa,strSQL444Placa
   
    Set rs444Placa = Server.CreateObject("ADODB.RecordSet")
	strSQL444Placa = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	rs444Placa.CursorLocation = 3
    rs444Placa.CursorType = 3

    rs444Placa.ActiveConnection = Conexao
	
	
	
	 rs444Placa.Open strSQL444Placa, Conexao 
	 
	 
	 
	 dim rs444Captacao,strSQL444Captacao
   
    Set rs444Captacao = Server.CreateObject("ADODB.RecordSet")
	strSQL444Captacao = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	rs444Captacao.CursorLocation = 3
    rs444Captacao.CursorType = 3

    rs444Captacao.ActiveConnection = Conexao
	
	
	
	
	
	 rs444Captacao.Open strSQL444Captacao, Conexao
	 
	 
	 dim rs444Captacao22,strSQL444Captacao22
   
    Set rs444Captacao22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Captacao22 = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	rs444Captacao22.CursorLocation = 3
    rs444Captacao22.CursorType = 3

    rs444Captacao22.ActiveConnection = Conexao
	
	
	
	 rs444Captacao22.Open strSQL444Captacao22, Conexao
	 
	
	
	dim rs444Origem,strSQL444Origem
   
    Set rs444Origem = Server.CreateObject("ADODB.RecordSet")
	strSQL444Origem = "SELECT origem.id_origem,origem.origem FROM origem  ORDER BY id_origem Desc" 
	
	
	rs444Origem.CursorLocation = 3
    rs444Origem.CursorType = 3

    rs444Origem.ActiveConnection = Conexao
	
	
	
	 rs444Origem.Open strSQL444Origem, Conexao
	
	dim rs444CaptacaoTextarea,strSQL444CaptacaoTextarea
   
    Set rs444CaptacaoTextarea = Server.CreateObject("ADODB.RecordSet")
	strSQL444CaptacaoTextarea = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	
	rs444CaptacaoTextarea.CursorLocation = 3
    rs444CaptacaoTextarea.CursorType = 3

    rs444CaptacaoTextarea.ActiveConnection = Conexao
	
	
	
	 rs444CaptacaoTextarea.Open strSQL444CaptacaoTextarea, Conexao	


     dim rs444Responsavel,strSQL444Responsavel
   
    Set rs444Responsavel = Server.CreateObject("ADODB.RecordSet")
	strSQL444Responsavel = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	 
	 
	 
	 rs444Responsavel.CursorLocation = 3
    rs444Responsavel.CursorType = 3

    rs444Responsavel.ActiveConnection = Conexao
	 
	 
	 
	 
	 rs444Responsavel.Open strSQL444Responsavel, Conexao


 dim rs444Responsavel22,strSQL444Responsavel22
   
    Set rs444Responsavel22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Responsavel22 = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id Desc" 
	
	rs444Responsavel22.CursorLocation = 3
    rs444Responsavel22.CursorType = 3

    rs444Responsavel22.ActiveConnection = Conexao
	
	
	
	
	 rs444Responsavel22.Open strSQL444Responsavel22, Conexao



'------------------------------selecionar os tipos de imóvel para o formulário-------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	rs444Tipo22.CursorLocation = 3
    rs444Tipo22.CursorType = 3

    rs444Tipo22.ActiveConnection = Conexao
	
	
	
	 rs444Tipo22.Open strSQL444Tipo22, Conexao







 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	rs444Tipo23.CursorLocation = 3
    rs444Tipo23.CursorType = 3

    rs444Tipo23.ActiveConnection = Conexao
	
	
	
	
	 rs444Tipo23.Open strSQL444Tipo23, Conexao




'-------------------------------------------------------------------------------------------------






%>		


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=345,height=180,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow33(abrejanela33) {
   openWindow33 = window.open(abrejanela33,'openWin33','width=450,height=400,resizable=yes')
   openWindow33.focus( )
   }

</SCRIPT>




<html>
<title>Formulário de inclusão de comprador</title>

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






if (doublecombo.txt_proprietario.value == "") {
        alert("Você precisa indicar o nome do comprador!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	
	
if (doublecombo.txt_data_futuro_contato_comprador.value == "0/0/2007 00:00:00") {
        alert("Você precisa indicar a data para o futuro contato de comprador!");
        doublecombo.txt_data_futuro_contato_comprador.focus();
		doublecombo.txt_data_futuro_contato_comprador.select();
        return false;
    }	
	
	
	

	
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do comprador!");
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
alert("O telefone do comprador só pode conter números!");
doublecombo.txt_telefone.focus();
doublecombo.txt_telefone.select();
return false;
}
}





	var strValidNumber1_6="1234567890,.";
for (nCount=0; nCount < doublecombo.stage22.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.stage22.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.stage22.focus();
doublecombo.stage22.select();
return false;
}
}


var strValidNumber1_7="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_valor_vend.value.length; nCount++) 
		{
strTempChar1_7=doublecombo.txt_valor_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_7.indexOf(strTempChar1_7,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor_vend.focus();
doublecombo.txt_valor_vend.select();
return false;
}
}




var strValidNumber1_8="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_condominio_vend.value.length; nCount++) 
		{
strTempChar1_8=doublecombo.txt_condominio_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_8.indexOf(strTempChar1_8,0)==-1) 
{
alert("O formulário Condomínio só pode conter números!");
doublecombo.txt_condominio_vend.focus();
doublecombo.txt_condominio_vend.select();
return false;
}
}





	
	if (doublecombo.stage22.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }
	
	
	if (doublecombo.stage22.value == "0,00") {
        alert("Você precisa determinar uma valor para o imóvel pretendido!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }

	
	
	
//------------------------------saldo devedor--------------------------







if (doublecombo.txt_ja_pago_devedor.value == "") {
        alert("O formulário valor já pago no saldo devedor está vazio!");
        doublecombo.txt_ja_pago_devedor.focus();
		doublecombo.txt_ja_pago_devedor.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_ja="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_ja_pago_devedor.value.length; nCount++) 
		{
strTempChar1_ja=doublecombo.txt_ja_pago_devedor.value.substring(nCount,nCount+1);
if (strValidNumber1_ja.indexOf(strTempChar1_ja,0)==-1) 
{
alert("O formulário valor já pago no saldo devedor só pode conter números!");
doublecombo.txt_ja_pago_devedor.focus();
doublecombo.txt_ja_pago_devedor.select();
return false;
}
}






if (doublecombo.txt_devendo_devedor.value == "") {
        alert("O formulário valor devido no saldo devedor está vazio!");
        doublecombo.txt_devendo_devedor.focus();
		doublecombo.txt_devendo_devedor.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_devendo="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_devendo_devedor.value.length; nCount++) 
		{
strTempChar1_devendo=doublecombo.txt_devendo_devedor.value.substring(nCount,nCount+1);
if (strValidNumber1_devendo.indexOf(strTempChar1_devendo,0)==-1) 
{
alert("O formulário valor devido no saldo devedor só pode conter números!");
doublecombo.txt_devendo_devedor.focus();
doublecombo.txt_devendo_devedor.select();
return false;
}
}














//------






//------------------------------Área Total--------------------------







if (doublecombo.txt_a_total_vend.value == "" && doublecombo.txt_pergunta.value == "sim") {
        alert("O formulário área total  está vazio!");
        doublecombo.txt_a_total_vend.focus();
		doublecombo.txt_a_total_vend.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_atotal="1234567890,.";
for (nCount=0; nCount < doublecombo.txt_a_total_vend.value.length; nCount++) 
		{
strTempChar1_atotal=doublecombo.txt_a_total_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_atotal.indexOf(strTempChar1_atotal,0)==-1 && doublecombo.txt_pergunta.value == "sim") 
{
alert("O formulário área total só pode conter números!");
doublecombo.txt_a_total_vend.focus();
doublecombo.txt_a_total_vend.select();
return false;
}
}






if (doublecombo.txt_a_total_vend.value == "00" && doublecombo.txt_pergunta.value == "sim") {
        alert("O formulário área total precisa de um valor!");
        doublecombo.txt_a_total_vend.focus();
		doublecombo.txt_a_total_vend.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	



//------


	
	
	
	
	
	
	
	
	


var strText2_4 = doublecombo.stage22.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.stage22.focus();
		
		doublecombo.stage22.select();
		
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

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_atendimento.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="incluir_compradores22.asp">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
  </tr>
  
  
  <tr>
    <td width="590" height="20"></td>
  </tr>
  
  <tr>
    <td width="590" height="120"><div align="center"><img src="simbol_comprador02.jpg"></img></div></td>
  </tr>
  
  
  
  <tr>
    <td width="590" height="30"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Formulário 
          de inclusão de compradores</strong></font></div></td>
  </tr>
  
 
  
  <tr>
      <td height="18">
<div align="center"> 
          <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi incluido com sucesso.</font> 
          <% end if %>
        </div></td>
  </tr>
  
   
  
  <tr><td width="590" height="18">&nbsp;</td></tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo cadastramento</font></div></td>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_responsavel_cadastramento_comprador" id="txt_responsavel_cadastramento_comprador" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="Internet" selected >Internet</option>
                      <% if not rs444Responsavel.eof then %>
                      <% While NOT rs444Responsavel.EoF %>
                      <option value="<% = rs444Responsavel("list_name") %>"> 
                      <% = rs444Responsavel("list_name") %>
                      </option>
                      <% rs444Responsavel.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Internet">Internet</option>
                      <%end if%>
                    </select></td>
              </tr>
			  
			 
			  
			  <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">data 
                      futuro contato</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data_futuro_contato_comprador" type="text" class="inputBox" id="txt_data_futuro_contato_comprador" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="0/0/2007 00:00:00" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			  
			  <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto 
                      futuro contato</font></div></td>
                  <td height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_assunto_futuro_contato_comprador" COLS=20 ROWS=10 class="inputBox" id="txt_assunto_futuro_contato_comprador" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			  
			  
			  <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Melhor 
                      hor&aacute;rio para visita</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="txt_melhor_horario_visita_comprador" size="1" class="inputBox" id="txt_melhor_horario_visita_comprador"  style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px;  color:FFFFFF; background: <%=claro%>">
                      <option  value="Manhâ">Manhã </option>
                      <option value="Tarde" >Tarde </option>
					   <option  value="Noite">Noite </option>
                      <option value="Manhã ou tarde" >Manhã ou tarde </option>
                     <option  value="Manhã ou noite">Manhã ou noite</option>
                      <option value="Tarde ou noite" >Tarde ou noite </option>
					  <option value="Qualquer horário">Qualquer horário</option>
					
					</select></td>
              </tr>
			  
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Origem 
                      do Comprador</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_origem_comprador" id="txt_origem_comprador" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     
                      <% if not rs444Origem.eof then %>
                      <% While NOT rs444Origem.EoF %>
                      <option value="<% = rs444Origem("origem") %>"> 
                      <% = rs444Origem("origem") %>
                      </option>
                      <% rs444Origem.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="não informado">não informado</option>
                      <%end if%>
                    </select>
                  </td>
              </tr>
			 
			   
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo atendimento</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_atendimento" id="txt_atendimento" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="Internet" selected >Internet</option>
                      <% if not rs444Captacao.eof then %>
                      <% While NOT rs444Captacao.EoF %>
                      <option value="<% = rs444Captacao("list_name") %>"> 
                      <% = rs444Captacao("list_name") %>
                      </option>
                      <% rs444Captacao.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Internet">Internet</option>
                      <%end if%>
                    </select> </td>
              </tr>
			 
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do interessado</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" id="txt_proprietario" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      residencial</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" id="txt_telefone" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                  </td>
              </tr>
			  
			  
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      comercial</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone02" type="text" id="txt_telefone02" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                  </td>
              </tr>
			  
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      celular</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone03" type="text" id="txt_telefone03" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                  </td>
              </tr>
			  
			  
			  
			  
			  
			  
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                      do interessado</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" id="txt_email" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=medio%>"></td>
              </tr>
               
			  
			 
              
			  
			  
              
			  
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="cqualquer" selected>Cidade</option>
					 <% if not rs3.eof then %>
				    <% While NOT Rs3.EoF %>
                    <option value="<% = Rs3("id_combo1") %>">
                    <% = Rs3("nome_combo1") %>
                    </option>
                    <% Rs3.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
                  </select>
                    <a href="javascript:newWindow3('form_incluir_cidade.asp')"></img></a></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" height="120" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" height="120" style="border:1px solid #FFFFFF;"> 
                    <select name="combo2" class="inputBox" style="HEIGHT: 120px; WIDTH: 150px; background:<%=medio%>" multiple size="8">
                    <option value="bqualquer">Bairro/Região</option>
					<% if not rs4.eof then%>
					<% While NOT Rs4.EoF %>
                    <option value="<% = Rs4("id_combo2") %>">
                    <% = Rs4("nome_combo2") %>
                    </option>
                    <% Rs4.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
					
					
                  </select>
                  </td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                     <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                    </select>
                  </td>
              </tr>
			  
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" size="1" id="txt_tipo" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                    	<% if not rs444Tipo22.eof then%>
					<% While NOT rs444Tipo22.EoF %>
                    <option value="<% = rs444Tipo22("tipo") %>">
                    <% =rs444Tipo22("tipo") %>
                    </option>
                    <% rs444Tipo22.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
                  </select>
                    </font></td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de quartos do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos" id="txt_quartos" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="não informado" selected>não informado</option>
                    <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na garagem do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="não informado" selected>não informado</option>
                    <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
              
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                    <option value="não informado" selected>não informado</option>
                    
                    <option value="vago">Vago</option>
					 <option value="ocupado por terceiros">Ocupado por terceiros</option>
                    <option value="ocupado pelo inquilino">Ocupado pelo inquilino</option>
					<option value="ocupado pelo proprietário">Ocupado por terceiros</option>
				  </select>
                  </td>
              </tr>
              
              
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      que deseja</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="example2" size="1" class="inputBox" id="example2"  style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px;  color:FFFFFF; background: <%=medio%>">
                      
                      <option  value="aluguel">Aluguel </option>
                      <option value="compra" selected>Compra </option>
                    </select> </td>
              </tr>
			   <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy 
                      do comprador</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby" id="txt_standby" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="excluido" selected>Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Faixa 
                      de pre&ccedil;o desejada</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="stage22" type="text" class="inputBox" id="stage22" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="13"> 
                  </td>
              </tr>
             
              <tr bgcolor="<%=claro%>">
                  <td width="290" height="100" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">D&ecirc; 
                      mais detalhes sobre o im&oacute;vel desejado por esse cliente 
                      e diga de que forma ele pretende pagar</font></div></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" COLS=20 ROWS=10 class="inputBox" id="txt_descricao" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			  
			   <tr bgcolor="<%=medio%>">
                  <td width="290" height="100" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Coloque 
                      aqui os dados confidenciais desse cliente</font></div></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao_confi" COLS=20 ROWS=10 class="inputBox" id="txt_descricao_confi" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			  
			  
			  <tr>
			      <td width="290" height="50"><div align="right"></div></td>
                  <td width="290" height="180"> <div align="left"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Caso 
                      esse cliente tenha um im&oacute;vel para dar como parte 
                      de pagamento na compra do im&oacute;vel que ele deseja ou 
                      aceite na venda do seu im&oacute;vel, outro im&oacute;vel 
                      como parte de pagamento responda &quot;sim&quot; abaixo:</strong></font></div></td>
			  </tr>
			   <tr> 
                  <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>"> 
                    <div align="center">&nbsp;</div></td>
                  <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<select name="txt_pergunta" id="txt_pergunta" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                
				<% if session("permissao")= "5" or session("permissao") = "6" or session("permissao") = "2" then %>
				
				  <option value="sim">Sim</option>
                  <option value="nao" selected>Não</option>
               <%else%>
			   <option value="nao" selected>Não</option>
			   <%end if%>
			   
			    </select>
                  </td>
            </tr>
			
			
			<tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Capta&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_captacao" id="txt_captacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="Internet" selected >Internet</option>
                      <% if not rs444Captacao22.eof then %>
                      <% While NOT rs444Captacao22.EoF %>
                      <option value="<% = rs444Captacao22("list_name") %>"> 
                      <% = rs444Captacao22("list_name") %>
                      </option>
                      <% rs444Captacao22.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Internet">Internet</option>
                      <%end if%>
                    </select> </td>
              </tr>
			  
			  
			   <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Origem 
                      da capta&ccedil;&atilde;o</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_origem_captacao" type="text" id="txt_origem_captacao" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
			  
			
			 <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo cadastramento</font></div></td>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_responsavel_cadastramento" id="txt_responsavel_cadastramento" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="Internet" selected >Internet</option>
                      <% if not rs444Responsavel22.eof then %>
                      <% While NOT rs444Responsavel22.EoF %>
                      <option value="<% = rs444Responsavel22("list_name") %>"> 
                      <% = rs444Responsavel22("list_name") %>
                      </option>
                      <% rs444Responsavel22.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Internet">Internet</option>
                      <%end if%>
                    </select></td>
              </tr>
			
			
			
			
			
			<tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presen&ccedil;a 
                      na Primeira P&aacute;gina</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_presenca_primeira_vend" size="1" class="inputBox" id="txt_presenca_primeira_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="incluido">Incluído</option>
                      <option value="excluido" selected>Excluído</option>
                    </select>
                  </td>
                </tr>
			 <tr bgcolor="<%=medio%>"> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                      do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_titulo_anuncio_vend" type="text" id="txt_titulo_anuncio_vend" size="38" maxlength="40" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"></td>
              </tr>
			  <tr bgcolor="<%=claro%>"> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_texto_anuncio_vend" type="text" id="txt_texto_anuncio_vend" size="38" maxlength="120" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>"></td>
              </tr>
			 
			 
             
			
			
			
			 <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                      do im&oacute;vel</font></div></td>
                  <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><input name="txt_endereco_vend" type="text" id="txt_endereco_vend" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>">
                  </td>
            </tr>
			
			 <tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Chaves 
                      do im&oacute;vel</font></div></td>
                  <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><input name="txt_chaves_do_imovel" type="text" id="txt_chaves_do_imovel" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>">
                  </td>
            </tr>
			
			
			
			
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" onChange="javascript:atualizacarros2(this.form);">
                      <option value="cqualquer" selected>Cidade</option>
					 <% if not rs33.eof then %>
				    <% While NOT Rs33.EoF %>
                    <option value="<% = Rs33("id_combo1") %>">
                    <% = Rs33("nome_combo1") %>
                    </option>
                    <% Rs33.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
                  </select>
                    <a href="javascript:newWindow3('form_incluir_cidade.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a> 
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo4" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" onChange="javascript:atualizacarros999(this.form);">
                      <option value="bqualquer" selected>Bairro/Região</option>
					<% if not rs44.eof then%>
					<% While NOT Rs44.EoF %>
                    <option value="<% = Rs44("id_combo2") %>">
                    <% = Rs44("nome_combo2") %>
                    </option>
                    <% Rs44.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
					
					
                  </select>
                    <a href="javascript:newWindow3('form_incluir_bairro.asp')"><img src="bt_mais02.jpg" width="18" height="18" border="0"></a> 
                  </td>
                </tr>
                 <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                    </select>
                    <a href="javascript:newWindow3('form_incluir_vila.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a> 
                  </td>
              </tr>
			  
			  
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo</font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                     	<% if not rs444Tipo23.eof then%>
					<% While NOT rs444Tipo23.EoF %>
                    <option value="<% = rs444Tipo23("tipo") %>">
                    <% =rs444Tipo23("tipo") %>
                    </option>
                    <% rs444Tipo23.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
                  </select>
                    </font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      do terreno (para casas)</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <input name="txt_a_total_vend" type="text" class="inputBox" id="txt_a_total_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      constru&iacute;da (para casas) e &uacute;til (para aptos)</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_a_constr_vend" type="text" class="inputBox" id="txt_a_constr_vend" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font> </td>
              </tr>
			  
			  
			  
			 <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      de frente</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_de_frente" type="text" class="inputBox" id="txt_metros_de_frente" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			   <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      de fundo</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_de_fundo" type="text" class="inputBox" id="txt_metros_de_fundo" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			   <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      lateral esquerda</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_lateral_esquerda" type="text" class="inputBox" id="txt_metros_lateral_esquerda" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			   <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      lateral direita</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_lateral_direita" type="text" class="inputBox" id="txt_metros_lateral_direita" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>" value="00" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos_vend" id="txt_quartos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                      <option value="não informado" selected>não informado</option>
                    <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Su&iacute;tes</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_suites" id="txt_suites" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="não informado" selected>não informado</option>
                   <option value="00">00</option>
				    <option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
			  
			  
			  
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_banheiros_vend" id="txt_banheiros_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="não informado" selected>não informado</option>
                    <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vagas 
                      na Garagem</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas_vend" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="não informado" selected>não informado</option>
                    <option value="00">00</option>
					<option value="01">01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                  </td>
              </tr>
                <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_negociacao_vend" id="txt_negociacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="aluguel" selected>Aluguel</option>
                    <option value="venda">Venda</option>
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="13">
                  </td>
              </tr>
			   <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do condom&iacute;nio</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_condominio_vend" type="text" class="inputBox" id="txt_condominio_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="13">
                  </td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
                      devedor </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_saldo_devedor" id="txt_saldo_devedor" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="sem saldo devedor" selected>Sem saldo devedor</option>
					  <option value="com saldo devedor" >Com saldo devedor</option>
                    
                    
                  </select>
                  </td>
              </tr>
			  
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      j&aacute; pago no saldo devedor</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_ja_pago_devedor" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      devido no saldo devedor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_devendo_devedor" type="text" class="inputBox" id="txt_valor2" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="0,00" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Placa</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_placa_vend" id="txt_placa_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="Sem Placa" selected >Sem Placa</option>
                      <% if not rs444Placa.eof then %>
                      <% While NOT rs444Placa.EoF %>
                      <option value="<% = rs444Placa("list_name") %>"> 
                      <% = rs444Placa("list_name") %>
                      </option>
                      <% rs444Placa.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="Sem Placa">Sem Placa</option>
                      <%end if%>
                    </select>
                  </td>
              </tr>
			  
			  
			    <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy 
                      </font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby_vend" id="txt_standby_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="excluido" selected>Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr>
			  
			  
			   <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Imóvel 
                      vendido/suspenso/com proposta</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_imovel_em_negociacao" id="teste" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="não informado" selected>Não informado</option>
					  <option value="Vendido pela Veja">Vendido pela Veja</option>
                    <option value="Vendido por outros">Vendido por outros</option>
                    <option value="Suspenso">Suspenso</option>
					<option value="Com proposta">Com proposta</option>
                  </select>
                  </td>
              </tr>
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao_vend" id="txt_ocupacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                      <option value="não informado" selected>não informado</option>
                    <option value="vago">vago</option>
                     <option value="ocupado por terceiros">Ocupado por terceiros</option>
                    <option value="ocupado pelo inquilino">Ocupado pelo inquilino</option>
					<option value="ocupado pelo proprietário">Ocupado pelo proprietário</option>
                    
                  </select>
                  </td>
              </tr>
			  <tr>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Qualidade 
                      do neg&oacute;cio</font></div></td>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_qualidade_vend" id="txt_qualidade_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="bom negócio" >Bom Negócio</option>
                    <option value="negócio comum" selected>Negócio Comum</option>
                    
                  </select></td>
              </tr>
			  
			  
			  
			  <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Melhor 
                      hor&aacute;rio para visita</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="txt_melhor_horario_visita" size="1" class="inputBox" id="example2"  style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px;  color:FFFFFF; background: <%=claro%>">
                      <option  value="Manhâ">Manhã </option>
                      <option value="Tarde" >Tarde </option>
					   <option  value="Noite">Noite </option>
                      <option value="Manhã ou tarde" >Manhã ou tarde </option>
                     <option  value="Manhã ou noite">Manhã ou noite</option>
                      <option value="Tarde ou noite" >Tarde ou noite </option>
					  <option value="Qualquer horário">Qualquer horário</option>
					
					</select></td>
              </tr>
			  
			  
			   <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">data 
                      futuro contato</font></div></td>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_futuro_contato" type="text" class="inputBox" id="txt_data_futuro_contato" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="0/0/2007 00:00:00" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			  
			  <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto 
                      futuro contato</font></div></td>
                  <td height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><textarea name="txt_assunto_futuro_contato" COLS=20 ROWS=10 class="inputBox" id="txt_assunto_futuro_contato" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			  
			  
			  
			  
              <tr>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">D&ecirc; 
                      mais detalhes sobre o im&oacute;vel e se poss&iacute;vel 
                      o motivo da venda</font></div></td>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="obs_imovel_vend" class="inputBox" id="obs_imovel_vend" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
              <tr bgcolor="<%=claro%>">
                  <td width="290" height="100" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Coloque 
                      aqui os dados confidenciais do propriet&aacute;rio</font></div></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="obs_proprietario_vend" class="inputBox" id="obs_proprietario_vend" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			  
			  
			  
			  
			  
			  
			  
			  
			  
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><input name="image" type="image"  src="bt_enviar001.jpg" width="145" height="18" border="0"></td>
                        <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar001.jpg" width="145" height="18" border="0"></a></td>
                      </tr>
                    </table></td>
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</form>



<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui é criada uma variável "groups" que receberá os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a variável group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero até o número de elementos do array "groups" */

group2[i2]=new Array()
/* aqui é criado o array "group" que receberá valores conforme o número de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receberá valores de opções. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receberá valores de opções. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("200,00 até 500,00","0000000200 0000000500")
group2[2][4]=new Option("500,00 até 1000,00","0000000500 0000001000")
group2[2][5]=new Option("1000,00 até 2000,00","0000001000 0000002000")
group2[2][6]=new Option("Mais de 2000,00","0000002000 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Menos de 20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.000,00 até 50.000,00","0000020000 0000050000")
group2[3][4]=new Option("50.000,00 até 100.000,00","0000050000 0000100000")
group2[3][5]=new Option("100.000,00 até 200.000,00","0000100000 0000200000")
group2[3][6]=new Option("Mais de 200.000,00","0000200000 1000000000")









/* aqui temos um array bidimensional "group" que receberá valores de opções. */


var temp2=document.doublecombo.stage22
/* aqui a variável "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui é criada a função "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que dá um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */


for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que é escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" será o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a variável "location" recebe os valores de "stage2" que corresponde ao endereço de
link para o carregamento de página. */


//-->
</script>



<%
'-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------

'-----------------------------
            rs4.close        
		   
           Set rs4 = Nothing
		   
'---------------------------------




'-----------------------------
           rs33.Close           
		   
           Set rs33 = Nothing
		   
'---------------------------------


'-----------------------------
           rs44.Close           
		   
           Set rs44 = Nothing
		   
'---------------------------------




'-----------------------------
           rs333.Close           
		   
           Set rs333 = Nothing
		   
'---------------------------------

































'-----------------------------
           rs444Placa.Close           
		   
           Set rs444Placa = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Captacao.Close           
		   
           Set rs444Captacao = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Captacao22.Close           
		   
           Set rs444Captacao22 = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Origem.Close           
		   
           Set rs444Origem = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Tipo22.Close           
		   
           Set rs444Tipo22 = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Tipo23.Close           
		   
           Set rs444Tipo23 = Nothing
		   
'---------------------------------


'-----------------------------
                    
		   
           Set rscaptacaoTextarea = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Responsavel22.Close           
		   
           Set rs444Responsavel22 = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Responsavel.Close           
		   
           Set rs444Responsavel = Nothing
		   
'---------------------------------





%>







<%
           rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
           %>
 

<% response.flush%>
  <%response.clear%>
  
 <%
Function EscreveFuncaoJavaScript ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo1.options[doublecombo.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas3 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas3.ActiveConnection = Conexao3
	
	
	rsMarcas3.Open SqlMarcas3, Conexao3







While NOT (rsMarcas3.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas3("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros3 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas3("id_combo1")&" order by nome_combo2"




Set rsCarros3 = Server.CreateObject("ADODB.RecordSet")

	rsCarros3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsCarros3.ActiveConnection = Conexao3
	
	
	rsCarros3.Open SqlCarros3, Conexao3







'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1  
While NOT (rsCarros3.EoF)

Response.Write "doublecombo.combo2.options[" & i & "] = new Option('" & rsCarros3("nome_combo2") & "','" & rsCarros3("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros3.MoveNext
Wend
Response.Write "doublecombo.combo2.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas3.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 




rsMarcas3.Close           
		   
           Set rsMarcas3 = Nothing
             
			rsCarros3.Close           
		   
           Set rsCarros3 = Nothing 






End Function




%>  
 
 <%
Function EscreveFuncaoJavaScript2 ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo3.options[doublecombo.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas33 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas33.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsMarcas33.ActiveConnection = Conexao3
	
	
	rsMarcas33.Open SqlMarcas33, Conexao3





While NOT (rsMarcas33.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas33("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros33 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas33("id_combo1")&" order by nome_combo2"






Set rsCarros33 = Server.CreateObject("ADODB.RecordSet")

	rsCarros33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros33.CursorType = 3
'indica o tipo de cursor utilizão

rsCarros33.ActiveConnection = Conexao3
	
	
	rsCarros33.Open SqlCarros33, Conexao3





'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "Bairro/Região" & "','" & "bqualquer" & "');"& vbcrlf
i = 1 
 
While NOT (rsCarros33.EoF)

Response.Write "doublecombo.combo4.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend

Response.Write "doublecombo.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');"& vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 





rsMarcas33.Close           
		   
           Set rsMarcas33 = Nothing
             
			rsCarros33.Close           
		   
           Set rsCarros33 = Nothing 





End Function
%> 
 
 
 <%
Function EscreveFuncaoJavaScript999 ( Conexao3 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros999 (doublecombo) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (doublecombo.combo4.options[doublecombo.combo4.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!



'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas999 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 ORDER BY nome_combo2" 
Set rsMarcas999 = Conexao333.Execute ( SqlMarcas999 )

While NOT (rsMarcas999.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3 FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""

Set rsCarros999 = Conexao3.Execute ( SqlCarros999 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
Response.Write "doublecombo.combo7.options[" & i  & "] = new Option('" & "Vila" & "','" & "vlqualquer" & "');"& vbcrlf
i = 1 
While NOT (rsCarros999.EoF)

Response.Write "doublecombo.combo7.options[" & i & "] = new Option('" & rsCarros999("nome_combo3") & "','" & rsCarros999("id_combo3") & "');" & vbcrlf 
i=i+1

rsCarros999.MoveNext
Wend


Response.Write "doublecombo.combo7.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "vlqualquer" & "');"


'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas999.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 





rsMarcas999.Close           
		   
           Set rsMarcas999 = Nothing
             
			rsCarros999.Close           
		   
           Set rsCarros999 = Nothing 




End Function
%> 
 
  
  
  <%  EscreveFuncaoJavaScript ( Conexao3 ) %>
  <%  EscreveFuncaoJavaScript2 ( Conexao3) %>


<%EscreveFuncaoJavaScript999 ( Conexao3 )%>


 <%
conexao3.close
set conexao3 = nothing


conexao333.close
set conexao333 = nothing


set conexao33 = nothing


%> 
</body>
</html>

