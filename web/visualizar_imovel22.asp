<!--#include file="dsn.asp"-->

<!--#include file="cores.asp"-->

<% response.buffer=True%>





<%
'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1   FROM combo1 ORDER BY nome_combo1" 





Set rs3 = Server.CreateObject("ADODB.RecordSet")

	rs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs3.ActiveConnection = Conexao3
	
	
	rs3.Open Sql3, Conexao3





Sql33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1   FROM combo1 ORDER BY nome_combo1" 





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











<%
Dim Conexao,strSQL,rs,intRecordCount,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")




if varCod_imovel = "" then
varCod_imovel = "0"
end if

 Set rs = Server.CreateObject("ADODB.RecordSet")
strSQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where cod_imovel="&varCod_imovel


Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   
RS.CursorLocation = 3
RS.CursorType = 3
RS.ActiveConnection = Conexao

        rs.Open strSQL, Conexao 

if not (rs.eof)then



varSucesso_imovel = request.QueryString("varSucesso_imovel")
   dim objFSO
  
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
   
   
  
   dim rs4,strSQL4
   
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
	
	
          rs4.CursorLocation = 3
          rs4.CursorType = 3
          rs4.ActiveConnection = Conexao
	
    
		rs4.Open strSQL4, Conexao
		
	
	
	
	
	dim rs555,strSQL555
   
    Set rs555 = Server.CreateObject("ADODB.RecordSet")
	strSQL555 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 where nome_combo1 ='"&rs("cidade")&"'  ORDER BY nome_combo1" 
	 
	 rs555.CursorLocation = 3
     rs555.CursorType = 3
     rs555.ActiveConnection = Conexao
	 
	 
	 
	 rs555.Open strSQL555, Conexao 
	
	
	
	dim Sql4Bairro,rs4Bairro
	  Set rs4Bairro = Server.CreateObject("ADODB.RecordSet")
Sql4Bairro = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where nome_combo2 like '"& rs("bairro") &"' and cidade_combo2 like '"& rs("cidade") &"' ORDER BY nome_combo2" 

        rs4Bairro.CursorLocation = 3
         rs4Bairro.CursorType = 3
           rs4Bairro.ActiveConnection = Conexao

           rs4Bairro.Open Sql4Bairro, Conexao

dim rs444Vila,strSQL444Vila
   
    Set rs444Vila = Server.CreateObject("ADODB.RecordSet")
	strSQL444Vila = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where nome_combo3 ='"&rs("vila")&"' and bairro_combo3 ='"&rs("bairro")&"' and cidade_combo3 ='"&rs("cidade")&"'   ORDER BY nome_combo3" 
	 
	 
	 rs444Vila.CursorLocation = 3
         rs444Vila.CursorType = 3
           rs444Vila.ActiveConnection = Conexao
	 
	 
	 rs444Vila.Open strSQL444Vila, Conexao		


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



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow44(abrejanela44) {
   openWindow44 = window.open(abrejanela44,'openWin44','width=603,height=500,resizable=yes,left=100,scrollbars=yes')
   openWindow44.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow55(abrejanela55) {
   openWindow55 = window.open(abrejanela55,'openWin55','width=603,height=500,resizable=yes,left=200,scrollbars=yes')
   openWindow55.focus( )
   }

</SCRIPT>



<html>

<title>Imóvel</title>
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
	
	
	if (doublecombo.txt_telefone.value == "") {
        alert("Você precisa indicar o telefone do comprador!");
        doublecombo.txt_telefone.focus();
		doublecombo.txt_telefone.select();
        return false;
    }
	
	
		
	if (doublecombo.txt_valor_vend.value == "0,00") {
        alert("Você precisa indicar um valor para o imóvel!");
        doublecombo.txt_valor_vend.focus();
		doublecombo.txt_valor_vend.select();
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



	
var strValidNumber1_total="1234567890";
for (nCount=0; nCount < doublecombo.txt_a_total_vend.value.length; nCount++) 
		{
strTempChar1_total=doublecombo.txt_a_total_vend.value.substring(nCount,nCount+1);
if (strValidNumber1_total.indexOf(strTempChar1_total,0)==-1) 
{
alert("A área total só pode conter números!");
doublecombo.txt_a_total_vend.focus();
doublecombo.txt_a_total_vend.select();
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













	
	if (doublecombo.stage22.value == "") {
        alert("O formulário valor do Imóvel pretendido está vazio!");
        doublecombo.stage22.focus();
		doublecombo.stage22.select();
        return false;
    }


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


<!--#include file="loggedin.asp"-->


<body  bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >


<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="atualizar_imovel22.asp?varCod_imovel=<%=varCod_imovel%>">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="590" height="48"><a href="visualizar_imovel22.asp?varCod_imovel=<%=varCod_imovel%>"><img src="top_resultado.jpg" width="590" height="48" border="0"></a></td>
  </tr>
  
  
   <tr>
    <td width="590" height="20"></td>
  </tr>
  
  <tr>
    <td width="590" height="120"><div align="center"><img src="simbol_imovel02.jpg"></img></div></td>
  </tr>
  
   <tr>
    <td width="590" height="30"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="imprimir_imovel22.asp?varCod_imovel=<%=varCod_imovel%>" style="color:#FFFFFF">Visualizar 
          impress&atilde;o</a></strong></font></div></td>
  </tr>
  
  
  
  <tr><td width="590" height="18">&nbsp;</td></tr>
  
  <tr><td height="18"><table width="590" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td><table width="580" border="0" cellspacing="0" cellpadding="0">
                     <tr>
                  <td width="290" height="18" bgcolor="<%=escuro%>" > 
                    <%
				   dim varRs444Permuta
	if rs("cod_permuta") <> "" then
	varRs444Permuta = rs("cod_permuta")
	else
	varRs444Permuta = "0"
	end if
				   
				   
				   
				   dim rs444Permuta,SQL444Permuta
 Set rs444Permuta = Server.CreateObject("ADODB.RecordSet")
 SQL444Permuta = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone='"& rs("telefone")&"' order by cod_permuta DESC" 
	
	
	rs444Permuta.CursorLocation = 3
         rs444Permuta.CursorType = 3
           rs444Permuta.ActiveConnection = Conexao
	
	
	rs444Permuta.open SQL444Permuta,Conexao,2,1  
	
			
	if  not rs444Permuta.eof then
				  
				 while not rs444Permuta.eof 
				  %>
                    <div align="center"><a href="javascript:newWindow55('visualizar_permuta22.asp?varCodPermuta=<%=rs444Permuta("cod_permuta")%>')"><img src="bt_foto22perm.jpg" width="290" height="18" border="0"></a></div>
                   
				   <%
				   rs444Permuta.movenext
				   wend
				   %>
				   
				    <%else%>
                    <%end if%>
                  </td>
				  
				  <td width="290" height="18" bgcolor="<%=escuro%>" > 
                    <%
				   dim varRs444Comprador
	if rs("cod_comprador") <> "" then
	varRs444Comprador = rs("cod_comprador")
	else
	varRs444Comprador = "0"
	end if
				   
				   
				   
				   dim rs444Comprador,SQL444Comprador
 Set rs444Comprador = Server.CreateObject("ADODB.RecordSet")
 SQL444Comprador = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone='"& rs("telefone")&"' order by cod_compradores DESC" 
	
	
	rs444Comprador.CursorLocation = 3
         rs444Comprador.CursorType = 3
           rs444Comprador.ActiveConnection = Conexao
	
	
	rs444Comprador.open SQL444Comprador,Conexao,2,1  
	
			
	if  not rs444Comprador.eof then
				  
				  
				  while not rs444Comprador.eof
				  
				  %>
                    
					
					
					<div align="center"><a href="javascript:newWindow44('visualizar_compradores22.asp?varCodCompradores=<%=rs444Comprador("cod_compradores")%>')"><img src="bt_foto22Compr.jpg" width="290" height="18" border="0"></a></div>
                    
					<%
					
					rs444Comprador.movenext
					wend
					
					
					%>
					
					<%else%>
                    <%end if%>
                  </td>
              </tr>
                    </table></td>
            <td width="5">&nbsp;</td>
          </tr>
        </table></td></tr>
  <tr>
  <td>
  <table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
            <td><table width="580" border="0" cellspacing="0" cellpadding="0" style="border:1px solid #FFFFFF;">
            
			 
			  <tr>
                <td width="580" height="334" bgcolor="<%=escuro%>"><% If objFSO.FileExists(Server.MapPath(rs("Foto_grande"))) = True Then%>
                    <div align="center"><img src="<%=rs("foto_grande")%>" name="photoslider" width="580" height="334"></img></div>
                      <% else %>
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Foto 
                      não disponível</strong></font></div>
                    <% end if %></td>
                
				
				
				
				
              </tr>
            </table></td>
          <td width="5">&nbsp;</td>
        </tr>
      </table>
  </td>
  </tr>
  
  <tr>
  <td>
  <table width="590" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td width="580"><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
				  <script language="JavaScript">
                         var photos=new Array()
                         var which=0
                         
photos[0]="<%=rs("foto_grande1")%>"
photos[1]="<%=rs("foto_grande2")%>"
photos[2]="<%=rs("foto_grande3")%>"
photos[3]="<%=rs("foto_grande4")%>"
photos[4]="<%=rs("foto_grande5")%>"
photos[5]="<%=rs("foto_grande6")%>"
photos[6]="<%=rs("foto_grande7")%>"
photos[7]="<%=rs("foto_grande8")%>"
photos[8]="<%=rs("foto_grande9")%>"
photos[9]="<%=rs("foto_grande10")%>"


 var tam = 0;
<% if rs("foto_grande1")<>"imovel00000.jpg"  then%>
                         var tam = 0;
						<%end if%>

<% if rs("foto_grande2")<>"imovel00000.jpg"  then %>
                         var tam = 1;
						<%end if%>
						
<% if rs("foto_grande3")<>"imovel00000.jpg"  then%>
                         var tam = 2;
						<%end if%>					 
                       
 <% if rs("foto_grande4")<>"imovel00000.jpg"  then %>
                         var tam = 3;
						<%end if%>
						
<% if rs("foto_grande5")<>"imovel00000.jpg"  then %>
                      var tam = 4;
						<%end if%>
						
<% if rs("foto_grande6")<>"imovel00000.jpg"  then %>
                      var tam = 5;
						<%end if%>
						
<% if rs("foto_grande7")<>"imovel00000.jpg"  then %>
                      var tam = 6;
						<%end if%>
												
<% if rs("foto_grande8")<>"imovel00000.jpg"  then %>
                      var tam = 7;
						<%end if%>
						
						
<% if rs("foto_grande9")<>"imovel00000.jpg"  then %>
                      var tam = 8;
						<%end if%>
																							
<% if rs("foto_grande10")<>"imovel00000.jpg"  then %>
                      var tam = 9;
						<%end if%>					   
					     function anterior(){
                           if (which>0){
                             which--
                           }else{
                             which=tam;
                           }
                           document.images.photoslider.src=photos[which]
                         }
                         function proxima(){
                           if (which<tam){
                             which++
                           }else{
                             which=0;
                           }
                           document.images.photoslider.src=photos[which]
                         }
                      </script>
                  <td width="290">&nbsp;</td>
                  <td width="290"> <table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:anterior()" class="link" onmouseover="window.status='Anterior'; return true" onmouseout="window.status=''"><img src="bt_anterior002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                        <td><% if  rs("foto_grande2")<>"imovel00000.jpg"  then%><a href="javascript:proxima()" class="link" onmouseover="window.status='Próxima'; return true" onmouseout="window.status=''"><img src="bt_proxima002.jpg" width="145" height="18" border="0"></a><%else%><%end if%></td>
                      </tr>
                    </table> </td>
                </tr>
              </table></td>
            <td width="5">&nbsp;</td>
          </tr>
        </table>
  
  </td>
  </tr>
  
  <tr>
      <td><table width="590" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="570"><div align="center">
			  <%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
        <%else%>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
        foi atualizado com sucesso.</font> 
        <% end if %>

			  
			  
			  
			  </div></td>
              
            <td> <a href="visualizar_fotos.asp?varCodImovel=<%=varCod_imovel%>"><img src="bt_mais03.jpg" width="18" height="18" border="0"></a></td>
            </tr>
          </table></td>
  </tr>
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="5">&nbsp;</td>
          <td><table width="580" border="0" cellspacing="0" cellpadding="0">
             
			 
			 <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quem atualizou esse imóvel</font></div></td>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="" type="text" value="<%if rs("quem_atualizou") <> "" then response.write rs("quem_atualizou") else response.write "não informado" end if%>"  size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"></td>
              </tr>
			 
			 
			 
			 
			<tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">data 
                      futuro contato</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data_futuro_contato" type="text" value="<%if rs("data_futuro_contato") <> "" then response.write rs("data_futuro_contato") else response.write "não informado" end if%>" id="txt_data_futuro_contato" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
			  
			  
			  <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto 
                      futuro contato</font></div></td>
                  <td height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_assunto_futuro_contato" COLS=20 ROWS=10 class="inputBox" id="txt_assunto_futuro_contato"  style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%if rs("assunto_futuro_contato") <> "" then response.write rs("assunto_futuro_contato") else response.write "não informado" end if%></textarea></td>
              </tr>
			   
			 
			 
			 
			 
			 <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Código de referência do imóvel</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs("cod_imovel")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			    <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Link para o imóvel</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%="http://www.imobiliariaveja.com.br/mostrar_imovel2.asp?varCodimovel="&rs("cod_imovel")%>" size="38" maxlength="200" align="left"></td>
              </tr> 
			 
			 
			 
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      de inclus&atilde;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_data" type="text" class="inputBox" id="txt_data" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs("data")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			  
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      da &uacute;ltima atualiza&ccedil;&atilde;o</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_atualizacao" type="text" class="inputBox" id="txt_data_atualizacao" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs("data_atualizacao")%>" size="38" maxlength="50" align="left"></td>
              </tr>
			 
			 
			 
			 
			  <tr>
                  <td height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Capta&ccedil;&atilde;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
				
				<%if  session("permissao") = "3" or session("permissao") = "6" then%>
				
				<select name="txt_captacao" id="txt_captacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                     <option value="<%=rs("captacao")%>" selected ><%=rs("captacao")%></option>
					  <option value="Internet">Internet</option>
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
                    </select>
					<%else%>
					<input name="txt_captacao" type="hidden" class="inputBox" id="txt_captacao" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs("captacao")%>" size="38" maxlength="50" align="left">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("captacao")%></font>
					<%end if%>
					</td>
              </tr>
			 
			 
			 <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Origem 
                      da capta&ccedil;&atilde;o</font></div></td>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_origem_captacao" value="<%if rs("origem_captacao") <> "" then response.write rs("origem_captacao") else response.write "não informado" end if%>" type="text" id="txt_origem_captacao" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"></td>
              </tr>
			  
			
			 <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo cadastramento</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_responsavel_cadastramento" value="<%if rs("responsavel_cadastramento") <> "" then response.write rs("responsavel_cadastramento") else response.write "não informado" end if%>" type="text" id="txt_responsavel_cadastramento" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>"></td>
              </tr>
			 
			 
			 
			 
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Presen&ccedil;a 
                      na Primeira P&aacute;gina</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_presenca_primeira" id="txt_presenca_primeira" size="1" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="<% if rs("presenca_primeira") <> "" then response.write rs("presenca_primeira") else response.write "excluido" end if%>"selected><% if rs("presenca_primeira") <> "" then response.write rs("presenca_primeira") else response.write "excluido" end if%></option>
					  <option value="incluido">Incluído</option>
                      <option value="excluido">Excluído</option>
                    </select>
                  </td>
                </tr>
			
			 
			 
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome 
                      do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>" value="<%=rs("proprietario")%>" size="38" maxlength="50" align="left"></td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      residencial do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>" value="<%=rs("telefone")%>" size="38" maxlength="20" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      comercial do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone02" value="<%if rs("telefone02") <> "" then response.write rs("telefone02") else response.write "não informado" end if%>" type="text" id="txt_telefone02" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>">
                  </td>
              </tr>
			  
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Telefone 
                      celular do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone03" value="<%if rs("telefone03") <> "" then response.write rs("telefone03") else response.write "não informado" end if%>" type="text" id="txt_telefone03" size="38" maxlength="20" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                  </td>
              </tr>
			  
			  
			  
			  
			  
			  
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
                      do propriet&aacute;rio</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="HEIGHT: 18px; WIDTH: 290px ; background:<%=claro%>" value="<%=rs("email")%>" size="38" maxlength="50" align="left"></td>
              </tr>
              
			  
			   <tr bgcolor="<%=medio%>"> 

                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo 
                      do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_titulo_anuncio_vend" type="text" class="inputBox" id="txt_titulo_anuncio_vend" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs("titulo_anuncio")%>" size="38" maxlength="40" align="left"></td>
              </tr>
			  <tr bgcolor="<%=claro%>"> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Texto</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    do An&uacute;ncio</font> </div></td>
                <td width="290" height="18" style="border:1px solid #FFFFFF;"><input name="txt_texto_anuncio_vend" type="text" class="inputBox" id="txt_texto_anuncio_vend" style="HEIGHT: 18px; WIDTH: 290px; background: <%=claro%>" value="<%=rs("texto_anuncio")%>" size="38" maxlength="120" align="left"></td>
              </tr>
			  
			   
			
			 <tr> 
              <td width="290" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF; background:<%=medio%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o 
                      do im&oacute;vel</font></div></td>
                  <td width="290" style="border:1px solid #FFFFFF; background:<%=medio%>"><input name="txt_endereco_vend" type="text" class="inputBox" id="txt_endereco_vend" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>" value="<%=rs("endereco")%>" size="38" maxlength="50" align="left"></td>
            </tr>
			
			<tr> 
              <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Chaves 
                      do im&oacute;vel</font></div></td>
                  <td width="290" style="border:1px solid #FFFFFF; background:<%=claro%>"><input name="txt_chaves_do_imovel" value="<%if rs("chaves_do_imovel") <> "" then response.write rs("chaves_do_imovel") else response.write "não informado" end if%>" type="text" id="txt_chaves_do_imovel" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>">
                  </td>
            </tr>
			
			
			
			
			
			  <tr> 
                  <td bgcolor="<%=medio%>" height="18" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      do im&oacute;vel </font></div></td>
                  <td bgcolor="<%=medio%>" height="18" style="border:1px solid #FFFFFF;"> 
                    <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" onChange="javascript:atualizacarros2(this.form);">
                     <option value="<% if rs("cidade") = "não informado" or rs555.eof then response.write "cqualquer" else response.write rs555("id_combo1") end if  %>" select><%if   rs("cidade") = "cqualquer" then response.write "não informado" else response.write rs("cidade") end if  %></option>
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      do im&oacute;vel </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo4" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" onChange="javascript:atualizacarros999(this.form);">
                      
                        <option value="<%if rs("bairro") = "não informado" or rs4Bairro.eof then response.write "bqualquer" else response.write rs4Bairro("id_combo2") end if%>" ><%if   rs("bairro") = "bqualquer" then response.write "não informado" else response.write rs("bairro") end if  %></option>
                       
                        <option value=""></option>
                  </select>
                    <a href="javascript:newWindow3('form_incluir_bairro.asp')"><img src="bt_mais02.jpg" width="18" height="18" border="0"></a> 
                  </td>
                </tr>
                 <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila 
                      do im&oacute;vel </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo7" class="inputBox" id="combo7" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                     <option value="<%if rs("vila") <> "não informado" and  rs("vila") <>"" and  not rs444Vila.eof then response.write rs444Vila("id_combo3") else response.write "vlqualquer" end if%>" selected><%if rs("vila") <> "não informado" and  rs("vila") <>"" then response.write rs("vila") else response.write "não informado" end if%></option>
				  <option value="vlqualquer">qualquer um</option>
                    </select>
                    <a href="javascript:newWindow3('form_incluir_vila.asp')"><img src="bt_mais01.jpg" width="18" height="18" border="0"></a> 
                  </td>
              </tr>
			  
			  
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel </font></div></td>
                  <td style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo_vend" size="1" id="txt_tipo_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="<%=rs("tipo")%>" selected><%if   rs("tipo") = "tqualquer" then response.write "não informado" else response.write rs("tipo") end if  %></option>					 
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      total / Terreno</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <input name="txt_a_total_vend" type="text" class="inputBox" id="txt_a_total_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="<%=rs("area_total")%>" size="12" maxlength="20">
                    <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font></td>
              </tr>
                <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
                      constru&iacute;da (para casas) &aacute;rea &uacute;til (para 
                      aptos) </font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_a_constr_vend" type="text" class="inputBox" id="txt_a_constr_vend" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>" value="<%=rs("area_construida")%>" size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m&sup2;</font> </font> </td>
              </tr>
			  
			  <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      de frente</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_de_frente" type="text" value="<%if rs("metros_de_frente") <> "" then response.write rs("metros_de_frente") else response.write "00" end if%>" class="inputBox" id="txt_metros_de_frente" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>"  size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			   <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      de fundo</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_de_fundo" type="text" value="<%if rs("metros_de_fundo") <> "" then response.write rs("metros_de_fundo") else response.write "00" end if%>" class="inputBox" id="txt_metros_de_fundo" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>"  size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			   <tr bgcolor="<%=medio%>"> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      lateral esquerda</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_lateral_esquerda" type="text" value="<%if rs("metros_lateral_esquerda") <> "" then response.write rs("metros_lateral_esquerda") else response.write "00" end if%>" class="inputBox" id="txt_metros_lateral_esquerda" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>"  size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			   <tr bgcolor="<%=claro%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Metros 
                      lateral direita</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <input name="txt_metros_lateral_direita" value="<%if rs("metros_lateral_direita") <> "" then response.write rs("metros_lateral_direita") else response.write "00" end if%>" type="text" class="inputBox" id="txt_metros_lateral_direita" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>"  size="12" maxlength="20">
                    <font color="#FFFFFF"> <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    m</font> </font> </td>
              </tr> 
			  
			  
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Quartos 
                      do im&oacute;vel </font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos_vend" id="txt_quartos_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                       <option value="<%=rs("quartos")%>" selected><% if rs("quartos") = "0" then response.write "não informado" else response.write rs("quartos") end if%></option>
                    <option value="não informado">não informado</option>
					
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
                      <option value="<%if rs("suites") <> "" then response.write rs("suites") else response.write "não informado" end if%>" selected><%if rs("suites") <> "" then response.write rs("suites") else response.write "não informado" end if%></option>
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Banheiros 
                      do im&oacute;vel </font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_banheiros_vend" id="txt_banheiros_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     <option value="<%=rs("banheiros")%>" selected><%=rs("banheiros")%></option>
                      <option value="não informado">não informado</option>
				    
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
                      na Garagem do im&oacute;vel</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas_vend" id="txt_vagas_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
                     <option value="<%=rs("vagas")%>" selected><% if rs("vagas") = "0" then response.write "não informado" else response.write rs("vagas") end if%></option>
                      <option value="não informado">não informado</option>
                   
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      </font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_negociacao_vend" id="txt_negociacao_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      
                     <option value="<%=rs("negociacao")%>" selected><%if rs("negociacao") = "nqualquer" then response.write "qualquer uma" else response.write rs("negociacao") end if%></option>
					 
					  <option value="aluguel" >Aluguel</option>
                      <option value="venda">Venda</option>
					  
                     
                     
                  </select>
                  </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do im&oacute;vel </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_valor_vend" type="text" class="inputBox" id="txt_valor_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="<%=FormatNumber(rs("valor"),2)%>" size="12" maxlength="13">
                  </td>
              </tr>
			  
			   <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      do condom&iacute;nio</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_condominio_vend" type="text" class="inputBox" id="txt_condominio_vend" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="<%if rs("condominio") <> "" then response.write FormatNumber(rs("condominio"),2) else response.write "0,00" end if%>" size="12" maxlength="13">
                  </td>
              </tr>
			  
			  
			   <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Saldo 
                      devedor </font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_saldo_devedor" id="txt_saldo_devedor" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="<%if rs("saldo_devedor") <> "" then response.write rs("saldo_devedor") else response.write "sem saldo devedor" end if%>" selected ><%if rs("saldo_devedor") <> "" then response.write rs("saldo_devedor") else response.write "sem saldo devedor" end if%></option>
					  <option value="sem saldo devedor" >Sem saldo devedor</option>
					  <option value="com saldo devedor" >Com saldo devedor</option>
                    
                    
                  </select>
                  </td>
              </tr>
			  
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      j&aacute; pago no saldo devedor</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_ja_pago_devedor" value="<%if rs("ja_pago_devedor") <> "" then response.write formatNumber(rs("ja_pago_devedor"),2) else response.write "0,00" end if %>" type="text" class="inputBox" id="txt_ja_pago_devedor" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>"  size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  
			  <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
                      devido no saldo devedor</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_devendo_devedor" type="text" class="inputBox" id="txt_devendo_devedor" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>" value="<%if rs("devendo_devedor") <> "" then response.write formatNumber(rs("devendo_devedor"),2) else response.write "0,00" end if%>" size="12" maxlength="30">
                  </td>
              </tr>
			  
			  
			  
			  
			  
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Placa</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_placa_vend" id="txt_placa_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                     
					  <option value="<%if rs("placa") <> "" then %><%=rs("placa")%><%else%><%="Sem Placa"%><%end if%>" select><%if rs("placa") <> "" then %><%=rs("placa")%><%else%><%="Sem Placa"%><%end if%></option>
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
                     
					 <% if session("permissao") = "6" then %>
					  <option value="<%=rs("standby")%>" selected><%=rs("standby")%></option>
                      <option value="excluido">Excluído</option>
                      <option value="incluido">Incluído</option>
                   <%else%>
				   
				   <option value="<%=rs("standby")%>" selected><%=rs("standby")%></option>
				   <%end if%>
				    </select>
                  </td>
              </tr>
			  
			  
			  
			  <tr bgcolor="<%=medio%>"> 
                  <td style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Imóvel vendido/suspenso/com proposta</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_imovel_em_negociacao" id="txt_imovel_em_negociacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
                      <option value="<%if rs("imovel_em_negociacao") <> "" then response.write rs("imovel_em_negociacao") else response.write "não informado" end if%>" selected><%if rs("imovel_em_negociacao") <> "" then response.write rs("imovel_em_negociacao") else response.write "não informado" end if%></option>
					  
					   <option value="não informado">Não informado</option>
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
                     <option value="<%=rs("ocupacao")%>" selected><%=rs("ocupacao")%></option>
					 <option value="não informado">não informado</option>
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
                       <option value="<%if rs("qualidade") <> "" then response.write rs("qualidade") else  response.write "não informado" end if%>" selected><%if rs("qualidade") <> "" then response.write rs("qualidade") else  response.write "não informado" end if%></option>
					  <option value="bom negócio" >Bom Negócio</option>
                    <option value="negócio comum" >Negócio Comum</option>
                    
                  </select></td>
              </tr>
			  
			  
			   <tr>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Melhor 
                      hor&aacute;rio para visita</font></div></td>
                  <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="txt_melhor_horario_visita" size="1" class="inputBox" id="txt_melhor_horario_visita"  style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px;  color:FFFFFF; background: <%=claro%>">
                     
					  <option value="<%if rs("melhor_horario_visita") <> "" then response.write rs("melhor_horario_visita") else response.write "não informado" end if%>" selected><%if rs("melhor_horario_visita") <> "" then response.write rs("melhor_horario_visita") else response.write "não informado" end if%></option>
                     
					 
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
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">detalhes 
                      sobre o im&oacute;vel e se poss&iacute;vel o motivo da venda</font></div></td>
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="obs_imovel_vend" class="inputBox" id="obs_imovel_vend" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%=rs("obs_imovel")%></textarea></td>
              </tr>
              <tr bgcolor="<%=claro%>">
                  <td width="290" height="100" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dados</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      confidenciais do propriet&aacute;rio</font></div></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="obs_proprietario_vend" class="inputBox" id="obs_proprietario_vend" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>" onKeyPress="return limitfield(this, 800)"><%=rs("obs_proprietario")%></textarea></td>
              </tr>
			 
			  <tr>
                  <td width="290" height="180"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="290" height="140"><div align="left"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Caso 
                      esse cliente aceite um im&oacute;vel como parte de pagamento 
                      deste im&oacute;vel que foi cadastrado acima ou queira dar 
                      este im&oacute;vel como parte de pagamento na compra de 
                      outro im&oacute;vel, responda &quot;sim&quot; na pergunta 
                      abaixo:</strong></font></div></td>
			  </tr>
			   
			  <tr> 
                  <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>"> 
                    <div align="center">&nbsp;</div></td>
                  <td width="290" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF; background:<%=claro%>"> 
                    <select name="txt_pergunta" id="txt_pergunta" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                
				 <% if session("permissao") = "5" or session("permissao") = "6" or session("permissao") = "2" then%>
				  <option value="sim">Sim</option>
                  <option value="nao" selected>Não</option>
               <%else%>
			   <option value="nao" selected>Não</option>
			   <%end if%>
				
				</select>
                  </td>
            </tr> 
			
			 <tr>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Respons&aacute;vel 
                      pelo cadastramento</font></div></td>
                  <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_responsavel_cadastramento_comprador" type="text" id="txt_responsavel_cadastramento_comprador" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>"></td>
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atendimento</font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><select name="txt_atendimento" id="txt_atendimento" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=medio%>">
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
                    </select></td>
              </tr>
			   
			   
			   
			   
			   
			   
            <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Origem 
                      do Comprador</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_origem" id="txt_origem" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                     
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
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
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vila</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      onde quer comprar ou alugar im&oacute;vel</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="combo5" class="inputBox" id="combo5" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                     <option value="vlqualquer" selected>Vila</option>
				  <option value="vlqualquer">qualquer um</option>
                    </select>
                  </td>
              </tr>
			  
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
                    <select name="txt_tipo" size="1" id="txt_tipo" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de quartos do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_quartos" id="txt_quartos" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
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
                      na garagem do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=claro%>">
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
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ocupa&ccedil;&atilde;o</font><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>">
                    <option value="não informado" selected>não informado</option>
                   
                    <option value="vago">Vago</option>
					 <option value="ocupado por terceiros">Ocupado por terceiros</option>
                    <option value="ocupado pelo inquilino">Ocupado pelo inquilino</option>
					<option value="ocupado pelo proprietário">Ocupado por terceiros</option>
				  </select>
                  </td>
              </tr>
              
              
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      que deseja</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><select name="example2" size="1" class="inputBox" id="example2"  style="HEIGHT: 18px; WIDTH: 149px ; font-size : 10px;  color:FFFFFF; background: <%=claro%>">
                      
                      <option  value="aluguel">Aluguel </option>
                      <option value="compra" selected>Compra </option>
                    </select> </td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Faixa 
                      de pre&ccedil;o desejada</font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="stage22" type="text" class="inputBox" id="stage22" style="HEIGHT: 18px; WIDTH: 150px; background:<%=medio%>" value="0,00" size="12" maxlength="13"> 
                  </td>
              </tr>
              <tr bgcolor="<%=claro%>"> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">StandBy</font></div></td>
                  <td style="border:1px solid #FFFFFF;"> 
                    <select name="txt_standby" id="txt_standby" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: <%=claro%>">
                      <option value="excluido" selected>Excluído</option>
                    <option value="incluido">Incluído</option>
                    
                  </select>
                  </td>
              </tr>
              <tr bgcolor="<%=medio%>">
                  <td width="290" height="100" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">D&ecirc; 
                      mais detalhes sobre o im&oacute;vel desejado por esse cliente 
                      e diga de que forma ele pretende pagar</font></div></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" COLS=20 ROWS=10 class="inputBox" id="txt_descricao" style="HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			 
			   <tr bgcolor="<%=claro%>">
                  <td width="290" height="100" style="border:1px solid #FFFFFF;" ><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Coloque 
                      aqui os dados confidenciais desse cliente</font></div></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao_confi" COLS=20 ROWS=10 class="inputBox" id="txt_descricao_confi" style="HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>" onKeyPress="return limitfield(this, 800)"></textarea></td>
              </tr>
			  
              <tr>
                  <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><input name="image" type="image"  src="bt_atualizar002.jpg" width="145" height="18" border="0"></td>
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

<%else%>
<div align="center"><strong><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imóvel 
  não encontrado</font></strong> 
  <% end if%>
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
           rs.Close
                    
           Set rs = Nothing
          
		  '-----------------------------
           rs444captacao22.Close           
		   
           Set rs444captacao22 = Nothing
		   
'---------------------------------





'-----------------------------
           rs4Bairro.Close           
		   
           Set rs4Bairro = Nothing
		   
'---------------------------------


'-----------------------------
           rs333.Close           
		   
           Set rs333 = Nothing
		   
'---------------------------------



'-----------------------------
           rs33.Close           
		   
           Set rs33 = Nothing
		   
'---------------------------------


		  
	'-----------------------------
           rs444Vila.Close           
		   
           Set rs444Vila = Nothing
		   
'---------------------------------	



'-----------------------------
                   
		   
           Set objfso = Nothing
		   
'---------------------------------

  '-----------------------------
           rs44.Close           
		   
           Set rs44 = Nothing
		   
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
           rs444Placa.Close           
		   
           Set rs444Placa = Nothing
		   
'---------------------------------	
		
		
		
	'-----------------------------
           rs444Tipo23.Close           
		   
           Set rs444Tipo23 = Nothing
		   
'---------------------------------



'-----------------------------
           rs444Permuta.Close           
		   
           Set rs444Permuta = Nothing
		   
'---------------------------------




'-----------------------------
           rs444Captacao.Close           
		   
           Set rs444Captacao = Nothing
		   
'---------------------------------




'-----------------------------
           rs555.Close           
		   
           Set rs555 = Nothing
		   
'---------------------------------



'-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------



'-----------------------------
           rs4.Close           
		   
           Set rs4 = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Comprador.Close           
		   
           Set rs444Comprador = Nothing
		   
'---------------------------------
	
		
		  
		  
		  
		   %>
</div>

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
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1   FROM combo1 ORDER BY nome_combo1" 



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


Set rsMarcas999 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas999.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsMarcas999.CursorType = 3
'indica o tipo de cursor utilizão

rsMarcas999.ActiveConnection = Conexao3
	
	
	rsMarcas999.Open SqlMarcas999, Conexao3




While NOT (rsMarcas999.EOF)

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas999("id_combo2") & "':" & vbcrlf







'Caso tenha sido essa marca selecionada... 


'Apagamos tudo o que tem na caixa dos carros!
Response.Write "doublecombo.combo7.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros999 = "SELECT combo3.id_combo3,combo3.id_combo2,combo3.nome_combo3,combo3.data_combo3,combo3.bairro_combo3,combo3.id_combo1,combo3.cidade_combo3  FROM combo3 where id_combo2 =" & rsMarcas999("id_combo2")&""


Set rsCarros999 = Server.CreateObject("ADODB.RecordSet")

	rsCarros999.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsCarros999.CursorType = 3
'indica o tipo de cursor utilizão

rsCarros999.ActiveConnection = Conexao3
	
	
	rsCarros999.Open SqlCarros999, Conexao3






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

'-----------------------------
           conexao.Close           
		   
           Set conexao = Nothing
		   
'---------------------------------



'-----------------------------
           conexao3.Close           
		   
           Set conexao3 = Nothing
		   
'---------------------------------




'-----------------------------
           conexao333.Close           
		   
           Set conexao333 = Nothing
		   
'---------------------------------






%>

</body>
</html>