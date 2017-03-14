<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->

<% response.buffer=True%>


<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "SAO BERNARDO"
end if


'Criando conexão com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn

'Abrindo a tabela MARCAS!
Sql3 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs3 = Conexao3.Execute ( Sql3 ) 

dim objFSO

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

%> 








<%
Dim Conexao,strSQL,rs,varCod_imovel,varSucesso_imovel
varCod_imovel = request.QueryString("varCod_imovel")
varSucesso_imovel = request.QueryString("varSucesso_imovel")
   
 
	
	
	 dim Conexao9,rs9
 Set Conexao9 = Server.CreateObject("ADODB.Connection")
	Set rs9 = Server.CreateObject("ADODB.RecordSet")
	Conexao9.Open dsn
	dim strSQL9
	
	dim varCodCompradores
	varCodCompradores=request.QueryString("varCodCompradores")
	
	dim varCod_Comprador02
	varCod_Comprador02=varCodCompradores
	
	
	
	 strSQL9 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou FROM compradores where cod_compradores="&varCodCompradores
	 rs9.CursorLocation = 3
      rs9.CursorType = 3
	  
	  
	 rs9.Open strSQL9, Conexao9
	 
	 dim vValor
	  vValor=rs9("valor")
   session("vValor")=vValor
   session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)
	 
	
	
	
	
'----------------------------verifica imóvel---------------------------
	
	
	dim rsVerifica2
	dim strSQLVerifica2
	
 Set rsVerifica2 = Server.CreateObject("ADODB.RecordSet")
    
	
	
	
	
	
	strSQLVerifica2 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where telefone='"&rs9("telefone")&"' or telefone02='"&rs9("telefone")&"' or telefone='"&rs9("telefone")&"' "
	 
   
   
rsVerifica2.CursorLocation = 3
rsVerifica2.CursorType = 3

        rsVerifica2.Open strSQLVerifica2, Conexao9 	
		
		
	
	
	
	
	
	
	'-----------------------------------------------------------------	
	
	
	
	
'-----------------incluir em compradores clicados-------------------

dim strSQL003
	 dim rs003
	 
	 Set rs003 = Server.CreateObject("ADODB.RecordSet")
	 
	 strSQL003 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02  FROM compradores where telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"' ORDER BY cod_compradores Desc"
	
	
	
	
	


   
RS003.CursorLocation = 3
RS003.CursorType = 3

dim vAtendimento02

dim vClienteDaInternet


        rs003.Open strSQL003, Conexao9 
	  if not rs003.eof then
	  vAtendimento02 = rs003("atendimento")
	  vClienteDaInternet = rs003("cod_compradores")
	  else
	  vClienteDaInternet = ""
	  vAtendimento02 = "não informado"
	  
	  end if
	 
	 if session("origem") <> "" then
	 session("origem") = "Site"
	 else
	 session("origem") = "Email"
	 end if
	 
	 
	 dim EnderecoIP

EnderecoIP = request.ServerVariables("REMOTE_ADDR")
	
	dim vCodImovel
	
	
	
	if not rsVerifica2.eof then
	vCodImovel = rsVerifica2("cod_imovel")
	end if
	
	
	 if not rs003.eof then
	  
	  Conexao9.execute"Insert into comprador_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao,atendimento,origem,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& vClienteDaInternet &"','"& EnderecoIP &"','"& now() &"','"& rs9("tipo") &"','"& rs9("quartos") &"','"& rs9("vagas") &"','"& rs9("cidade") &"','"& rs9("bairro") &"','"& rs9("valor") &"','"& rs9("negociacao") &"','"& vAtendimento02 &"','"& session("origem") &"','"& session("vOrigem_Franquia") &"')"
	
	   'Conexao9.execute"Insert into comprador_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao,atendimento,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& "00" &"','"& EnderecoIP &"','"& now() &"','"& "0" &"','"& "0"  &"','"& "0"  &"','"& "0"  &"','"& "0"  &"','"& "0" &"','"& "0"  &"','"& vAtendimento02 &"','"& session("origem") &"')"
	
	  else
	  
	  Conexao9.execute"Insert into comprador_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao,atendimento,origem,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& vClienteDaInternet &"','"& EnderecoIP &"','"& now() &"','"& "não informado" &"','"& "00" &"','"& "00" &"','"& "00" &"','"& "não informado" &"','"& "00" &"','"& "não informado" &"','"& "não informado" &"','"& session("origem") &"','"& session("vOrigem_Franquia") &"')"
	
	
	  
	  end if
	
	
	
	 
	'Conexao9.execute"Insert into comprador_clicado(nome,telefone,codigo_clicado,endereco_ip,data,tipo,quartos,vagas,cidade,bairro,valor,negociacao,atendimento,origem) values( '"& session("nome") &"','"& session("telefone") &"','"& rs9("cod_compradores") &"','"& EnderecoIP &"','"& now() &"','"& rs9("tipo") &"','"& rs9("quartos") &"','"& rs9("vagas") &"','"& rs9("cidade") &"','"& rs9("bairro") &"','"& rs9("valor") &"','"& rs9("negociacao") &"','"& vAtendimento02 &"','"& session("origem") &"')"
	
	
	
	
	'----------------------------------------------------------------------------------------







'----------------------------------------------------------------------	 
		
%>		



<script>
function isValidDigitNumber (doublecombo)
{
{





{
if (doublecombo.txt_email.value == "") {
		
	} else {
		prim = doublecombo.txt_email.value.indexOf("@")
		if(prim < 2) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("@",prim + 1) != -1) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(".") < 1) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(" ") != -1) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("zipmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("hotmeil.com") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(".@") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("@.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(".com.br.") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("/") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("[") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("]") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("(") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf(")") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		}
		if(doublecombo.txt_email.value.indexOf("..") > 0) {
			alert("O e-mail informado parece não estar correto.");
			doublecombo.txt_email.focus();
			doublecombo.txt_email.select();
			return false;
		
		
		
		}
		
		
	}

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









	
	if (doublecombo.txt_proprietario.value == "") {
        alert("O formulário Proprietário do Imóvel está vazio!");
        doublecombo.txt_proprietario.focus();
		doublecombo.txt_proprietario.select();
        return false;
    }
	
	if (doublecombo.combo2.value == "") {
        alert("O formulário Bairro do Imóvel está vazio!");
        doublecombo.combo1.focus();
		
        return false;
    }
	
	if (doublecombo.combo1.value == "") {
        alert("O formulário Cidade do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }
	
	
	
	
	if (doublecombo.txt_endereco.value == "") {
        alert("O formulário Endereço do Imóvel está vazio!");
        doublecombo.txt_endereco.focus();
		doublecombo.txt_endereco.select();
        return false;
    }
	
	
	if (doublecombo.blob.value == "") {
        alert("O formulário Foto Grande está vazio!");
        doublecombo.blob.focus();
		doublecombo.blob.select();
        return false;
    }
	
	 vfile = doublecombo.blob.value;
    tfile = vfile.length;
    
    if (vfile.substr(tfile - 4, 4) != ".jpg" && vfile.substr(tfile - 4, 4) != ".gif") {
        alert("O arquivo do formulário Foto Grande deverá possuir o formato (.jpg) ou (.gif)!");
        doublecombo.blob.value == vfile.substr(tfile - 4, 4);
		doublecombo.blob.focus();
		doublecombo.blob.select();
		
		
        return false;
    }
	
	
	

var strVerif = doublecombo.blob.value;
var	strVerif_n = strVerif.length;
if (strVerif.substring(strVerif_n - 15,strVerif_n - 9) != "imovel" ){

       alert("Você escolheu o arquivo errado, o nome do arquivo certo começa com 'imovel' e mais cinco numerais.");
       doublecombo.blob.focus();
		doublecombo.blob.select();
		
		
		
return false;

}


var strVerif2 = doublecombo.blob.value;
var	strVerif_n2 = strVerif2.length;
if (strVerif2.substring(strVerif_n2 - 15,strVerif_n) == "imovel00000.jpg" ){

       alert("Você escolheu o arquivo errado, imovel00000.jpg não pode ser enviado.");
       doublecombo.blob.focus();
		doublecombo.blob.select();
		
		
		
return false;

}


	

//---------------------------------configuração do combo foto_pequena---------------------


	if (doublecombo.blob2.value == "") {
        alert("O formulário Foto Pequena está vazio!");
        doublecombo.blob2.focus();
		doublecombo.blob2.select();
        return false;
    }
	
	 vfile2 = doublecombo.blob2.value;
    tfile2 = vfile2.length;
    
    if (vfile2.substr(tfile2 - 4, 4) != ".jpg" && vfile2.substr(tfile2 - 4, 4) != ".gif") {
        alert("O arquivo do formulário Foto Pequena deverá possuir o formato (.jpg) ou (.gif)!");
        doublecombo.blob2.value == vfile2.substr(tfile2 - 4, 4);
		doublecombo.blob2.focus();
		doublecombo.blob2.select();
		
		
        return false;
    }
	
	
	

var strVerif2 = doublecombo.blob2.value;
var	strVerif_n2 = strVerif2.length;
if (strVerif2.substring(strVerif_n2 - 20,strVerif_n2 - 9) != "mini_imovel" ){

       alert("Você escolheu o arquivo errado, o nome do arquivo certo começa com 'mini_imovel' e mais cinco numerais.");
       doublecombo.blob2.focus();
		doublecombo.blob2.select();
		
		
		
return false;

}


var strVerif3 = doublecombo.blob2.value;
var	strVerif_n3 = strVerif3.length;
if (strVerif3.substring(strVerif_n3 - 20,strVerif_n3) == "mini_imovel00000.jpg" ){

       alert("Você escolheu o arquivo errado, mini_imovel00000.jpg não pode ser enviado.");
       doublecombo.blob2.focus();
		doublecombo.blob2.select();
		
		
		
return false;

}




//--------------------------------------------------------------------










	
	
	
	
	
		var strValidNumber1_5="1234567890,";
for (nCount=0; nCount < doublecombo.txt_a_total.value.length; nCount++) 
		{
strTempChar1_5=doublecombo.txt_a_total.value.substring(nCount,nCount+1);
if (strValidNumber1_5.indexOf(strTempChar1_5,0)==-1) 
{
alert("O formulário Área Total só pode conter números!");
doublecombo.txt_a_total.focus();
doublecombo.txt_a_total.select();
return false;
}
}
	
	

	

var strValidNumber1_4="1234567890,";
for (nCount=0; nCount < doublecombo.txt_a_constr.value.length; nCount++) 
		{
strTempChar1_4=doublecombo.txt_a_constr.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1) 
{
alert("O formulário Área Construída só pode conter números!");
doublecombo.txt_a_constr.focus();
doublecombo.txt_a_constr.select();
return false;
}
}



if (doublecombo.txt_valor.value == "") {
        alert("O formulário Valor está vazio!");
        doublecombo.txt_valor.focus();
		doublecombo.txt_valor.select();
        return false;
    }



	
	
	
	
	
	
	








	

	
	
	


	var strValidNumber1_6="1234567890,";
for (nCount=0; nCount < doublecombo.txt_valor.value.length; nCount++) 
		{
strTempChar1_6=doublecombo.txt_valor.value.substring(nCount,nCount+1);
if (strValidNumber1_6.indexOf(strTempChar1_6,0)==-1) 
{
alert("O formulário Valor só pode conter números!");
doublecombo.txt_valor.focus();
doublecombo.txt_valor.select();
return false;
}
}

var strText2_4 = doublecombo.txt_valor.value;
var s_strText2_4 = strText2_4.length
if (strText2_4.substring((s_strText2_4 - 3), (s_strText2_4 - 2)) != ","){

       alert("A vírgula do formulário Valor está fora do lugar!");
       doublecombo.txt_valor.focus();
		
		doublecombo.txt_valor.select();
		
return false;

}
//-----------


//----------------------

prim2_4 = doublecombo.txt_valor.value.indexOf(",")
if(doublecombo.txt_valor.value.indexOf(",",prim2_4 + 1) != -1) {
			alert("O formulário Valor não contêm a vírgula do valor-moeda");
			doublecombo.txt_valor.focus();
			doublecombo.txt_valor.select();
			return false;
		}







	
	
	
   
	
	

	
	
}



{







//------------- Verifica se é numérico---------------------



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
//-----------------------------------------------

}
}









</script>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=610,height=490,resizable=yes')
   openWindow3.focus( )
   }

</SCRIPT>



<html>

<title>Comprador</title>
<head>
<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: <%=claro%>;}
</STYLE>

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

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow001(abrejanela001) {
   openWindow001 = window.open(abrejanela001,'openWin001','width=605,height=500,resizable=yes,scrollbars=yes')
   openWindow001.focus( )
   }

</SCRIPT>
<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow002(abrejanela002) {
   openWindow002 = window.open(abrejanela002,'openWin002','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow002.focus( )
   }

</SCRIPT>
</head>

<!--#include file="style_imoveis.asp"-->


<body onload=doublecombo.txt_proprietario.focus(); bgcolor="#f7ecbf" bottommargin="30" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0" >
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="post" action="incluir_compradores01.asp">
<table width="590" border="0" cellspacing="0" cellpadding="0">
  
  
  <tr>
    <td><table width="590" border="0" cellspacing="0" cellpadding="0">
  <tr>
            <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></td>
  </tr>
 <tr>
            <td height="80"><table width="580" height="70" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td> 
              <div align="center"><font color="#9d9249" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
                      a ficha desse comprador abaixo:<br>
				 <br>
				  <%
		  if session("telefone") = "43621135" then
		  
		  response.write "O telefone desse comprador é "&rs9("telefone")
		   end if
		  %>
				
				
				
				</strong></font></div></td>
                </tr>
              </table> </td>
  </tr>
  <tr>
    <td><table width="589" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="3">&nbsp;</td>
          <td width="584"><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
              </tr>
             <tr>
                        <td height="30" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#9d9249" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente</strong></font></div></td>
                        <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#9d9249" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs9("atendimento")%></strong></font></td>
              </tr>
			 
			 
			 
			 <tr>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo 
                            do comprador</font></div></td>
                <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf; font-size:12px" value="<%=rs9("cod_compradores")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Uacute;tima 
                            atualiza&ccedil;&atilde;o </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px"  value="<%=rs9("data_atualizacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			 
			 
			 
			 
              <tr>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Meu 
                      nome </font></div></td>
                <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf;font-size:12px" value="<%=rs9("nome")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Telefone</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
				
				
				<%
		  
		  dim SqlTelefone_acesso
		  dim rsTelefone_acesso
		  
		  SqlTelefone_acesso = "SELECT * FROM telefone_acesso where telefone_acesso like '"&session("telefone")&"' ORDER BY cod_telefone_acesso ASC" 

Set rsTelefone_acesso = Server.CreateObject("ADODB.RecordSet")

	rsTelefone_acesso.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsTelefone_acesso.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsTelefone_acesso.ActiveConnection = Conexao9
	
	
	rsTelefone_acesso.Open sqlTelefone_acesso, Conexao9
		  
		  
		  
		  
		  if not rsTelefone_acesso.eof  then
		  
		  'response.write "O telefone desse imóvel é "&rs("telefone")&" e o endereço "&rs("endereco")
		   'end if
		  %>
          
				
				
				<input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px" value="<%=rs9("telefone")%>" size="38" maxlength="33" align="left">
              <%else%>
			  
			  <input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px" value="não informado" size="38" maxlength="33" align="left">
			  <%end if%>
			  </tr>
                
			  
			 
              
			   <tr>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Email</font></div></td>
                <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf;font-size:12px" value="n&atilde;o informado" size="38" maxlength="33" align="left"></td>
              </tr>
			  
              
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Endere&ccedil;o</font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px" value="n&atilde;o informado" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      que estou interessado</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><input name="txt_cidade" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf;font-size:12px" value="<%=rs9("cidade")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                            que estou interessado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_bairro" type="text" class="inputBox" id="txt_proprietario3" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px" value="<%=rs9("bairro")%>" size="38" maxlength="33" align="left">
                  </td>
              </tr>
              <tr>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;">
                    <input name="txt_tipo" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf;font-size:12px" value="<%=rs9("Tipo")%>" size="38" maxlength="33" align="left">
                    </td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de quartos do im&oacute;vel desejado</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_quartos" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px" value="<%=rs9("quartos")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
			  
			   <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;meros 
                      de vagas na garagem desejado</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf;font-size:12px" value="<%=rs9("vagas")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Negocia&ccedil;&atilde;o 
                      que eu quero</font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;font-size:12px" value="<%=rs9("negociacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Faixa 
                      de pre&ccedil;o que eu quero</font></div></td>
                  <td bgcolor="#f7ecbf" style="border:1px solid #FFFFFF;"><input name="txt_valor" type="text" class="inputBox" id="txt_proprietario3" style="border-color : #f7ecbf;color:#9d9249;HEIGHT: 18px; WIDTH: 290px; background:#f7ecbf;font-size:12px" value=" <%if vValor <> "vqualquer" then%><%=FormatNumber(session("vValor"),2)%> <%else%>não informado<%end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
			 
			                           
			 
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif">Aqui 
                            tem a descri&ccedil;&atilde;o do im&oacute;vel que 
                            eu quero</font></div></td>
                    </tr>
                    <tr> 
                              <td width="290" height="82" bgcolor="#f7ecbf" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="border-color : <%=claro%>;color:#9d9249;HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; font-size:12px" onKeyPress="return limitfield(this, 800)"><%=rs9("descricao")%></textarea></td>
              </tr>
              <tr>
                        <td>&nbsp; </td>
                        <td><a href="comprador_proposta.asp?varCodImovel=<%=varCodImovel%>&varCod_comprador02=<%=varCod_comprador02%>"><img src="bt_agenda04.jpg" width="289" height="18" border="0"></a></td>
              </tr>
            </table></td>
          <td width="10">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table></td>
  </tr>
</table>
</form>
 

<%








	
	
	
 
 
 
 
dim strSQL002


strSQL002 ="SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  FROM imoveis where (telefone like'"& rs9("telefone")&"' or telefone02 like'"& rs9("telefone")&"' or telefone03 like'"& rs9("telefone")&"') ORDER BY cod_imovel Desc"

'----------------------------------------------------Fim da instrução SQL---------------------------------
  
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 



Set RS = Server.CreateObject("ADODB.Recordset")
'um objeto recordset é instânciado.

Dim LinkTemp
'essa variável vai ser usada como contador







dim intPage
'essa variável vai receber um valor inicial "1" que mostra que estamos na primeira página.

dim intPageCount
'Essa variável vai receber o valor da quantidade de páginas do recordset.

dim intRecordCount
'Essa variável vai receber o número de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a variável intPage recebe o valor "1" na primeira página.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conexão o recordset utilizará.
	
RS.Open strSQL002, Conn, 1, 3
'o recordset é aberto
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por página.

RS.CacheSize = RS.PageSize
'o Cache também conterá 20 registros por página.

intPageCount = RS.PageCount
'A variável intPageCount recebe o valor do número de página do recordset retornado.

intRecordCount = RS.RecordCount
'A variável intRecordCount recebe o valor do número de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then



%>







<table width="591" border="0"  cellpadding="0" cellspacing="0">
  <tr>
    <td height="100">
<div align="center"><font color="#FF0000" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Veja 
        abaixo o im&oacute;vel que esse comprador tem para vender,dar como parte 
        de pagamento ou permutar.</strong></font></div></td>
  </tr>
</table>

<%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage é maior que o número de páginas então intPage é igual ao número de páginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage é menor ou igual a zero então intPage igual a "1"
	'a variável intPage sempre vai ser forçada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados então.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a página exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a variável intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posição exata do primeiro registro da página correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage é igual ao número de páginas no recordset , estamos na última 
			'página então.
				intFinish = intRecordCount
				'a variável intFinish recebe o valor do número do último recordset.
				'intFinish corresponde ao valor do último registro da página correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a variável intFinish recebe o valor de intStart + o valor
				'do número de registros na página menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros então
		For intRecord = 1 to RS.PageSize
		'um contador inRecord é colocado até o número de registros na página.
%>
<% varCodimovel = rs("COD_Imovel") %>

<table width="591" height="190" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="591"><table width="580" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td bgcolor="#e9dca8"> 
            <table width="570" height="170" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td><table width="570" height="170" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="210" height="170"><table width="210" border="0" cellspacing="0" cellpadding="0">
                                      <tr>
                                        <td width="210" height="128" style="border:2px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rs("foto_pequena"))) = True Then%><img src="<%=rs("foto_pequena")%>" width="206" height="124" border="0"></img><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><strong><a href="javascript:newWindow3('mostrar_imovel_comprador2.asp?varCodimovel=<%=varCodimovel%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>&origem=<%="Site"%>')" style="color:#9d9249;text-decoration:none;">Foto não disponível</a></strong></font></div><%end if%></td>
                                      </tr>
                                      <tr>
                                        <td width="210" height="42"><table width="210" height="42" border="0" cellpadding="0" cellspacing="0">
                                              <tr>
                                              <td bgcolor="#f7ecbf"><div align="center"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Meu 
                                      im&oacute;vel </strong></font></div></td>
                                            </tr>
                                          </table></td>
                                      </tr>
                                    </table></td>
                      <td><table width="355" height="170" border="0" align="right" cellpadding="0" cellspacing="0">
                          <tr>
                            <td bgcolor="#f7ecbf"><div align="center">
                                <table width="355" border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td><table width="355" border="2" bordercolor="#f7ecbf" cellspacing="0" cellpadding="0">
                                        <tr bgcolor="#e9dca8"> 
                                          <td width="118" height="20"> 
                                            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><strong><%=rs("cidade")%></strong></font></div></td>
                                          <td width="118" height="20"> 
                                            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><strong><%=rs("bairro")%></strong></font></div></td>
                                          <td width="118" height="20"> 
                                            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><strong><%=rs("area_construida")%>m²</strong></font></div></td>
                                        </tr>
                                      </table></td>
                                  </tr>
                                  <tr>
                                    <td height="20" bgcolor="#e9dca8" style="border:1px solid #f7ecbf;"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><strong><%=formatnumber(rs("valor"),2)%></strong></font></div></td>
                                  </tr>
                                  <tr>
                                    <td height="130"><div align="center"><font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("obs_imovel")%><br>
                                        <br>
                                        Atualizado em:<%=rs("data_atualizacao")%><br>
                                        <br>
                                        <br>
                                        Referência:<strong><%=rs("cod_imovel")%></strong></font></div></td>
                                  </tr>
                                </table>
                                <font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>






<%
RS.MoveNext


	  





 'acima é feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next



RS.close
Set RS = Nothing

	
%></tr>
</table> 
<table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&varCodCompradores=<%=varCodCompradores%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" size="1" > 
            </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%> 
            <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
            <a href="?page=<%=intPage + 1%>&varCodCompradores=<%=varCodCompradores%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>"><b><font color="#000000" face="Verdana, arial" size="1">Próximo</font></b></a><a href=""> 
            </a> 
            <%End If%>
            </font></div></td>
        </tr>
      </table>


<%End If


Else

%>
  <% 
response.write "<html><body bgcolor='EAA813'><br><br><br><center><font size='1' face='Verdana, Arial, Helvetica, sans-serif'><strong></strong></font></center></body></html>"

%>


<br>
<% end if %>
<%

set objFSO = nothing



  
 




%>
<%
           rs9.Close
           'fecha a conexão
           
           Set rs9 = Nothing
		   
		  rsVerifica2.close
		   
		   set rsVerifica2 = nothing
		   
		   
		   Conexao9.Close
		   set conexao9 = nothing
           %>
<% response.flush%>
<%response.clear%>
</body>
</html>

