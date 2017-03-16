<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<!--#include file="style_imoveis.asp"-->
<%response.Buffer = true %>

<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "S?o Bernardo"
end if

'--------------------------Fazer conex?o-------------------------

 dim SqlImoveis001
 dim rsImoveis001



'Criando conex?o com o banco de dados! 
Set Conexao3 = Server.CreateObject("ADODB.Connection")
Conexao3.Open dsn



'---------------Pegar dados do cliente-----------------



session("nome") = request.form("txt_nome")

if session("nome") = "" then


session("nome") = request.querystring("nome")


end if





session("telefone") = request.form("txt_telefone")

if session("telefone") = "" then


session("telefone") = request.querystring("telefone")


end if





session("email") = request.form("txt_email")

if session("email") = "" then


session("email") = request.querystring("email")


end if



'-------------------------------------------------------------------





'-----------listagem de cidades-----------------------
'Criando conex?o com o banco de dados! 


'Abrindo a tabela MARCAS!
Sql33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



Set rs33 = Server.CreateObject("ADODB.RecordSet")

	rs33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs33.CursorType = 3
'indica o tipo de cursor utiliz?o

rs33.ActiveConnection = Conexao3


rs33.Open Sql33, Conexao3






'-----------------------------------------------------------
dim vCidade2
 dim vCidade
   
   if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")
end if
   
   
    vCidade2=request.form("combo3")
	
	
	
	session("vCidade2") = vCidade2
	 if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")

end if





'---------------------listagem de bairros-----------------------




dim rs44,strSQL44,Conexaoo
   
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	
	
	strSQL44 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4  ORDER BY nome_combo2" 
	
	
	if session("vCidade2") <> "cqualquer" then
	
	strSQL44 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 ="&int(session("vCidade2"))&"  ORDER BY nome_combo2" 
	else
	strSQL44 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	end if
	
	
	

	rs44.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs44.CursorType = 3
'indica o tipo de cursor utiliz?o

rs44.ActiveConnection = Conexao3


	
	
	
	
	
	
	
	rs44.Open strSQL44, Conexao3






'----------------------------------------------------------------





'-----------------------------Listagem de tipos----------------------


 dim rs444Tipo23,strSQL444Tipo23
   
    Set rs444Tipo23 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo23 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo FROM tipo  ORDER BY tipo ASC" 
	
	
	
	
	rs444Tipo23.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444Tipo23.CursorType = 3
'indica o tipo de cursor utiliz?o

rs444Tipo23.ActiveConnection = Conexao3


rs444Tipo23.Open strSQL444Tipo23, Conexao3





'---------------------------------------------------------------------





'----------------------------Buscar informa??es para o SQL---------------



'------------------Buscar informa??es de cidade---------------------

 
	  
	
	if session("vCidade2") <> "cqualquer" and session("vCidade2") <> "" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1 ="&session("vCidade2")
 
 rs2.open SQL2,Conexao3,2,1
 
 vCidade = rs2("nome_combo1")
 
 rs2.close
 
 set rs2 = nothing
 
 else
 vCidade = vCidade2
 end if

	session("vCidade")= vCidade







'-------------------------------------------------------------------


'------------------------pegar os dados dos bairros---------------


dim vBairro2
dim vBairro
	 vBairro2=request.form("combo4")
	 session("vBairro2") = vBairro2
	 if session("vBairro2") = "" then
session("vBairro2") = request.querystring("vBairro2")

end if
	 
	 if session("vBairro2") <> "bqualquer" and session("vBairro2") <> ""  then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="& session("vBairro2")
 
 rs3.open SQL3,Conexao3,2,1

 vBairro = rs3("nome_combo2")
 
 rs3.close
 
 set rs3 = nothing
 
 else
 vBairro = vBairro2
	end if                                      
									
	 
	 
	 
	 session("vBairro")= vBairro






'-------------------------------------------------------------------





'----------------------pegar dados de tipo----------------------------


dim vTipo



 vTipo=request.form("txt_tipo")
  
  
  
  if vTipo = "" then
  
  vTipo = request.querystring("vTipo")
  
  end if
  
  
   session("vTipo") = vTipo



'-----------------------------Pegar vagas----------------------------

dim vVagas



 vVagas=request.form("txt_vagas")
 
 if vVagas = "" then
 
 vVagas = request.querystring("vVagas")
 
 end if
 
 
  session("vVagas")=vVagas




'---------------------------Quartos------------------------------------

dim vQuartos

 vQuartos=request.form("txt_quartos")
 
 if vQuartos = "" then
 
 vQuartos = request.querystring("vQuartos")
 
 end if
 
 
  session("vQuartos")=vQuartos



'------------------------Negocia??o------------------------------

dim vNegociacao

 vNegociacao=request.form("example22")


if vNegociacao = "" then

vNegociacao = request.querystring("vNegociacao")

end if


if vNegociacao = "compra" then
'vNegociacao = "venda"
end if

 session("vnegociacao") = vNegociacao






'----------------------pegar valor-----------------------------

dim vValor

 vValor=request.form("stage222")
   
   if vValor = "" then
   vValor = request.querystring("vValor")
   end if
   
   if vValor = "" then

vValor = "0"
end if
   
   
   session("vValor")=vValor
   session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)


dim vValorMedio

vValorMedio = ((int(session("vValor1"))+int(session("vValor2")))/2)







'----------------------------------------------------------------------






'--------------------------------------------------------------------------










'------------------montar a string---------------------------------


'-------------------------Cidade-----------------------------------
 dim stringCidade
 dim stringIndex
 stringIndex = " where cod_compradores <>"&"0"&"" 

if  session("vCidade") <> "cqualquer" and  session("vCidade") <> "" then
stringCidade = " and (cidade='"& session("vCidade")&"' or cidade='"& "n?o informado"&"') "
else
stringCidade = ""
end if
'-----------------------------------------------------------------------


'------------------------------------Bairro-----------------------------

dim stringBairro
 
  if session("vBairro") <> "bqualquer" and  session("vBairro") <> "" then
 
  stringBairro = " and (bairro like '%"&session("vBairro")&"%' or bairro like '%"&"n?o informado"&"%') "
  
 
 
  
  else
 
  stringBairro = ""
  
  end if





'------------------------------------------------------------------




'-------------------------Tipo---------------------------------------

'--------------------------------------Tipo-------------------------------


	dim stringTipo
 
  if session("vTipo")<>"tqualquer" and  session("vTipo")<>""  then
  stringTipo = " and Tipo like'%"&session("vTipo")&"%'"
  else
  stringTipo = ""
  end if
 






'-------------------------------------------------------------------



'-------------------------------Negociacao-------------------------------
	
	
	
	dim stringNegociacao
	
	if session("vNegociacao") = "venda" then
	session("vNegociacao") = "compra"
	end if
	
 
  if session("vNegociacao")<>"nqualquer" and session("vNegociacao")<>"" then
  stringNegociacao = " and Negociacao ='"&session("vNegociacao")&"'"
  else  
  stringNegociacao = ""
  end if
  	
	
	
	'-------------------------------------------------------------------
	'---------------------------Quartos------------------------------


if  session("vQuartos") <> "qqualquer" and session("vQuartos") <> "" then
stringQuartos = " and quartos <="&int(session("vQuartos"))&""
else
stringQuartos = ""
end if

'---------------------------------------------------------------------------

'---------------------------Vagas------------------------------


if  session("vVagas") <> "vgqualquer" and session("vVagas") <> "" then
stringVagas = " and vagas <="&int(session("vVagas"))&""
else

stringVagas = ""
end if





	
	
	 '----------------------------------Valor--------------------------------
	 
	
	 
	 dim stringValor
 
  if session("vValor")<>"vqualquer" and  session("vValor")<>"" then
  stringValor = " and Valor >="& session("vValor1") &" and Valor <= "& session("vValor2") &""
  else
  stringValor = "" 
  end if
   
	 
	 '----------------------------------------------------------------------
	 
	 
'------------------------------Pegar os dois im?veis em destaque------------



dim rsFrontPage,SQLFrontPage,objFSO 

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

Set rsFrontPage = Server.CreateObject("ADODB.RecordSet")

SQLFrontPage = "SELECT * FROM imoveis where presenca_primeira like '"&"incluido"&"' ORDER BY cod_imovel DESC"

rsFrontPage.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rsFrontPage.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rsFrontPage.ActiveConnection = Conexao3


rsFrontPage.open SQLFrontPage,Conexao3

dim intRecordCount2


intRecordCount2 = rsFrontPage.RecordCount



dim EnderecoIP , vData
  vData = now()
  
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")

 if  vCidade2 <> ""  then
 
  Conexao3.execute"Insert into compradores_procurados(nome,telefone,Cidade, bairro ,tipo,negociacao,valor,enderecoIP,data,quartos,vagas,origem_franquia,telefone_quem_procurou) values( '"& session("nome") &"','"& session("telefone") &"','"& session("vCidade") &"','"& session("vBairro") &"','"& session("vTipo") &"','"& session("vNegociacao") &"','"& session("vValor") &"','"& EnderecoIP &"','"& vData &"','"& session("vQuartos") &"','"& session("vVagas") &"','"& session("vOrigem_Franquia") &"','"& session("telefone") &"')"
  
  end if





%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Busca de compradores</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=605,height=530,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>




<script>

// Verifica se somente n?meros foram digitados no campo
function isValidDigitNumber2 (form) 



{




{

if (form.combo3.value == "cqualquer") {
		alert("Por favor, informe a cidade onde se localiza o seu imóvel.");
		form.combo3.focus();
		
		return false;
}


if (form.combo4.value == "bqualquer") {
		alert("Por favor, informe o bairro onde se localiza o seu imóvel.");
		form.combo4.focus();
		
		return false;
}


if (form.txt_tipo.value == "tqualquer") {
		alert("Por favor, informe o tipo do seu imóvel.");
		form.txt_tipo.focus();
		
		return false;
}


if (form.txt_quartos.value == "qqualquer") {
		alert("Por favor, informe quantos quartos tem o seu imóvel.");
		form.txt_quartos.focus();
		
		return false;
}


if (form.txt_vagas.value == "vgqualquer") {
		alert("Por favor, informe quantas vagas na garagem tem o seu imóvel.");
		form.txt_vagas.focus();
		
		return false;
}


if (form.example22.value == "nqualquer") {
		alert("Por favor, escolha uma negocia??o.");
		form.example22.focus();
		
		return false;
}


if (form.stage222.value == "vqualquer") {
		alert("Por favor, escolha um valor.");
		form.stage222.focus();
		
		return false;
}

}
}


</script>



</head>

<body topmargin="0" bottommargin="0" leftmargin="0" rightmargin="0">

<form name="form"  method="post" onSubmit="return isValidDigitNumber2(this);" action="procurar_compradores001.asp">
<table width="794" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="794" height="106"><img src="top01.jpg" width="794" height="106"></td>
  </tr>
  <tr>
      <td height="60" bgcolor="#e6dca9" > 
        <div align="center">
          <div align="center"><font color="#e0a94e" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000" size="3">Esta 
            é a página para você encontrar um comprador ou inquilino para o seu 
            imóvel</font></strong></font><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font></div>
          <font color="#9d9249" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td> 
  </tr>
  <tr>
    <td height="260"><table width="794" height="260" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="390" height="260" bgcolor="#e0a94e"><table width="380" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="380" height="250" bgcolor="#e6dca9"><table width="370" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td bgcolor="#e0a94e"><table width="360" height="230" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td bgcolor="#e6dca9"><table width="356" height="226" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                        <input name="txt_nome" type="text" class="inputBox"  id="txt_nome" style="HEIGHT: 18px; WIDTH: 350px; font-size : 12px;  background: #b2802c; color:#FFFFFF; font-size:12px;" onfocus="form.txt_nome.value=''" value="<% if session("nome") <> "" then response.write session("nome") else response.write "Seu nome:" end if%>" size="30" maxlength="30">
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                        <input name="txt_telefone" type="text" class="inputBox"  id="txt_telefone" style="HEIGHT: 18px; WIDTH: 350px; font-size : 12px;  background: #b2802c; color:#FFFFFF;" onfocus="form.txt_telefone.value=''" value="<% if session("telefone") <> "" then response.write session("telefone") else response.write "Seu telefone:" end if%>" size="30" maxlength="30">
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                        <input name="txt_email" type="text" class="inputBox"  id="txt_email" style="HEIGHT: 18px; WIDTH: 350px; font-size : 12px;  background: #b2802c; color:#FFFFFF;" onfocus="form.txt_email.value=''" value="<% if session("email") <> "" then response.write session("email") else response.write "Seu email:" end if%>" size="30" maxlength="30">
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                      <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 350px; font-size : 12px;  background: #b2802c; color:#FFFFFF;" onChange="javascript:atualizacarros2(this.form);">
                                        <option value="cqualquer" >Em qual cidade Esta localizado seu imóvel?</option>
          
		    <% if not rs33.eof then %>
            <% While NOT Rs33.EoF %>
			
             <option value="<% = Rs33("id_combo1") %>"<%if session("vCidade2")<> "cqualquer" then%><%if int(rs33("id_combo1")) = int(session("vCidade2")) then response.write "selected" else response.write "" end if %><%end if%>> 
		   
            <% = Rs33("nome_combo1") %>
            </option>
            <% Rs33.MoveNext %>
            <% Wend %>
            <option value="cqualquer">qualquer uma</option>
            <%else%>
            <option value=""></option>
            <%end if%>
          </select>
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                      <select name="combo4"  class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 350px; font-size : 12px;  background: #b2802c; color:#FFFFFF;">
                                       <option value="bqualquer" >Em qual bairro Esta localizado seu imóvel?</option>
            <% if not rs44.eof then%>
            <% While NOT Rs44.EoF %>
             <option value="<% = Rs44("id_combo2") %>" <% if session("vBairro2") <> "bqualquer" then if int(Rs44("id_combo2")) <> int(session("vBairro2"))  then response.write "" else response.write "selected" end if end if %>> 
                
            
			
			<% = Rs44("nome_combo2") %>
            </option>
            <% Rs44.MoveNext %>
            <% Wend %>
            <option value="bqualquer">qualquer um</option>
            <% else %>
            <option value=""></option>
            <% end if %>
          </select>
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                      <select name="txt_tipo" size="1"  class="inputBox" id="txt_tipo" style="HEIGHT: 18px; WIDTH: 350px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                      
									  <option value="<%if session("vTipo") <> "tqualquer" and session("vTipo") <> "" then  response.write session("vTipo") else response.write "tqualquer" end if%>" selected><%if session("vTipo") <> "tqualquer" and session("vTipo") <> "" then  response.write session("vTipo") else response.write "Qual o tipo de imóvel que o sr(a) tem ?" end if%></option>
				 
									   
                  <option value="tqualquer">Qualquer um</option>
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
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                      <select name="txt_quartos" size="1"  class="inputBox" id="txt_quartos" style="HEIGHT: 18px; WIDTH: 350px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                        
										 <option value="<% if session("vQuartos") <> "qqualquer" and session("vQuartos") <> "" then response.write session("vQuartos") else response.write "qqualquer" end if%>"><% if session("vQuartos") <> "qqualquer" and session("vQuartos") <> "" then response.write session("vQuartos") else response.write "Quantos quartos tem o seu imóvel?" end if%></option>
										 
                  <option value="qqualquer">Qualquer um</option>
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
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center">
                                      <select name="txt_vagas" size="1"  class="inputBox" id="txt_vagas"  style="HEIGHT: 18px; WIDTH: 350px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                       
									   <option value="<% if session("vVagas") <> "vgqualquer" and session("vVagas") <> "" then response.write session("vVagas") else response.write "vgqualquer" end if%>"><% if session("vVagas") <> "vgqualquer" and session("vVagas") <> "" then response.write session("vVagas") else response.write "Quantas vagas na garagem tem o seu imóvel?" end if%></option>
									   
									    
                  <option value="vgqualquer">Qualquer um</option>
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
                                    </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center"> 
                                        <select name="example22" size="1" class="inputBox" id="example22" style="HEIGHT: 18px; WIDTH: 350px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;" onChange="redirect3(this.options.selectedIndex)">
                                          <option value="<% if session("vnegociacao") <> "nqualquer" and session("vnegociacao") <> ""  then response.write session("vnegociacao") else response.write "nqualquer" end if%>" selected>
                                          <% if session("vnegociacao") <> "nqualquer" and session("vnegociacao") <> ""    then 
				
				if session("vnegociacao") = "aluguel" then 
				response.write "Aluguel"
				end if
				
				if session("vnegociacao") = "compra" then 
				response.write "Vender"
				end if
				
				if session("vnegociacao") = "venda" then 
				response.write "Vender"
				end if
				
				 else 
				 
				 response.write "O que o sr(a) quer fazer com seu imóvel?" end if%>
                                          </option>
                                          <option  value="nqualquer">Qualquer 
                                          um</option>
                                          <option  value="aluguel">Alugar </option>
                                          <option value="venda">Vender</option>
                                        </select>
                                      </div></td>
                                </tr>
                                <tr> 
                                  <td width="356" height="20"> 
                                    <div align="center"> 
                                        <select name="stage222" size="1" class="inputBox" id="stage222"  style="HEIGHT: 18px; WIDTH: 350px; ; font-size : 12px; background: #b2802c; color:#FFFFFF;">
                                          <option value="vqualquer" selected>Qual 
                                          a faixa de valor que o sr(a) pretende 
                                          trabalhar ?</option>
                                          <option value="vqualquer">Qualquer um</option>
                                          <% if session("vnegociacao") <> "venda" and session("vnegociacao") <> "compra" and  session("vnegociacao") <> "" then %>
                                         
										  <option value="<%=session("vValor")%>" <% if session("vValor") <> "" and  session("vValor") <> "vqualquer" then response.write "selected" else response.write "" end if%>>
                                          <% if session("vValor") <> "vqualquer" and session("vValor") <> ""  then response.write FormatNumber(session("vValor1"),2)&" at? "&FormatNumber(session("vValor2"),2) else response.write "Qual a faixa de valor que o sr(a) pretende trabalhar ?" end if%>
                                          </option>
                                          <option value="0000000000 0000000200">At? 
                                          200,00</option>
                                          <option value="0000000201 0000000500">201,00 
                                          at? 500,00</option>
                                          <option value="0000000501 0000000750">501,00 
                                          at? 750,00</option>
                                          <option value="0000000751 0000001000">751,00 
                                          at? 1000,00</option>
                                          <option value="0000001001 0000001500">1001,00 
                                          at? 1500,00</option>
                                          <option value="0000001501 0000002000">1501,00 
                                          at? 2000,00</option>
                                          <option value="0000002001 0000002500">2001,00 
                                          at? 2500,00</option>
                                          <option value="0000002501 0000003000">2501,00 
                                          at? 3000,00</option>
                                          <option value="0000003001 0000003500">3001,00 
                                          at? 3500,00</option>
                                          <option value="0000003501 0000004000">3501,00 
                                          at? 4000,00</option>
                                          <option value="0000004001 1000000000">Acima 
                                          de 4000,00</option>
                                          <%else%>
                                          
                                         <option value="vqualquer" selected>Qual o valor que o sr(a) pretende trabalhar?</option>
                                                       <option value="<%=session("vValor")%>" <%if session("vValor") <> "vqualquer" and session("vValor") <> "" then response.write "selected" else response.write "" end if %> >
                                          <% if session("vValor") <> "vqualquer" and session("vValor") <> "" then response.write FormatNumber(session("vValor1"),2)&" at? "&FormatNumber(session("vValor2"),2) else response.write "Qual faixa de valor o sr(a) pretende trabalhar?" end if%>
                                          </option>
													    <option value="vqualquer">Qualquer 
                                                        um</option>
                                                        <option value="0000000000 0000020000">At? 
                                                        20.000,00</option>
                                                        <option value="0000020001 0000050000">20.001,00 
                                                        at? 50.000,00</option>
                                                        <option value="0000050001 0000080000">50.001,00 
                                                        at? 80.000,00</option>
                                                        <option value="0000080001 0000110000">80.001,00 
                                                        at? 110.000,00</option>
                                                        <option value="0000110001 0000150000">110.001,00 
                                                        at? 150.000,00</option>
                                                        <option value="0000150001 0000200000">150.001,00 
                                                        at? 200.000,00</option>
                                                        <option value="0000200001 0000250000">200.001,00 
                                                        at? 250.000,00</option>
                                                        <option value="0000250001 0000300000">250.001,00 
                                                        at? 300.000,00</option>
                                                        <option value="0000300001 0000350000">300.001,00 
                                                        at? 350.000,00</option>
                                                        <option value="0000350001 0000400000">350.001,00 
                                                        at? 400.000,00</option>
														
														<option value="0000400001 0000600000">400.001,00 
                                                        at? 600.000,00</option>
														<option value="0000600001 0000800000">600.001,00 
                                                        at? 800.000,00</option>
														
														<option value="0000800001 0001000000">800.001,00 
                                                        at? 1000.000,00</option>
														
                                                        <option value="0001000001 1000000000">Acima 
                                                        de 1000.000,00</option>
                                          <%end if%>
                                        </select>
                                      </div></td>
                                </tr>
                                <tr> 
                                    <td width="356" height="20"> <div align="right">
                                        <input name="image" type="image" src="bt_procurar404.jpg" width="201" height="18">
                                      </div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          <td width="404" height="260" bgcolor="#e9dca8"><table width="404" height="260" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td width="10" height="260">&nbsp;</td>
                <td width="187" height="260"><table width="177" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td style="border:1px solid #FFFFFF;">
					  
					  <%if not rsFrontPage.eof and  rsFrontPage.RecordCount >= 2 then %>
					  
					  
					  <table width="167" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="167" height="240" bgcolor="#e0a94e" ><table width="157" height="230" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td width="157" height="30" bgcolor="#e6dca9"><div align="center"><font color="#ba9142" size="2" face="Verdana, Arial, Helvetica, sans-serif">Destaque</font></div></td>
                                </tr>
                                <tr>
                                  <td width="157" height="5"></td>
                                </tr>
                                <tr>
                                  <td width="157" bgcolor="#e6dca9" height="100" style="border:1px solid #FFFFFF;"><% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="155" height="98" border="0"></img></a><%else%><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto n?o dispon?vel</strong></a></font></div><%end if%></td>
                                </tr>
                                <tr>
                                  <td width="157" height="5"></td>
                                </tr>
                                <tr>
                                  <td width="157" height="90"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#FFFFFF"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table>
						
						
						</td>
                    </tr>
                  </table></td>
                <td width="10" height="260">&nbsp;</td>
                <td width="187" height="260"><table width="177" height="250" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
					<%=rsFrontPage.movenext%>
					
					
					<%else%>
						
						<%end if%>
                      <td style="border:1px solid #FFFFFF;">
					  
					  <% if not rsFrontPage.eof and  rsFrontPage.RecordCount >= 2 then%>
					  
					  
					  <table width="167" height="240" border="0" align="center" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="167" height="240" bgcolor="#e0a94e" ><table width="157" height="230" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr>
                                  <td width="157" height="30" bgcolor="#e6dca9"><div align="center"><font color="#ba9142" size="2" face="Verdana, Arial, Helvetica, sans-serif">Destaque</font></div></td>
                                </tr>
                                <tr>
                                  <td width="157" height="5"></td>
                                </tr>
                                <tr>
                                    <td width="157" height="100" bgcolor="#e6dca9" style="border:1px solid #FFFFFF;">
                                      <% If objFSO.FileExists(Server.MapPath(rsFrontPage("foto_pequena"))) = True Then%>
                                      <a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')"><img src="<%=rsFrontPage("foto_pequena")%>" width="155" height="98" border="0"></img></a>
                                      <%else%>
                                      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#9d9249"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#9d9249"><strong>Foto 
                                        n?o dispon?vel</strong></a></font></div>
                                      <%end if%></td>
                                </tr>
                                <tr>
                                  <td width="157" height="5"></td>
                                </tr>
                                <tr>
                                  <td width="157" height="90"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('mostrar_imovel2.asp?varCodimovel=<%=rsFrontPage("cod_imovel")%>')" style="text-decoration:none;color:#FFFFFF"><%=rsFrontPage("texto_anuncio")%></a></font></div></td>
                                </tr>
                              </table></td>
                          </tr>
                        </table>
						
						<%else%>
						
						
						<%end if%>
						</td>
                    </tr>
                  </table></td>
                <td width="10" height="260">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  
  
   <%
 
 
 

dim rs444VerificaConta,strSQL444VerificaConta
   
    Set rs444VerificaConta = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where (telefone like '"&session("telefone")&"' or telefone02 like '"&session("telefone")&"' or telefone03 like '"&session("telefone")&"') and atendimento <>'"&"internet"&"' and atendimento <>'"&"n?o informado"&"' " 
	
	
	
	rs444VerificaConta.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rs444VerificaConta.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rs444VerificaConta.ActiveConnection = Conexao3
	
	
	
	
	
	 rs444VerificaConta.Open strSQL444VerificaConta, Conexao3
	

if  not rs444VerificaConta.eof and session("telefone") <> "" then






vCadastrado = "sim"

%>
 
 
  
  <tr>
  <td>
  
  <table width="794" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="200" bgcolor="#e0a94e"> 
        <div align="center">
          <table width="785" height="190" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td bgcolor="#e6dca9"><table width="600" height="150" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
    <td width="600"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font size="5"><a href="acessoLink02.asp?varTelefone=<%=session("telefone")%>" style="color:red;text-decoration:none;"  target="_blank">Ol&aacute; 
                              sr(a) <%=session("nome")%></a></font><a href="acessoLink03.asp?varTelefone=<%=session("telefone")%>" style="color:red;text-decoration:none;" target="_blank"><br>
                              Obrigado por retornar ao nosso site, voc&ecirc; est&aacute; conosco desde o dia <%=rs444VerificaConta("data")%> 
                              o seu atendente &eacute; o sr(a) <%=rs444VerificaConta("atendimento")%>,querendo que algum dos compradores listados visite seu im&oacute;vel, 
                              procure seu atendente ou agente uma visita na ficha do comprador escolhido</a></strong></font></div></td>
  </tr>
</table>
 </td>
            </tr>
          </table>
		  </div>
</table>
<tr><td height="20"></td></tr>
</td>
</tr>
  
  <%else%>
  
  
  <%end if%>
  
  
  
  
 
 <%
 dim strSQL
 
 
 
 if session("vCidade") <> "" then
 strSQL = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores "&stringIndex&stringCidade&stringBairro&stringVila&stringTipo&stringNegociacao&stringQuartos&stringVagas&stringValor&" and atendimento <> 'internet' and ( standby like 'comprador OK') ORDER  BY Cod_compradores DESC"
 
 else
 
 strSQL = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where cod_compradores =0"
 
 end if
 
 dim RS
 
Set RS = Server.CreateObject("ADODB.Recordset")
'um objeto recordset ? inst?nciado.

Dim LinkTemp
'essa vari?vel vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = "#537497"
color2 = "#94ADC8"
'as vari?veis acima s?o usadas para trocar a cor das tabelas que conter?o os valores
'dos recordsets.






dim intPage
'essa vari?vel vai receber um valor inicial "1" que mostra que estamos na primeira p?gina.

dim intPageCount
'Essa vari?vel vai receber o valor da quantidade de p?ginas do recordset.

dim intRecordCount
'Essa vari?vel vai receber o n?mero de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a vari?vel intPage recebe o valor "1" na primeira p?gina.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conex?o o recordset utilizar?.

RS.MaxRecords = 50

	
RS.Open strSQL, Conn, 1, 3
'o recordset ? aberto
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por p?gina.

RS.CacheSize = RS.PageSize
'o Cache tamb?m conter? 20 registros por p?gina.

intPageCount = RS.PageCount
'A vari?vel intPageCount recebe o valor do n?mero de p?gina do recordset retornado.

intRecordCount = RS.RecordCount
'A vari?vel intRecordCount recebe o valor do n?mero de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
 
 %>
 
 
  
  <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount end if
'se intPage ? maior que o n?mero de p?ginas ent?o intPage ? igual ao n?mero de p?ginas.

	If CInt(intPage) <= 0 Then intPage = 1 end if
	'se intPage ? menor ou igual a zero ent?o intPage igual a "1"
	'a vari?vel intPage sempre vai ser for?ada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados ent?o.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a p?gina exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a vari?vel intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posi??o exata do primeiro registro da p?gina correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage ? igual ao n?mero de p?ginas no recordset , estamos na ?ltima 
			'p?gina ent?o.
				intFinish = intRecordCount
				'a vari?vel intFinish recebe o valor do n?mero do ?ltimo recordset.
				'intFinish corresponde ao valor do ?ltimo registro da p?gina correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a vari?vel intFinish recebe o valor de intStart + o valor
				'do n?mero de registros na p?gina menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros ent?o
		For intRecord = 1 to RS.PageSize
		'um contador inRecord ? colocado at? o n?mero de registros na p?gina.
%>
  
  
  
  <%
  varCodCompradores = rs("cod_compradores")
  
  %>
  <tr>
    <td width="794" height="190"><table width="784" height="180" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td style="border:1px solid #ddddc5;"><table width="774" height="170" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td bgcolor="#e9dca8"><table width="774" height="170" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td height="20"><div align="center"><font color="#996600" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Procuro 
                            por : <%=rs("cidade")%> ,<%=rs("bairro")%> 
                            , <%=rs("tipo")%> com<%=rs("vagas")%> vagas na garagem 
                            e <%=rs("quartos")%> dormit?rios no valor de <%=FormatNumber(rs("valor"),2)%></strong></font></div></td>
                      </tr>
                      <tr>
                        <td height="150"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('visualizar_comprador01.asp?varCodCompradores=<%=varCodCompradores%>')" style="color:#996600;text-decoration:none;"><%=rs("descricao")%> 
                            <br>
                      <strong>Se quiser saber mais clique aqui.</strong></a></font></div>
					  <br>
					  <%
					 
					
					 
					  SqlImoveis001 = "SELECT imoveis.telefone,imoveis.telefone02,imoveis.telefone03,imoveis.cod_imovel,imoveis.tipo,imoveis.valor FROM imoveis where telefone like '"&rs("telefone")&"' or telefone02 like '"&rs("telefone")&"' or telefone03 like '"&rs("telefone")&"' ORDER BY cod_imovel Desc" 

Set rsImoveis001 = Server.CreateObject("ADODB.RecordSet")

	rsImoveis001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rsImoveis001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava??o.

rsImoveis001.ActiveConnection = Conexao3
	
	
	rsImoveis001.Open sqlImoveis001, Conexao3
					  
		if not rsImoveis001.eof then 
		
		
		
		
		
		
					  %>
					  <br>
					  
					  
					  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FF0000">Aten&ccedil;&atilde;o 
                            este comprador tem <%=rsImoveis001("tipo")%> no valor de <%=formatnumber(rsImoveis001("valor"),2)%> reais para entrar 
                            no neg&oacute;cio,clique no texto acima para saber 
                            mais. </font><br>
                            </strong></font></div>
					  <%
					  end if
					  
					  rsImoveis001.close

                     set rsImoveis001 = nothing
					  
					  
					  %>
					  </td>
                      </tr>
                    </table></td>
              </tr>
            </table></td>
        </tr>
      </table> </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  
  <%
RS.MoveNext


	  





 'acima ? feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
  
  
  
  <tr>
    <td>&nbsp;</td>
  </tr>
  
</table>
<br>

<table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a p?gina atual for maior que "1" ent?o o link anteriro ? colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vValor=<%=session("vValor")%>&vTipo=<%=session("vTipo")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vNegociacao=<%=session("vNegociacao")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="text-decoration:none;"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          
        <td> 
          <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
            <strong> 
            <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
            <!-- se p?gina atual ? menor que o total de p?ginas e intPage maior que um
			  ou seja, se n?o estiver na primeira p?gina e nem na ?ltima ent?o. -->
            <font color="#000000">P?gina</font> <%=cInt(intPage)%> <font color="#000000">de</font> 
            <%=cInt(intPageCount)%> </strong></font> <strong> 
            <%End If%>
            </strong></font> </div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
            <%If cInt(intPage) < cInt(intPageCount)  Then%> 
            <!-- se intPage ? menor que o n?mero de p?ginas ent?o colocar o bot?o pr?ximo -->
            <a href="?page=<%=intPage + 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vVila=<%=session("vVila")%>&vVila2=<%=session("vVila2")%>&vValor=<%=session("vValor")%>&vTipo=<%=session("vTipo")%>&vQuartos=<%=session("vQuartos")%>&vVagas=<%=session("vVagas")%>&vNegociacao=<%=session("vNegociacao")%>&nome=<%=session("nome")%>&telefone=<%=session("telefone")%>&email=<%=session("email")%>" style="text-decoration:none;"><b><font color="#000000" face="Verdana, arial" size="1">Pr?ximo</font></b></a> 
            <%end if%> 
            
            </font></div></td>
        </tr>
      </table>

<% end if%>

<%else%>
  <table width="794" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="300" bgcolor="#e0a94e"> 
        <div align="center">
          <table width="785" height="290" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td bgcolor="#e6dca9"><div align="center"><font color="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                  <% if vCidade2 <> "" then %>
                  N?o foi encontrado nenhum comprador/inquilino para o seu imóvel. 
                  <%else%>
                  Esta é a página para você encontrar um comprador/inquilino para 
                  o seu imóvel. 
                  <%end if%>
                  </strong></font></div></td>
            </tr>
          </table>
        </div></td>
  </tr>
</table>


<% end if%>


</form>



<script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups22=document.form.example22.options.length
/* Aqui ? criada uma vari?vel "groups" que receber? os valores 
do combo example. */



var group22=new Array(groups22)
/* aqui a vari?vel group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups22; i2++)
/* aqui temos um contador de zero at? o n?mero de elementos do array "groups" */

group22[i2]=new Array()
/* aqui ? criado o array "group" que receber? valores conforme o n?mero de elementos
do array "groups". */

group22[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receber? valores de op??es. */


group22[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receber? valores de op??es. */

group22[2][0]=new Option("Qual a faixa de valor que o sr(a) pretende trabalhar ?","vqualquer")
group22[2][1]=new Option("Qualquer Valor","vqualquer")
group22[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group22[2][3]=new Option("201,00 at? 500,00","0000000201 0000000500")
group22[2][4]=new Option("501,00 at? 750,00","0000000501 0000000750")
group22[2][5]=new Option("751,00 at? 1000,00","0000000751 0000001000")
group22[2][6]=new Option("1001,00 at? 1500,00","0000001001 0000001500")
group22[2][7]=new Option("1501,00 at? 2000,00","0000001501 0000002000")
group22[2][8]=new Option("2001,00 at? 2500,00","0000002001 0000002500")
group22[2][9]=new Option("2501,00 at? 3000,00","0000002501 0000003000")
group22[2][10]=new Option("3001,00 at? 3500,00","0000003001 0000003500")
group22[2][11]=new Option("3501,00 at? 4000,00","0000003501 0000004000")
group22[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group22[3][0]=new Option("Qual a faixa de valor que sr(a) pretende trabalhar?","vqualquer")
group22[3][1]=new Option("Qualquer Valor","vqualquer")
group22[3][2]=new Option("At?  20.000,00","0000000000 0000020000")
group22[3][3]=new Option("20.001,00 at? 50.000,00","0000020001 0000050000")
group22[3][4]=new Option("50.001,00 at? 80.000,00","0000050001 0000080000")
group22[3][5]=new Option("80.001,00 at? 110.000,00","0000080001 0000110000")
group22[3][6]=new Option("110.001,00 at? 150.000,00","0000110001 0000150000")
group22[3][7]=new Option("150.001,00 at? 200.000,00","0000150001 0000200000")
group22[3][8]=new Option("200.001,00 at? 250.000,00","0000200001 0000250000")
group22[3][9]=new Option("250.001,00 at? 300.000,00","0000250001 0000300000")
group22[3][10]=new Option("300.001,00 at? 350.000,00","0000300001 0000350000")
group22[3][11]=new Option("350.001,00 at? 400.000,00","0000350001 0000400000")
group22[3][12]=new Option("400.001,00 at? 600.000,00","0000400001 0000600000")
group22[3][13]=new Option("600.001,00 at? 800.000,00","0000600001 0000800000")
group22[3][14]=new Option("800.001,00 at? 1000.000,00","0000800001 0001000000")
group22[3][15]=new Option("Acima de 1000.000,00","0001000001 1000000000")








/* aqui temos um array bidimensional "group" que receber? valores de op??es. */


var temp22=document.form.stage222
/* aqui a vari?vel "temp" recebe os valores do segundo combo o "stage2" */

function redirect3(x2){
/* aqui ? criada a fun??o "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp22.options.length-1;m2>0;m2--)
temp22.options[m2]=null
/* aqui temos um contador "m" que d? um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group22[x2].length;i2++){
temp22.options[i2]=new Option(group22[x2][i2].text,group22[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que ? escolhido no
primeiro combo "example".*/

}
temp22.options[0].selected=true
}
/* aqui o array "temp.options[0]" ser? o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location22=temp22.options[temp22.selectedIndex].value
}

/* aqui  a vari?vel "location" recebe os valores de "stage2" que corresponde ao endere?o de
link para o carregamento de p?gina. */


//-->
</script>


<%
'-------------------Cadastrar busca-----------------------

  dim rs444Imovel,SQL444Imovel
 Set rs444Imovel = Server.CreateObject("ADODB.RecordSet")
 SQL444Imovel = "SELECT imoveis.cod_imovel,imoveis.telefone,imoveis.telefone02,imoveis.telefone03  FROM imoveis where (telefone like'"& session("telefone")&"' or telefone02 like'"& session("telefone")&"' or telefone03 like '"& session("telefone")&"') order by cod_imovel DESC" 
	
	
	rs444Imovel.CursorLocation = 3
         rs444Imovel.CursorType = 3
           rs444Imovel.ActiveConnection = Conexao3
	
	
	rs444Imovel.open SQL444Imovel,Conexao3,2,1  
	
			
	if   rs444Imovel.eof then




session("vValorMedio") = vValorMedio

dim vNegociacao002

if session("vNegociacao") = "compra" then

vNegociacao002 = "venda"

else
vNegociacao002 = "aluguel"
end if

Conexao3.execute"Insert into imoveis(proprietario,endereco,telefone,email,foto_grande,foto_pequena,foto_grande1,foto_grande2,foto_grande3,foto_grande4,foto_grande5,foto_grande6,foto_grande7,foto_grande8,foto_grande9,foto_grande10,link_foto,cidade,bairro,tipo,area_total,area_construida,quartos,banheiros,vagas,negociacao,valor,data,obs_imovel,obs_proprietario,presenca_primeira,titulo_anuncio,texto_anuncio,standby,ocupacao,captacao,data_atualizacao,vila,placa,condominio,cod_permuta,cod_comprador,qualidade,indexador_indicacoes,origem_captacao,data_captacao,cliques_no_imovel,tarja02,data01_tarja02,data02_tarja02,imovel_em_negociacao,data_contato,origem_franquia) values( '"& session("nome") &"','"& "n?o informado" &"','"& session("telefone") &"','"& session("email") &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "imovel00000.jpg" &"','"& "icon_foto2.gif" &"','"& vCidade &"','"& vBairro &"','"& vTipo &"','"& "0" &"','"& "0" &"','"& session("vQuartos") &"','"& "n?o informado" &"','"& session("vVagas")&"','"& vNegociacao002 &"','"& int(session("vValorMedio")) &"','"& now() &"','"& "n?o informado" &"','"& "n?o informado" &"','"& "excluido" &"','"& "n?o informado" &"','"& "n?o informado" &"','"& "excluido" &"','"& "n?o informado" &"','"& "internet" &"','"& now() &"','"& "n?o informado" &"','"& "Sem Placa"&"','"& "0" &"','"& "0" &"','"& "0" &"','"& "neg?cio comum" &"','"&"0"&"','"&"Busca de compradores"&"','"& now()&"','"& "0"&"','"& "sim"&"','"& day(now())&"','"& day(DateAdd("d", 15, now()))&"','"& "imóvel n?o contatado" &"','"& now() &"','"& session("vOrigem_Franquia") &"')"	 
	


end if
'------------------------------------------------------------------------
%>






<% response.flush%>
  <%response.clear%>


<%
Function EscreveFuncaoJavaScript2 ( Conexao3 )
'O parametro conexao receber? uma conexao aberta!
'Em funcoes, geralmente n?o criamos objetos do tipo conex?es!
'Opte por sempre deixar sua fun??o o mais compat?vel poss?vel com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (form) {" & vbcrlf

'Essa fun??o JavaScript recebe o form em que Estao os campos a serem atualizados!
'Veja na chamada da fun??o no m?todo OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual op??o foi selecionada!! 
Response.Write "switch (form.combo3.options[form.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as op??es de carro!
SqlMarcas33 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1 FROM combo1 ORDER BY nome_combo1" 



Set rsMarcas33 = Server.CreateObject("ADODB.RecordSet")

	rsMarcas33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rsMarcas33.CursorType = 3
'indica o tipo de cursor utiliz?o

rsMarcas33.ActiveConnection = Conexao3


rsMarcas33.Open SqlMarcas33, Conexao3




While NOT rsMarcas33.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas33("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros33 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 WHERE id_combo1 =" & rsMarcas33("id_combo1")&" order by nome_combo2"



Set rsCarros33 = Server.CreateObject("ADODB.RecordSet")

	rsCarros33.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor ? criado
'se no cliente ou no servidor.

rsCarros33.CursorType = 3
'indica o tipo de cursor utiliz?o

rsCarros33.ActiveConnection = Conexao3


rsCarros33.Open SqlCarros33, Conexao3





'Fazemos um loop por todos os carros, criando uma nova op??o no SELECT! 
 i = 0 
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "Qual o bairro do seu imóvel?" & "','" & "bqualquer" & "');"& vbcrlf
i = 1
While NOT rsCarros33.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend
Response.Write "form.combo4.options[" & i  & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" 
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma d?vida da sua utiliza??o! 
Response.Write "break;" & vbcrlf

'Pr?xima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da fun??o! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 



rsMarcas33.close

set rsMarcas33 = nothing

rsCarros33.close

set rsCarros33 = nothing

End Function
%> 



<%  EscreveFuncaoJavaScript2 ( Conexao3 ) %>

<%'strSQL%>

</body>
</html>
