<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->

<!--#include file="cores.asp"-->





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

%> 


<%


dim varCidade,stringCidade,varBairro,stringBairro,varNegociacao
dim stringNegociacao,varQuartos,stringQuartos,varCidade2

 varCidade2 = request.querystring("combo1")
 if varCidade2 = "" then
 varCidade2 = "cqualquer"
 end if
 
session("varCidade2") = varCidade2
 
 if varCidade2 <> "cqualquer" then
 dim rrs2,SSQL2




 Set rrs2 = Server.CreateObject("ADODB.RecordSet")
 SSQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1="&varCidade2
 
 
 


	rrs2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rrs2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rrs2.ActiveConnection = Conexao3
	
	

 
 
 
 rrs2.open SSQL2,Conexao3,2,1
 
 varCidade = rrs2("nome_combo1")
 
 '-----------------------------
           rrs2.Close           
		   
           Set rrs2 = Nothing
		   
'---------------------------------
 
 else
 varCidade = varCidade2
 end if
 
 if request.QueryString("varCidade")<>"" then
  varCidade = request.QueryString("varCidade")
 session("varCidade") = varCidade
 else
 
 session("varCidade") = varCidade
 end if
 
 
 
 
 
 
 
 
 
 
 
 
 
varBairro2 = request.querystring("combo2")
if varBairro2 = "" then
varBairro2 = "bqualquer"
end if

 if varBairro2 <> "bqualquer" then
	  dim rrs3,SSQL3,conexao5
	 
 Set rrs3 = Server.CreateObject("ADODB.RecordSet")
 SSQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="&varBairro2
 
 rrs3.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rrs3.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rrs3.ActiveConnection = Conexao3
 
 
 rrs3.open SSQL3,Conexao5,2,1

 varBairro = rrs3("nome_combo2")
 
 '-----------------------------
           rrs3.Close           
		   
           Set rrs3 = Nothing
		   
'---------------------------------
 
 
 else
 varBairro = varBairro2
	end if                                      
									
	  if request.QueryString("varBairro")<>"" then
  varBairro = request.QueryString("varBairro")
 session("varBairro") = varBairro
 else
 
 session("varBairro") = varBairro
 end if




varNegociacao = request.querystring("example2")
varQuartos = request.querystring("txt_quartos")





dim varValor, varValor1,varValor2
 

varValor = request.QueryString("stage22")

session("varValor")=varValor
  


if request.QueryString("stage22")<>"" then
varValor = request.QueryString("stage22")
 session("varValor") = varValor
 else
 
 session("varValor") = varValor
 end if
 
  if request.QueryString("varValor")<>"" then
  varValor = request.QueryString("varValor")
 session("varValor") = varValor
 else
 
 varValor = session("varValor") 
 end if 
	
 session("varValor1")=left(varValor,10)
   session("varValor2")=right(varValor,10)	
	
	
	varValor1 = session("varValor1")
	varValor2 = session("varValor2")
	
	 if request.QueryString("varNegociacao")<>"" then
varNegociacao = request.QueryString("varNegociacao")
 session("varNegociacao") = varNegociacao
 else
 
 session("varNegociacao") = varNegociacao
 end if
   
   
   
   
    if request.QueryString("varQuartos")<>"" then
 varQuartos = request.QueryString("varQuartos")
 session("varQuartos") = varQuartos
 else
 
 session("varQuartos") = varQuartos
 end if



dim varNotFind



dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
   
   
   conexao.open dsn
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
	rs4.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs4.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs4.ActiveConnection = Conexao
	
	
	
	
	
	
	rs4.Open strSQL4, Conexao

dim rs4Verifica 
dim strSQL4Verifica
dim permissao

permissao = "não"

Set rs4Verifica = Server.CreateObject("ADODB.RecordSet")



dim rs444,strSQL444
   
   





%>




<html>
<head>
<title>Ligar urgente</title>

<script>

function check(acao){
if(document.Formulario.selTodos.checked){
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked = acao;
}
}
else
{
e = document.Formulario.elements;
for(i=0;i<e.length;i++){
if(e[i].type == "checkbox") e[i].checked =! acao;
}
}



}





</script>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2020(abrejanela2020) {
   openWindow2020 = window.open(abrejanela2020,'openWin2020','width=800,height=600,resizable=yes,scrollbars=yes')
   openWindow2020.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=800,height=600,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3232(abrejanela3232) {
   openWindow3232 = window.open(abrejanela3232,'openWin3232','width=610,height=500,resizable=yes,scrollbars=yes')
   openWindow3232.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE="Javascript">
<!--

//showSubTopNav();
//showSubLeftNav(0, 1);

var popupVisible = false;

function show_info_popup(thisObj,menu_id,vertical_offset) {
	if (popupVisible == false) {
		menuObj = document.getElementById(menu_id);
		position = getAnchorPosition(thisObj.id);
		moveObject(menu_id,position.x+35,position.y - vertical_offset);
		changeObjectVisibility(menu_id,'visible');
		popupVisible = true;
	}
}

function hide_info_popup(thisObj,menu_id) {
	menuObj = document.getElementById(menu_id);
	// moveObject(menu_id,1,1);
	changeObjectVisibility(menu_id,'hidden');
	popupVisible = false;
}

function changeObjectVisibility(objectId, newVisibility) {
    // get a reference to the cross-browser style object and make sure the object exists
    var styleObject = getStyleObject(objectId);
    if(styleObject) {
	styleObject.visibility = newVisibility;
	return true;
    } else {
    	return false;
    }
}

function getStyleObject(objectId) {
     if(document.getElementById(objectId)){
	   return (document.getElementById(objectId).style);
     } else {
	   return false;
     }
}

function moveObject(objectId, newXCoordinate, newYCoordinate) {
    var styleObject = getStyleObject(objectId);
    if(styleObject) {
	styleObject.left = newXCoordinate;
	styleObject.top = newYCoordinate;
    }
}

function getAnchorPosition(anchor_id) {// This function will return an Object with x and y properties
	var position=new Object();
	// Logic to find position
	position.x=AnchorPosition_getPageOffsetLeft(document.getElementById(anchor_id));
	position.y=AnchorPosition_getPageOffsetTop(document.getElementById(anchor_id));
	return position;
}

function AnchorPosition_getPageOffsetLeft (el) {
	var ol=el.offsetLeft;
	while((el=el.offsetParent) != null) {
	  ol += el.offsetLeft;
	}
	return ol;
}

function AnchorPosition_getPageOffsetTop (el) {
	var ot=el.offsetTop;
	while( (el=el.offsetParent) != null) {
	  ot += el.offsetTop;
	}
	return ot;
}
//-->
</SCRIPT>











</head>
<body onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<div align="center"><br>
  <br>
  <br>
  <font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ligar urgente.</strong></font> </center> 
</div>
<center>
</center>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere
dim varCod_imovel

SearchWhere = request.querystring("SearchWhere")
SearchFor = request.querystring("SearchFor")

session("SearchWhere") = SearchWhere
session("SearchFor") = SearchFor


Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio
color2 = claro




Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

 
 
 dim vDataAtual
 
if len(now()) = 19 then
vDataAtual = left(now(),11)


end if


if len(now()) = 18 then
vDataAtual = left(now(),10)


end if


if len(now()) = 17 then
vDataAtual = left(now(),9)


end if


 
  

if session("permissao") <> 6 then

SQL = "select  compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(vDataAtual, " ") = 0
		SQL = SQL & "data_ligar_urgente like '%" _
			& left(vDataAtual, instr(vDataAtual," ") - 1) & "%' or "
		vDataAtual = Right(vDataAtual, len(vDataAtual) - instr(vDataAtual," "))
	loop
	if len(vDataAtual) > 1 then
		SQL = SQL & "data_ligar_urgente like '%" & vDataAtual & "%' and atendimento like '"& Session("nome_id") &"' and (standby <>'comprou com a Veja' and  standby <>'comprou com outro' and standby <>'suspenso' and standby <>'cliente inexistente' and standby <>'cliente com proposta') "&" ORDER  BY cod_compradores DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" and atendimento like '"& Session("nome_id") &"' and (standby like 'comprador a contatar' or standby like 'comprador OK' or standby like'comprou com outro' or standby like'não informado') ORDER  BY cod_compradores DESC"
	end if




else




SQL = "select  compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and "
do until instr(vDataAtual, " ") = 0
		SQL = SQL & "data_ligar_urgente like '%" _
			& left(vDataAtual, instr(vDataAtual," ") - 1) & "%' or "
		vDataAtual = Right(vDataAtual, len(vDataAtual) - instr(vDataAtual," "))
	loop
	if len(vDataAtual) > 1 then
		SQL = SQL & "data_ligar_urgente like '%" & vDataAtual & "%' "&" and (standby <>'comprou com a Veja' and  standby <>'comprou com outro' and standby <>'suspenso' and standby <>'cliente inexistente' and standby <>'cliente com proposta') ORDER  BY cod_compradores DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" and (standby like 'comprador a contatar' or standby like 'comprador OK' )  ORDER  BY cod_compradores DESC"
	end if







end if




%>
<%

Dim LinkTemp
'essa variável vai ser usada como contador


'as variáveis acima são usadas para trocar a cor das tabelas que conterão os valores
'dos recordsets.






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
	
RS.Open SQL, Conn, 1, 3
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
'verifica se existem registros retornados.
%><br>
<center></center>

<form  Method="Post" name="Formulario" action="multi_excluir_imovel.asp?varCod_imovel=<%=varCod_imovel%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&page=<%=cInt(intPage)%>" >
  <table width="900" border="0" cellspacing="0" cellpadding="0">
    <tr bgcolor="#000000"> 
	
	
    
	
	<td width="133" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendimento</strong></font></div></td>
      
      
		  
		<td width="40" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Indica</strong></font></div></td>  
		  
		  
	  <td width="250" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Raz&atilde;o 
          do ligar urgente</strong></font></div></td>
     
	  
      
     
	  
	  
	  <td width="135" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
          de ligar urgente</strong></font></div></td>
    
	<td width="140" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
          de inclus&atilde;o</strong></font></div></td>
    
	
	</tr>
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
    <%










'------------------------------------------------

%>
    <%
If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If

dim vValor


 session("vValor1")=left(vValor,10)
   session("vValor2")=right(vValor,10)


%>
    <% session("page")=intPage%>
   
    <tr> 
     
	  
     
	 
	 <td width="133" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("atendimento")%></strong></font></div></td>
	 
	 
	  
	  <td width="40" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>



<%



'------------------------Cidade---------------------------

dim stringIndex2
dim stringCidade2


stringIndex2 = " where cod_imovel<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "não informado" and rs("cidade") <> ""   then
stringCidade2 = " and cidade='"&rs("cidade")&"'"
else
stringCidade2 = ""
end if

 '--------------------------Bairro----------------------------








if ( rs("bairro") <> "qualquer um" and  rs("bairro") <> "não informado" and  rs("bairro") <> "") then


 
dim Numero_Indicacoes
dim Numero_Indicacoes02




Numero_Indicacoes = 0
Numero_Indicacoes02 = 0


dim soma02
dim soma

soma = 0
soma02 = 0

dim Variavel
dim Retorno
dim contar
Variavel =  rs("bairro")
Retorno = Split(rs("bairro"),", ")

contar=0

dim stringBairro3
dim stringBairro4
dim stringBairro5

for contar=0 to UBound(Retorno)

stringBairro3 = "and ( "
stringBairro4 = " Bairro='"&Retorno(contar)&"'or  " &stringBairro4

stringBairro5 = " cod_imovel=0)"


stringBairro2 = stringBairro3&stringBairro4&stringBairro5







next

stringBairro3 = ""
stringBairro4 = ""
stringBairro5 = ""


else
stringBairro2 = ""
end if








 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" and rs("tipo") <> ""  then




 
dim Numero_IndicacoesTipo
dim Numero_Indicacoes02Tipo




Numero_IndicacoesTipo = 0
Numero_Indicacoes02Tipo = 0


dim soma02Tipo
dim somaTipo

somaTipo = 0
soma02Tipo = 0

dim VariavelTipo
dim RetornoTipo
dim contarTipo
VariavelTipo =  rs("tipo")
RetornoTipo = Split(rs("tipo"),", ")

contarTipo=0

dim stringTipo3
dim stringTipo4
dim stringTipo5

for contarTipo=0 to UBound(RetornoTipo)

stringTipo3 = "and ( "
stringTipo4 = " tipo='"&RetornoTipo(contarTipo)&"'or  " &stringTipo4

stringTipo5 = " cod_imovel=0)"


stringTipo2 = stringTipo3&stringTipo4&stringTipo5







next

stringTipo3 = ""
stringTipo4 = ""
stringTipo5 = ""


else
stringTipo2 = ""
end if







 '------------------------------------------------------------- 







'-------------------Negociação---------------------------
if rs("negociacao") = "Compra"  then
vNegocio = "venda"
end if

if rs("negociacao") = "compra" then
vNegocio = "venda"
end if

if rs("negociacao") = "Aluguel" then
vNegocio = "aluguel"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if


if  rs("negociacao") <> "qualquer um" and rs("negociacao") <> "" and rs("negociacao") <> "nqualquer" and rs("negociacao") <> "não informado" and rs("negociacao") <> "" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> int(0) and rs("quartos") <> "" then
stringQuartos2 = " and quartos >="&rs("quartos")&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> int(0) and rs("vagas") <> "" then
stringVagas2 = " and vagas >="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------------Valor-----------------------------------


   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)




'stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""

stringValor2 = " and Valor <="& vValorMaior &""



'---------------------------------Condominio-----------------------------------



dim stringCondominio


Porcentual02 = int(rs("condominio"))*10/100
   


   vCondominioMenor = int(rs("condominio")) - int(Porcentual02)
   vCondominioMaior = int(rs("condominio")) + int(Porcentual02)




if  int(rs("condominio")) <> 0 and rs("condominio") <> ""  then

'stringCondominio = " and Condominio >="& vCondominioMenor &" and Condominio <="& vCondominioMaior &""
stringCondominio = "  and Condominio <="& vCondominioMaior &""

else
stringCondominio = ""
end if


'---------------------------------------------------------------------------


'---------------------------------Área Construida-----------------------------------



dim stringAreaConstruida


Porcentual03 = int(rs("area_construida"))*10/100
   


   vAreaConstruidaMenor = int(rs("area_construida")) - int(Porcentual03)
   vAreaConstruidaMaior = int(rs("area_construida")) + int(Porcentual03)



if  int(rs("area_construida")) <> 0 and rs("area_construida") <> "" then
'stringAreaTotal = " and area_total >="& vAreaTotalMenor &" and area_total <="& vAreaTotalMaior &""
stringAreaConstruida = " and area_construida >="& vAreaConstruidaMenor &""


else
stringAreaConstruida = ""
end if


'---------------------------------------------------------------------------













'-------------------------------Suítes-----------------------------------------


dim stringSuites
 
if  rs("suites") <> "suiqualquer" and rs("suites") <> "não" and rs("suites") <> "0" and rs("suites") <>"00" and rs("suites") <>"" then
stringSuites = "  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina
 
if  rs("piscina") <> "pisqualquer" and rs("piscina") <> "não" and rs("piscina") <> "0" and rs("piscina") <>"00" and rs("piscina") <>"" then
stringPiscina = "  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria
 
if  rs("portaria") <> "porqualquer" and rs("portaria") <> "não" and rs("portaria") <> "0" and rs("portaria") <>"00" and rs("portaria") <>"" then
stringPortaria = "  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal
 
if  rs("quintal") <> "quiqualquer" and rs("quintal") <> "não" and rs("quintal") <> "0" and rs("quintal") <>"00" and rs("quintal") <>"" then
stringQuintal = "  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras
 
if  rs("quadras") <> "quaqualquer" and rs("quadras") <> "não" and rs("quadras") <> "0" and rs("quadras") <>"00" and rs("quadras") <>"" then
stringQuadras = "  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula
 
if  rs("edicula") <> "ediqualquer" and rs("edicula") <> "não" and rs("edicula") <> "0" and rs("edicula") <>"00" and rs("edicula") <>"" then
stringEdicula = "  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Ocupação--------------------------------------
dim stringOcupacao
 
if  rs("ocupacao") <> "oqualquer" and rs("ocupacao") <> "não informado" and rs("ocupacao") <> ""  then
stringOcupacao = "  and ocupacao ='"&rs("ocupacao")&"'  and ocupacao IS NOT NULL"




else

stringOcupacao = ""
end if



'--------------------------------------------------------------------------------





dim stringStandby

stringStandby = " and ( imovel_em_negociacao like  '"&"imóvel OK"&"') "

 Set rs444 = Server.CreateObject("ADODB.RecordSet")
'se no cliente ou no servidor.


	strSQL444 = "SELECT imoveis.cod_imovel FROM imoveis"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio&stringAreaConstruida&stringSuites&stringPiscina&stringPortaria&stringQuintal&stringQuadras&stringEdicula&stringOcupacao&stringStandby&" ORDER  BY cod_imovel DESC"
	
	
	
	
 varIndicacaoCidade = rs("cidade")
	 varIndicacaoBairro = rs("bairro")
	 varIndicacaoNegociacao = rs("negociacao")
	 varIndicacaoTipo = rs("tipo")
	 varIndicacaoQuartos = rs("quartos")
	 varIndicacaoVagas = rs("vagas")
	 varIndicacaoValor = rs("Valor")
	
	 varIndicacaoCondominio = rs("condominio")
	 varIndicacaoAreaConstruida = rs("area_construida")
	 varIndicacaoSuites = rs("suites")
	 varIndicacaoPiscina = rs("Piscina")
	 varIndicacaoPortaria = rs("portaria")
	 varIndicacaoQuintal = rs("quintal")
	 varIndicacaoQuadras = rs("quadras")
	  varIndicacaoEdicula = rs("edicula")
	  varIndicacaoOcupacao = rs("ocupacao")
	
	 
	varCodIndicacao = "'"&strSQL444&"'"
	 
		
Rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
Rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.
 
	 
	 rs444.Open strSQL444,Conexao 
	 
	 
	   
	   
     %>
 
<% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or  session("permissao") = "5" or  session("permissao") = "6"  then %><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow22('indicacao_compradores22.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>&varIndicacaoCondominio=<%=varIndicacaoCondominio%>&varIndicacaoAreaConstruida=<%=varIndicacaoAreaConstruida%>&varIndicacaoSuites=<%=varIndicacaoSuites%>&varIndicacaoPiscina=<%=varIndicacaoPiscina%>&varIndicacaoPortaria=<%=varIndicacaoPortaria%>&varIndicacaoQuintal=<%=varIndicacaoQuintal%>&varIndicacaoQuadras=<%=varIndicacaoQuadras%>&varIndicacaoEdicula=<%=varIndicacaoEdicula%>&varIndicacaoOcupacao=<%=varIndicacaoOcupacao%>&varCodCompradores=<%=rs("cod_compradores")%>')"><%=rs444.recordcount%><br></a></strong></font><%else%><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444.recordcount%><br></strong></font><%end if%>

	 <%
	 
 do while not rs444.eof 

 
 rs444.movenext
loop
 
 

  %>
  
  <%
  
 
 
 rs444.close
 set rs444 = nothing
  
 %>

 
 
 <%
 
 









%>	


 


 

</font></div>







</font></div></td>
	 
	 
	 
	  <td width="250" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><br><a href="javascript:newWindow2020('visualizar_compradores33.asp?varCodCompradores=<%=rs("cod_compradores")%>')"><%=rs("assunto_ligar_urgente")%></a><br></strong></font></div><br></td>
	  
	  
	  
       
	 
	
	 
	 
      
		  
		  
      
      
	  
	  
     
	 
	  <td width="135" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("data_ligar_urgente")%></strong></font></div></td>
    
	 <td width="140" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("data")%></strong></font></div></td>
    
	
	</tr>
    <%
'-----------------------------------------------









rs.movenext
If RS.EOF Then Exit for
Next

%>
  </table>
</form>





<table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#003366"> 
        <%If cInt(intPage) > 1 Then%>
        <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
      <a href="?page=<%=intPage - 1%>" style="color:#000000">
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
        <font color="#000000">Página</font> <%=cInt(intPage)%> <font color="#000000">de</font> 
        <%=cInt(intPageCount)%> </font> 
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
        <a href="?page=<%=intPage + 1%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Próximo</b></font> 
        </a> 
        <%End If%>
        </font></div></td>
        </tr>
      </table>










 <%else%>
 
  <table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
 </table>
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Compradores n&atilde;o encontrados</font></div>
</font> 
<%
End If%>
<%else%>
<table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
	
 </table>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Compradores</font><font color="<%=escuro%>"> 
  n&atilde;o encontrados</font></div>
</font> 
<%
End if
%>
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
group2[2][3]=new Option("201,00 até 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 até 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 até 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 até 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 até 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 até 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 até 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 até 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 até 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Até  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 até 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 até 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 até 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 até 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 até 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 até 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 até 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 até 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 até 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("Acima de 400.000,00","0000400001 1000000000")









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
           'fecha a conexão
           
		   
           Set rs = Nothing
		   
		   
		'-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------  


'-----------------------------
           rs4.Close           
		   
           Set rs4 = Nothing
		   
'---------------------------------

Set rs4Verifica = Nothing

 
		   
		   
		    conexao3.close
		   set conexao3 = nothing
		   
		   
		   conexao.close
		   set conexao = nothing
		   
           %>
  <% response.flush%>
  <%response.clear%>
  
  <br>
  <br>
  <center>
</center>
</body>
</html>

