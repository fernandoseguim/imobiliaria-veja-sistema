<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->

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


 
 
 if session("permissao") = "6" then


SQL = "SELECT TOP 10 imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta  FROM imoveis where (data02_tarja02 ='"&day(vDataAtual)&"' or data01_tarja02 ='"&day(vDataAtual)&"')  and  data_contato NOT LIKE '" & vDataAtual & "%' and imovel_em_negociacao <> 'Vendido pela Veja' and imovel_em_negociacao <> 'Vendido por outros' and imovel_em_negociacao <> 'Im�vel inexistente' order by cod_imovel DESC"
	
else

'SQL = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta  FROM imoveis where (data01_tarja02 ='"&day(vDataAtual)&"' or data01_tarja02 ='"&day(vDataAtual)&"') and   data_contato NOT LIKE '%" & vDataAtual & "%' and captacao like '" & session("nome_id") & "' and imovel_em_negociacao <> 'Vendido pela Veja' and imovel_em_negociacao <> 'Vendido por outros' order by cod_imovel DESC"
	
	
SQL = "SELECT TOP 10 imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta  FROM imoveis where (data02_tarja02 ='"&day(vDataAtual)&"' or data01_tarja02 ='"&day(vDataAtual)&"') and data_contato NOT LIKE '" & vDataAtual & "%' and captacao like '" & session("nome_id") & "' and imovel_em_negociacao <> 'Vendido pela Veja' and imovel_em_negociacao <> 'Vendido por outros' and imovel_em_negociacao <> 'Im�vel inexistente'  order by cod_imovel DESC"
		
	
	
	
end if
 
 
 
Dim LinkTemp
'essa vari�vel vai ser usada como contador


'as vari�veis acima s�o usadas para trocar a cor das tabelas que conter�o os valores
'dos recordsets.






dim intPage
'essa vari�vel vai receber um valor inicial "1" que mostra que estamos na primeira p�gina.

dim intPageCount
'Essa vari�vel vai receber o valor da quantidade de p�ginas do recordset.

dim intRecordCount
'Essa vari�vel vai receber o n�mero de recordsets existentes.

If Request.QueryString("page") = "" Then
	intPage = 1	
Else
	intPage = Request.QueryString("page")
End If
'aqui a vari�vel intPage recebe o valor "1" na primeira p�gina.
	
RS.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

RS.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

RS.ActiveConnection = Conn
'a propriedade ActiveConnection indica qual conex�o o recordset utilizar�.
	
RS.Open SQL, Conn, 1, 3
'o recordset � aberto
	
RS.PageSize = 20
'Aqui configura-se o recordset para 20 registros por p�gina.

RS.CacheSize = RS.PageSize
'o Cache tamb�m conter� 20 registros por p�gina.

intPageCount = RS.PageCount
'A vari�vel intPageCount recebe o valor do n�mero de p�gina do recordset retornado.

intRecordCount = RS.RecordCount
'A vari�vel intRecordCount recebe o valor do n�mero de registros retornados no recordset.

 
 
 
  '-------------------Verificar se j� foram atualizados 10 clientes---------
   
   dim rs444Indicacao02
   dim strSQL444Indicacao02
   
   
   Set rs444Indicacao02 = Server.CreateObject("ADODB.RecordSet")
	
	 

	'strSQL444Indicacao = "SELECT TOP "&vNumero&" compradores.area_construida,compradores.area_total,compradores.condominio,compradores.cod_compradores   from compradores where atendimento like '"&vAtendimento01&"'"

strSQL444Indicacao02 = "SELECT * FROM imoveis where   data_contato  LIKE '%" & vDataAtual & "%' and  captacao  LIKE '%" & session("nome_id") & "%'  order by cod_imovel DESC"
	

         
         rs444Indicacao02.CursorLocation = 3
        rs444Indicacao02.CursorType = 3
        
 
            rs444Indicacao02.ActiveConnection = Conexao




	rs444Indicacao02.Open strSQL444Indicacao02, Conexao,1,3
   
   
   rs444Indicacao02.PageSize = 10
'Aqui configura-se o recordset para 20 registros por p�gina.

rs444Indicacao02.CacheSize = rs444Indicacao02.PageSize
'o Cache tamb�m conter� 20 registros por p�gina.

intPageCount02 = rs444Indicacao02.PageCount
'A vari�vel intPageCount recebe o valor do n�mero de p�gina do recordset retornado.

intRecordCount02 = rs444Indicacao02.RecordCount
   
   while not rs444Indicacao02.eof 
   
   
   
   rs444Indicacao02.movenext
   wend
   
   
    dim vAtualizado
   
   vAtualizado = "n�o"
   
   
   if int(intRecordCount02) >= int(intRecordCount) then
   session("ok") = "ok"
   
   
   vAtualizado = "sim"
   
   'response.redirect "archive_imoveis.asp?varCodOK="&session("ok")&""
   
   end if
   
   '-------------------------------------------------------------------------
   
   



%>





<html>
<head>
<title>Fidelizar im�vel</title>

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
   openWindow2020 = window.open(abrejanela2020,'openWin2020','width=610,height=500,resizable=yes,scrollbars=yes')
   openWindow2020.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
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









<script language="JavaScript">
var today=new Date();
var todaysec=today.getSeconds();

function xpop(){
if (confirm("Voc� precisa atualizar mais clientes para sair dessa p�gina.")){
window.open('archive_futuro_contato_imovel02.asp', todaysec+'floyd','width=605,height=530,resizable=yes,scrollbars=yes,Left=0,Top=0')

}
else {

(confirm("Voc� precisa atualizar mais clientes para sair dessa p�gina."))
window.open('archive_futuro_contato_imovel02.asp', todaysec+'floyd','width=605,height=530,resizable=yes,scrollbars=yes,Left=0,Top=0')

}   
}
</script>


</head>
<body <%if vAtualizado <> "sim" then %>onBlur="_blank"  onUnload="xpop()"<%else%><%end if%>  onload=document.forms.b2.SearchFor.focus(); topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<br>
<center>
<br>
  <br>
  <br>
  <font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Im&oacute;veis 
  a serem atualizados hoje.</strong></font> 
</center>
<br>
<center>
  <a href="archive_futuro_contato_imovel02.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Carregar 
  P�gina</strong></font></a> 
</center>

<center>
</center>
<%

 
 
 
 
  

'if session("permissao") = "6" then

'SQL = "select * from imoveis where "
'do until instr(day(vDataAtual), " ") = 0
'		SQL = SQL & "data_futuro_contato like '%" _
'			& left(vDataAtual, instr(vDataAtual," ") - 1) & "%' or "
'		vDataAtual = Right(vDataAtual, len(vDataAtual) - instr(vDataAtual," "))
	'loop
	'if len(vDataAtual) > 1 then
	'	SQL = SQL & "data_futuro_contato like '%" & vDataAtual & "%' "&" ORDER  BY data_atualizacao DESC"
	'else
	'	SQL = left(SQL, len(SQL) - 4)
	'	SQL = SQL&"  ORDER  BY data_atualizacao DESC"
	'end if

'else


'SQL = "select * from imoveis where "
'do until instr(vDataAtual, " ") = 0
'		SQL = SQL & "data_futuro_contato like '%" _
'			& left(vDataAtual, instr(vDataAtual," ") - 1) & "%' or "
'		vDataAtual = Right(vDataAtual, len(vDataAtual) - instr(vDataAtual," "))
'	loop
'	if len(vDataAtual) > 1 then
	'	SQL = SQL & "data_futuro_contato like '%" & vDataAtual & "%' and captacao like '" & session("nome_id") & "'  "&" ORDER  BY data_atualizacao DESC"
	'else
	'	SQL = left(SQL, len(SQL) - 4)
	'	SQL = SQL&" and captacao like '" & session("nome_id") & "'  ORDER  BY data_atualizacao DESC"
	'end if


'end if







'----------------------------Atualiza��o da tarja------------------------------------------



Set rs444Indicacao = Server.CreateObject("ADODB.RecordSet")
	
	 

	'strSQL444Indicacao = "SELECT TOP "&vNumero&" compradores.area_construida,compradores.area_total,compradores.condominio,compradores.cod_compradores   from compradores where atendimento like '"&vAtendimento01&"'"

strSQL444Indicacao = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta  FROM imoveis where (data01_tarja02 ='"&day(vDataAtual)&"' or data01_tarja02 ='"&day(vDataAtual)&"')  and  data_contato NOT LIKE '%" & vDataAtual & "%' and imovel_em_negociacao <> 'Vendido pela Veja' and imovel_em_negociacao <> 'Vendido por outros' order by cod_imovel DESC"
	

         
         rs444Indicacao.CursorLocation = 3
        rs444Indicacao.CursorType = 3
        
 
            rs444Indicacao.ActiveConnection = Conexao




	rs444Indicacao.Open strSQL444Indicacao, Conexao
	
	dim num
	num = 0
	
	 if not rs444Indicacao.eof  then 
				     While   not rs444Indicacao.eof
                   
                    'Conexao.execute"update imoveis set imovel_em_negociacao='"&"n�o informado"&"' where imovel_em_negociacao IS NULL" 
                   
				    Conexao.execute"update  imoveis set tarja02='"&"n�o"&"' where  cod_imovel="&rs444Indicacao("cod_imovel")
                   
				   
                   rs444Indicacao.MoveNext 
                     Wend 
					
					else
					
					end if
	
	
	
	'-------------------------------
   rs444Indicacao.close
  
  
  set rs444Indicacao = nothing
 '------------------------------






'-----------------------------------------------------------------------------------------















%>
<%

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%><br>
<%'vDataAtual%>
<center></center>
<form  Method="Post" name="Formulario" action="multi_excluir_imovel.asp?varCod_imovel=<%=varCod_imovel%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&varCidade=<%=session("varCidade")%>&varBairro=<%=session("varBairro")%>&varNegociacao=<%=session("varNegociacao")%>&varQuartos=<%=session("varQuartos")%>&page=<%=cInt(intPage)%>" >
  <table width="800" border="0" cellspacing="0" cellpadding="0">
    <tr bgcolor="#000000"> 
      <td width="132" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>C&oacute;digo 
          de refer&ecirc;ncia </strong></font></div></td>
	 
	 <td width="40" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Indicar</strong></font></div></td>
	 
	 
	  <td width="134" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Propriet&aacute;rio</strong></font></div></td>
      
	  <td width="133" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
      <td width="133" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email</strong></font></div></td>
     
	  
	  <td width="133" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Capta&ccedil;&atilde;o</strong></font></div></td>
      
	  <td width="135" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
          de inclus&atilde;o</strong></font></div></td>
    </tr>
    <%	
If CInt(intPage) > CInt(intPageCount) Then intPage = intPageCount
'se intPage � maior que o n�mero de p�ginas ent�o intPage � igual ao n�mero de p�ginas.

	If CInt(intPage) <= 0 Then intPage = 1
	'se intPage � menor ou igual a zero ent�o intPage igual a "1"
	'a vari�vel intPage sempre vai ser for�ada a receber o valor "1".
	
		If intRecordCount > 0 Then
		'se existirem registros retornados ent�o.
			 
			 RS.AbsolutePage = intPage
			'a propriedade AbsolutePage determina a p�gina exata que o registro atual
			'reside
			
			intStart = RS.AbsolutePosition
			'a vari�vel intStart recebe o valor da propriedade AbsolutePosition que
			'corresponde a posi��o exata do primeiro registro da p�gina correspondente.
			
			
			
			If CInt(intPage) = CInt(intPageCount) Then
			'se intPage � igual ao n�mero de p�ginas no recordset , estamos na �ltima 
			'p�gina ent�o.
				intFinish = intRecordCount
				'a vari�vel intFinish recebe o valor do n�mero do �ltimo recordset.
				'intFinish corresponde ao valor do �ltimo registro da p�gina correspondente.
			Else
				intFinish = intStart + (RS.PageSize - 1)
				'a vari�vel intFinish recebe o valor de intStart + o valor
				'do n�mero de registros na p�gina menos "1".
			End if
		End If
	If intRecordCount > 0 Then
	'se houver registros ent�o
		For intRecord = 1 to RS.PageSize
		'um contador inRecord � colocado at� o n�mero de registros na p�gina.
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
      <td width="132" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("cod_imovel")%></strong></font></div></td>
	
	 <td width="40" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>

<%

Set rs444 = Server.CreateObject("ADODB.RecordSet")




'------------------------Cidade---------------------------

stringIndex2 = " where cod_compradores<>"&"0"&""


if rs("cidade") <> "qualquer um" and rs("cidade") <> "n�o informado"  then
stringCidade2 = " and (cidade='"&rs("cidade")&"' or cidade='"&"n�o informado"&"')"
else
stringCidade2 = ""
end if



 '--------------------------Bairro----------------------------

if rs("bairro") <> "qualquer um" and rs("bairro") <> "n�o informado" then
stringBairro2 = " and (Bairro like '%"&rs("bairro")&"%' or Bairro like'%"&"n�o informado"&"%')"
else
stringBairro2 = ""
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if rs("tipo") <> "qualquer um" and rs("tipo") <> "tqualquer" then
stringTipo2 = " and Tipo like'%"&rs("Tipo")&"%'"
else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negocia��o---------------------------
vNegocio = "Compra"
if rs("negociacao") = "venda" then
vNegocio = "compra"
end if

if rs("negociacao") = "aluguel" then
vNegocio = "aluguel"
end if

if  rs("negociacao") <> "qualquer um" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  rs("quartos") <> 0 then
stringQuartos2 = " and quartos<="&rs("quartos")&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------


'---------------------------Vagas------------------------------


if  rs("vagas") <> 0 then
stringVagas2 = " and vagas <="&rs("vagas")&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------





'---------------------------------Valor-----------------------------------



 
   Porcentual = int(rs("valor"))*10/100
   


   vValorMenor = int(rs("valor")) - int(Porcentual)
   vValorMaior = int(rs("valor")) + int(Porcentual)
  








stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""




dim stringStandby

stringStandby = " and standby = '"&"excluido"&"'"










'---------------------------------------------------------------------------



	strSQL444 = "SELECT compradores.cod_compradores  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	
	 varIndicacaoCidade = rs("cidade")
	 varIndicacaoBairro = rs("bairro")
	 varIndicacaoNegociacao = rs("negociacao")
	 varIndicacaoTipo = rs("tipo")
	 varIndicacaoQuartos = rs("quartos")
	 varIndicacaoVagas = rs("vagas")
	 varIndicacaoValor = rs("Valor")
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
	varCodIndicacao = "'"&strSQL444&"'"
	 
		
Rs444.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

Rs444.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.
 
	 
	 rs444.Open strSQL444,Conexao 
	 
	   
     %>
          <% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6"  then %>
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow22('indicacao_imoveis22.asp?varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoVagas=<%=varIndicacaoVagas%>&varIndicacaoValor=<%=varIndicacaoValor%>')"><%=rs444.RecordCount%></a></strong></font> 
          <%else%>
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444.RecordCount%></strong></font> 
          <%end if%>
          <%
	 
 do while not rs444.eof 

 
 
 rs444.movenext
loop
 
 rs444.close
  
 
 
set rs444 = nothing








%>







</strong></font></div></td>
	
	
	
	  <td width="134" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow2020('visualizar_imovel33.asp?varCod_imovel=<%=rs("cod_imovel")%>')"><%=rs("proprietario")%></a></strong></font></div></td>
	  
       <td width="133" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("telefone")%> 
          </strong></font></div></td>
	 
	  <td width="133" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("email")%></strong></font></div></td>
	 
      
		  
		  
      
      
	  
	  <td width="133" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("captacao")%></font></div></td>
      <td width="135" height="30" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("data")%></font></div></td>
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
        <!-- se a p�gina atual for maior que "1" ent�o o link anteriro � colocado na 
			  na tela .-->
      <a href="?page=<%=intPage - 1%>" style="color:#000000">
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se p�gina atual � menor que o total de p�ginas e intPage maior que um
			  ou seja, se n�o estiver na primeira p�gina e nem na �ltima ent�o. -->
        <font color="#000000">P�gina</font> <%=cInt(intPage)%> <font color="#000000">de</font> 
        <%=cInt(intPageCount)%> </font> 
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage � menor que o n�mero de p�ginas ent�o colocar o bot�o pr�ximo -->
        <a href="?page=<%=intPage + 1%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Pr�ximo</b></font> 
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
<div align="center"I><font color="<%=escuro%>">Im&oacute;veis</font><font color="<%=escuro%>"> 
  n&atilde;o encontrados</font></div>
</font> 
<br>
<%
session("ok") = "ok"
'response.redirect "archive_imoveis.asp?varCodOK="&session("ok")&""

%>


<center>
  <a href="archive_imoveis.asp?varCodOK=<%=session("ok")%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Pr�xima p�gina 
  </strong></font></a> 
</center>

<%
End If%>
<%else%>
<table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">&nbsp;</td>
      
    </tr>
	
 </table>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Im&oacute;veis</font><font color="<%=escuro%>"> 
  n&atilde;o encontrados</font></div>
</font>

<br>

<%

session("ok") = "ok"
'response.redirect "archive_imoveis.asp?varCodOK="&session("ok")&""


%>


<center>
  <a href="archive_imoveis.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Pr�xima p�gina 
  </strong></font></a> 
</center>

 
<%
End if
%>
   
<%
  rs.Close
           'fecha a conex�o
           Conexao.Close
           Set rs = Nothing
           %>
  <% response.flush%>
  <%response.clear%>
  <br>
  <br>
  <center>
<%'SQL%>

</center>
</body>
</html>

