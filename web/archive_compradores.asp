<% response.Buffer = true %>

<!--#include file="dsn.asp"-->
<!--#include file="style6_imoveis.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="loggedin.asp"-->
<%


dim vOK

vOK = request.QueryString("varCodOK")

if vOK = "" then
vOK = session("ok")
end if

'if (vOK <> "ok" and session("permissao") <> "5" and session("permissao") <> "6") then
'response.redirect "archive_futuro_contato_comprador02.asp"

'end if




session("acessos")

dim vdataUrgente
dim vdataUrgente2
dim vdataUrgente3

vdataUrgente = now()


dim stringCidade2,stringBairro2,stringNegociacao2,stringQuartos2,stringVagas2,stringValor2,stringTipo2
dim stringIndex2
dim vNegocio
dim vValorMenor,vValorMaior
dim varCodIndicacao

dim varIndicacaoCidade
dim varIndicacaoBairro
dim varIndicacaoNegociacao
dim varIndicacaoQuartos
dim varIndicacaoVagas
dim varIndicacaoValor
dim varIndicacaoTipo


vValorMenor = int("0")
vValorMaior = int("0")

dim porcentual

%>



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


dim varCidade,stringCidade,varBairro,stringBairro,varNegociacao,varVagas
dim stringNegociacao,varQuartos,stringQuartos,varCidade2,stringVagas

 varCidade2 = request.querystring("combo1")


if varCidade2 = "" then
varCidade2 = request.querystring("varCidade2")
session("varCidade2") = varCidade2
end if




 if varCidade2 = "" then
 varCidade2 = "cqualquer"
 end if
 
session("varCidade2") = varCidade2


 
 if varCidade2 <> "cqualquer" then
 dim rrs2,SSQL2



 Set rrs2 = Server.CreateObject("ADODB.RecordSet")
 SSQL2 = "select combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  from combo1 where id_combo1="&varCidade2
 
 
 
          rrs2.CursorLocation = 3
          rrs2.CursorType = 3

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
 
 if session("varCidade2") = "" then
session("varCidade2") = "cqualquer"
end if
 
 
 
 '------------------------------pegar tipo--------------------------------

 dim varTipo
 
 varTipo =request.QueryString("txt_tipo")
 
 if varTipo = "" then
 varTipo = request.querystring("varTipo")
 end if
 
    if request.QueryString("varTipo")<>"" then
 varTipo = request.QueryString("varTipo")
 session("varTipo") = varTipo
 else
 
 session("varTipo") = varTipo
 end if

 
 
 
 
 
 
 
 
 
varBairro2 = request.querystring("combo2")



if varBairro2 = "" then
varBairro2 = request.querystring("varBairro2")

end if




if varBairro2 = "" then
varBairro2 = "bqualquer"
session("varBairro2") = "bqualquer"
end if

 if varBairro2 <> "bqualquer" then
	  dim rrs3,SSQL3,conexao5
	 
	 
	 
	 
 Set rrs3 = Server.CreateObject("ADODB.RecordSet")
 SSQL3 = "select combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  from combo2 where id_combo2 ="&varBairro2
 
 
 
 
rrs3.CursorLocation = 3
rrs3.CursorType = 3

rrs3.ActiveConnection = Conexao3
 
 
 
 rrs3.open SSQL3,Conexao3,2,1

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
 
 session("varBairro2") = varBairro2



if session("varBairro2") = "" then
session("varBairro2") = request.querystring("varBairro2")
end if


if session("varBairro2") = "" then
session("varBairro2") = "bqualquer"
end if




varNegociacao = request.querystring("example2")
varQuartos = request.querystring("txt_quartos")





dim varValor, varValor1,varValor2
 

varValor = request.QueryString("stage22")

session("varValor")=varValor
  
  if varValor = "" then
  
  varValor = "vqualquer"
  
  end if


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
	
	if session("varValor") = "" then
	session("varValor") = "vqualquer"
	end if
	
	
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


varVagas = request.querystring("txt_vagas")

 if varVagas = "" then
varVagas = request.QueryString("varVagas")
end if

 if request.QueryString("varVagas")<>"" then
 varVagas = request.QueryString("varVagas")
 session("varVagas") = varVagas
 else
 
 session("varVagas") = varVagas
 end if


'-----------Suítes-------------------------

dim varSuites

varSuites = request.querystring("txt_suites")

 if varSuites ="" then
 varSuites = request.QueryString("varSuites")
 session("varSuites") = varSuites
 else
 
 session("varSuites") = varSuites
 end if
 
 if session("varSuites") = "" then
 session("varSuites") = "suiqualquer"
 end if
 
'-------------------------------


'-----------Piscina-------------------------

dim varPiscina

varPiscina = request.querystring("txt_piscina")

 if varPiscina ="" then
 varPiscina = request.QueryString("varPiscina")
 session("varPiscina") = varPiscina
 else
 
 session("varPiscina") = varPiscina
 end if
 
 if session("varPiscina") = "" then
 session("varPiscina") = "pisqualquer"
 end if
 
'-------------------------------




'-----------Portaria-------------------------

dim varPortaria

varPortaria = request.querystring("txt_portaria")

 if varPortaria ="" then
 varPortaria = request.QueryString("varPortaria")
 session("varPortaria") = varPortaria
 else
 
 session("varPortaria") = varPortaria
 end if
 
 if session("varPortaria") = "" then
 session("varPortaria") = "porqualquer"
 end if
 
'-------------------------------


'-----------Quintal-------------------------

dim varQuintal

varQuintal = request.querystring("txt_quintal")

 if varQuintal ="" then
 varQuintal = request.QueryString("varQuintal")
 session("varQuintal") = varQuintal
 else
 
 session("varQuintal") = varQuintal
 end if
 
 if session("varQuintal") = "" then
 session("varQuintal") = "quiqualquer"
 end if
 
'-------------------------------


'-----------Quadras-------------------------

dim varQuadras

varQuadras = request.querystring("txt_quadras")

 if varQuadras ="" then
 varQuintal = request.QueryString("varQuadras")
 session("varQuadras") = varQuadras
 else
 
 session("varQuadras") = varQuadras
 end if
 
 if session("varQuadras") = "" then
 session("varQuadras") = "quaqualquer"
 end if
 
'-------------------------------


'-----------Edícula-------------------------

dim varEdicula

varEdicula = request.querystring("txt_edicula")

 if varEdicula ="" then
 varEdicula = request.QueryString("varEdicula")
 session("varEdicula") = varEdicula
 else
 
 session("varEdicula") = varEdicula
 end if
 
 if session("varEdicula") = "" then
 session("varEdicula") = "ediqualquer"
 end if
 
'-------------------------------


dim varOcupacao

varOcupacao = request.querystring("txt_ocupacao")

 if varOcupacao ="" then
 varOcupacao = request.QueryString("varOcupacao")
 session("varOcupacao") = varOcupacao
 else
 
 session("varOcupacao") = varOcupacao
 end if
 
 if session("varOcupacao") = "" then
 session("varOcupacao") = "ocuqualquer"
 end if
 
'-------------------------------





'----------------------------Área Total------------------------------------
	
	dim varAreaTotal, varAreaTotal1,varAreaTotal2
 

varAreaTotal = request.QueryString("txt_area_total")

session("varAreaTotal")=varAreaTotal
  


if request.QueryString("txt_area_total")<>"" then
varAreaTotal = request.QueryString("txt_area_total")
 session("varAreaTotal") = varAreaTotal
 else
 
 session("varAreaTotal") = varAreaTotal
 end if
 
  if request.QueryString("varAreaTotal")<>"" then
  varAreaTotal = request.QueryString("varAreaTotal")
 session("varAreaTotal") = varAreaTotal
 else
 
 varAreaTotal = session("varAreaTotal") 
 end if 
	
 session("varAreaTotal1")=left(varAreaTotal,10)
   session("varAreaTotal2")=right(varAreaTotal,10)	
	
	
	varvarAreaTotal1 = session("varvarAreaTotal1")
	varvarAreaTotal2 = session("varvarAreaTotal2")
	
	if session("varAreaTotal") = "" then
	session("varAreaTotal") = "arequalquer"
	end if
	
	
	
	
	
	dim varCondominio, varCondominio1,varCondominio2
 

varCondominio = request.QueryString("txt_condominio")



session("varCondominio")=varCondominio



dim varStandby
 
 varStandby = request.QueryString("txt_standby")
 
if varStandby = "" then
varStandby = request.QueryString("varStandby")
session("varStandby")=varStandby
else
session("varStandby")=varStandby

end if


if varStandby = "" then
varStandby = "stanqualquer"
end if


dim varCondicoes
 

varCondicoes = request.QueryString("txt_Condicoes")

session("varCondicoes") = varCondicoes

if varCondicoes = "" then
varCondicoes = request.QueryString("varCondicoes")
session("varCondicoes")=varCondicoes
else
session("varCondicoes")=varCondicoes
 end if 
 
 if varCondicoes = "" then
 varCondicoes = "condiqualquer"
 session("varCondicoes")=varCondicoes
 end if
 


if request.QueryString("txt_condominio")<>"" then
varValor = request.QueryString("txt_condominio")
 session("varCondominio") = varCondominio
 else
 
 session("varCondominio") = varCondominio
 end if
 
  if request.QueryString("varCondominio")<>"" then
  varCondominio = request.QueryString("varCondominio")
 session("varCondominio") = varCondominio
 else
 
 varCondominio = session("varCondominio") 
 end if 
	
 session("varCondominio1")=left(varCondominio,10)
   session("varCondominio2")=right(varCondominio,10)	
	
	
	varCondominio1 = session("varCondominio1")
	varCondominio2 = session("varCondominio2")
	
	if session("varCondominio") = "" then
	session("varCondominio") = "conqualquer"
	end if
	




dim varNotFind



dim rs4,strSQL4,Conexao
   Set Conexao = Server.CreateObject("ADODB.Connection")
    Set rs4 = Server.CreateObject("ADODB.RecordSet")
	if session("varCidade2") <> "cqualquer" then
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = "&int(session("varCidade2"))&"  ORDER BY nome_combo2" 
	else
	strSQL4 = "SELECT combo2.id_combo2,combo2.id_combo1,combo2.nome_combo2,combo2.data_combo2,combo2.cidade_combo2  FROM combo2 where id_combo1 = 4 ORDER BY nome_combo2"
	end if
	
	Conexao.Open dsn
	
	
	
	 
rs4.CursorLocation = 3
rs4.CursorType = 3

rs4.ActiveConnection = Conexao
	
	
	
	rs4.Open strSQL4, Conexao




dim rs444,strSQL444
   
    Set rs444 = Server.CreateObject("ADODB.RecordSet")



 dim rs444Atendimento,strSQL444Atendimento
   
    Set rs444Atendimento = Server.CreateObject("ADODB.RecordSet")
	strSQL444Atendimento = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha where origem_franquia like '"&session("vOrigem_Franquia")&"' ORDER BY id" 
	
	
	rs444Atendimento.CursorLocation = 3
    rs444Atendimento.CursorType = 3

    rs444Atendimento.ActiveConnection = Conexao
	
	
	
	
	
	 rs444Atendimento.Open strSQL444Atendimento, Conexao



dim varAtendimento
varAtendimento = request.querystring("txt_atendimento")
if varAtendimento = "" then
varAtendimento = request.querystring("varAtendimento")
end if

session("varAtendimento") = varAtendimento




'------------------------selecionar tipo-------------------------


 dim rs444Tipo22,strSQL444Tipo22
   
    Set rs444Tipo22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Tipo22 = "SELECT tipo.id_tipo,tipo.tipo,tipo.data_tipo  FROM tipo  ORDER BY tipo ASC" 
	
	
	
	
	
	rs444Tipo22.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444Tipo22.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444Tipo22.ActiveConnection = Conexao
	
	
	
	
	
	 rs444Tipo22.Open strSQL444Tipo22, Conexao





'----------------------------------------------------------





%>




<html>
<head>
<title>Compradores</title>

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

function newWindow2(abrejanela2) {
   openWindow2 = window.open(abrejanela2,'openWin2','width=800,height=600,resizable=yes,scrollbars=yes')
   openWindow2.focus( )
   }

</SCRIPT>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow22(abrejanela22) {
   openWindow22 = window.open(abrejanela22,'openWin22','width=603,height=500,resizable=yes,scrollbars=yes')
   openWindow22.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2222(abrejanela2222) {
   openWindow2222 = window.open(abrejanela2222,'openWin2222','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow2222.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3121(abrejanela3121) {
   openWindow3121 = window.open(abrejanela3121,'openWin3121','width=620,height=500,resizable=yes,scrollbars=yes')
   openWindow3121.focus( )
   }

</SCRIPT>






<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3333(abrejanela3333) {
   openWindow3333 = window.open(abrejanela3333,'openWin3333','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow3333.focus( )
   }

</SCRIPT>



<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow30303(abrejanela30303) {
   openWindow30303 = window.open(abrejanela30303,'openWin30303','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow30303.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow5555(abrejanela5555) {
   openWindow5555 = window.open(abrejanela5555,'openWin5555','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow5555.focus( )
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
		moveObject(menu_id,position.x+100,position.y - vertical_offset);
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
<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (b2) 



{



var strValidNumber1_4="1234567890";
for (nCount=0; nCount < b3.SearchFor.value.length; nCount++) 
		{
strTempChar1_4=b3.SearchFor.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1 && b3.SearchWhere.value == "telefone" ) 
{
alert("Ao colocar seu telefone, digite apenas números!");
b3.SearchFor.focus();
b3.SearchFor.select();
return false;
}
}



}






</script>


<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber2 (b1) 



{



var strValidNumber1_4="1234567890";
for (nCount=0; nCount < b1.SearchFor.value.length; nCount++) 
		{
strTempChar1_4=b1.SearchFor.value.substring(nCount,nCount+1);
if (strValidNumber1_4.indexOf(strTempChar1_4,0)==-1 && b1.SearchWhere.value == "telefone" ) 
{
alert("Ao colocar seu telefone, digite apenas números!");
b1.SearchFor.focus();
b1.SearchFor.select();
return false;
}
}



}






</script>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow7777(abrejanela7777) {
   openWindow7777 = window.open(abrejanela7777,'openWin7777','width=600,height=500,resizable=yes,scrollbars=yes')
   openWindow7777.focus( )
   }

</SCRIPT>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow6666(abrejanela6666) {
   openWindow6666 = window.open(abrejanela6666,'openWin6666','width=700,height=500,resizable=yes,scrollbars=yes')
   openWindow6666.focus( )
   }

</SCRIPT>


<script language="javascript">
function funScroll()
{
window.scrollTo(0,500)

}		
</script>

<script language=javascript>
function confirmacao(){
 if (confirm("tem certeza que você quer excluir esse item?"))
  {
  return true;
  }
  else
  {
  return false;
  }
}
</script>



<%
'onLoad="funScroll()" || scrolling automático
%>

</head>
<body  topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">


<center>
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td><table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="675" border="0" bordercolor="#FFFFFF" cellspacing="0" cellpadding="0">
          <tr> 
            
                  <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis.asp" target="_blank">Im&oacute;veis</a></strong></font></div></td>
                  <td width="134" height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores.asp" target="_blank">Compradores</a></strong></font></div></td>
                  <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta.asp" target="_blank">Permuta</a></strong></font></div></td>
            
          <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
            <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta.asp" target="_blank">Proposta</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email.asp" target="_blank">Email</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow7777('procurar_avaliacao_corretor.asp')" style="color:#FFFFFF">Avaliação </a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ligar_urgente_comprador.asp" target="_blank" style="color:#FFFFFF">Ligar 
                urgente</a></strong></font></div></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imovel_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Imóveis 
                clicados</a></strong></font></div></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_contas_procuradas_corretor.asp" target="_blank" style="color:#FFFFFF">Contas 
                acessadas</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_imovel.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                imóvel</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_futuro_contato_comprador.asp" target="_blank" style="color:#FFFFFF">Fidelizar 
                compradores</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo02.asp" target="_blank" style="color:#FFFFFF">Captação 
                bloco</a></strong></font></div>
				<%else%>
				<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação 
                bloco</strong></font></div>
				
				
				<%end if%></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="form_via_codigo01.asp" target="_blank" style="color:#FFFFFF">Atendente 
                bloco</a></strong></font></div>
			<%else%>
			
			<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente 
                bloco</strong></font></div>
			
			<%end if%></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_financiamentos.asp" target="_blank" style="color:#FFFFFF">Financiamentos</a></strong></font></div>
			<%else%>
			<div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Financiamentos</strong></font></div>
			<%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp" target="_blank" style="color:#FFFFFF">Cidade</a></strong></font></div>
			  <%else%>
             
			  <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div>
			 
			  <%end if%></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro.asp" target="_blank" style="color:#FFFFFF">Bairro</a></strong></font></div>
			  <%else%>
			  <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div>
			  
              <%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_vila.asp" target="_blank" style="color:#FFFFFF">Vila</a></strong></font></div>
			  <%else%>
			  
			   <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div>
			  
              <%end if%></td>
            <td width="135" height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_comprador_clicado_corretor.asp" target="_blank" style="color:#FFFFFF">Compradores 
                Clicados</a></strong></font></div></td>
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_procurados.asp" target="_blank">Im&oacute;veis 
          procurados</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_referencia_procurados.asp" target="_blank">Refer&ecirc;ncias 
          procuradas</a></strong></font></div></td>
  </tr>
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_procurados.asp" target="_blank">Permutantes 
          procurados</a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_origem.asp" target="_blank">Origem</a></strong></font></div>
	  <%else%>
	  <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Origem</strong></font></div>
      <%end if%>
	  
	  </td>
            <td width="135" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_tipo.asp" target="_blank">Tipos de imóveis</a></strong></font></div>
			  <%else%>
              
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipos de imóveis</strong></font></div>
		<% end if %>  
		</td> 
		   
		    <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_enviado.asp" target="_blank">Emails 
                enviados </a></strong></font></div></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_oficial.asp" target="_blank">Proposta oficial 
                 </a></strong></font></div></td>
    
  </tr>
  
          <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_procurados.asp" target="_blank">Compradores procurados</a></strong></font></div></td>
            
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_imoveis.asp" target="_blank" style="color:#FFFFFF">corretores externos imóveis</a></strong></font></div>
			  <%else%>
              <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>corretores externos imóveis</strong></font></div>
			
			<%end if%> 
            </td>
            
            <td width="135" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_corretores_fora_compradores.asp" target="_blank" style="color:#FFFFFF">Corretores externos compradores</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Corretores externos compradores</strong></font></div>
			
			<%end if%> </td> 
		   
		    
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if session("permissao") = "6" and (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_visualiza_paginas.asp" target="_blank" style="color:#FFFFFF">Visualização de páginas</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Visualização página</strong></font></div>
			
			<%end if%></td>
            <td width="134" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <% if  (Lcase(session("vOrigem_Franquia"))) = (Lcase(session("vOrigem_Franquia02"))) then%>
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_de_fora.asp" target="_blank" style="color:#FFFFFF">Proposta 
                de Fora</a></strong></font></div>
			<%else%>
            <div align="center"><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>Proposta de fora</strong></font></div>
			
			<%end if%></td>
    
  </tr>
   <tr> 
            <td width="136" height="20" bgcolor="<%=claro%>" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_interno.asp" target="_blank">Email interno</a></strong></font></div></td>
            
            <td width="134" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;" > 
              <div align="center"></div>
			  </td>
            
            <td width="135" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td> 
		   
		    
            <td width="136" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td>
            <td width="134" height="20" bgcolor="#FFFFFF" style="color:#FFFFFF;border:1px solid #FFFFFF;"> 
              <div align="center"></div>
			</td>
    
  </tr>
  
</table></td>
  </tr>
  </table></td>
  </tr>
  
  
  <tr>
      <td height="80"><table width="800" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="133" height="80"><img src="simbol_comprador.jpg" width="133" height="80" border="0"></img></td>
            <td><table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="10" height="10" bgcolor="#4d4343"></td>
          </tr>
        </table> </div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="<%=escuro%>">Comprou 
                          com a Veja</font></strong></font></strong></font></td>
                      </tr>
                    </table></td>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
                                <td width="10" height="10" bgcolor="#f7e302"></td>
          </tr>
        </table></div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="<%=escuro%>">Suspenso</font></strong></font></td>
                      </tr>
                    </table></td>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
                                <td width="10" height="10" bgcolor="#1956C6"></td>
          </tr>
        </table></div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FFFFFF">s</font><font color="<%=escuro%>">Proposta</font></strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
                                <td width="10" height="10" bgcolor="#9a9090"></td>
          </tr>
        </table></div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="<%=escuro%>">Comprou 
                          com outro</font></strong></font></td>
                      </tr>
                    </table></td>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
                                <td width="10" height="10" bgcolor="#0fb5ab"></td>
          </tr>
        </table></div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#FFFFFF">s</font><font color="<%=escuro%>">Cliente 
                          inexistente </font></strong></font></td>
                      </tr>
                    </table></td>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
                                <td width="10" height="10" bgcolor="#a753f6"></td>
          </tr>
        </table></div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="<%=escuro%>">S&oacute; 
                          permuta </font></strong></font></td>
                      </tr>
                    </table></td>
                </tr>
				
				<tr>
                  <td width="200" height="20"><table width="200" height="20" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td width="20"><div align="center"><table width="10" height="10" border="0" cellspacing="0" cellpadding="0">
          <tr>
                                <td width="10" height="10" bgcolor="#61b4e8"></td>
          </tr>
        </table></div></td>
						<td width="10">&nbsp;</td>
                        <td width="180"><font color="#006699" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="<%=escuro%>">Comprador 
                          a contatar</font></strong></font></td>
                      </tr>
                    </table></td>
                  <td width="200" height="20">&nbsp;</td>
                  <td width="200" height="20">&nbsp;</td>
                </tr>
				
				
				
              </table></td>
          </tr>
        </table></td>
  </tr>
  
  <tr><td height="20"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>A 
          sua permissão é <%=session("permissao")%></strong></font></div></td></tr>
</table>
<br>

<%
 '----------------------------------------------------
 dim hora, dia, mes, ano

hora = hour(now())
dia = day(now())
mes = month(now())
ano = year(now())
  'Abrindo a tabela MARCAS!

dim Sql001
dim rs001

Sql001 = "SELECT * FROM cliente_online where data_full like '"&hora&"/"&dia&"/"&mes&"/"&ano&"' and atendimento like '"&session("nome_id")&"' ORDER BY cod_online ASC" 

Set rs001 = Server.CreateObject("ADODB.RecordSet")

	rs001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs001.ActiveConnection = Conexao3
	
	
	rs001.Open sql001, Conexao3



if not rs001.eof then 
                response.write "Veja Seus clientes online"
				   %>
				   
				   <table width="400" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="150"><iframe src="cliente_online01.asp" name="meio" width="400px" height="150px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>
				   
				   
				   <%
				   else
				  
	end if
'------------------------------------------------------------------------------
  
  
  %>
  
  
  
  <br>
  <%
 '----------------------------------------------------
 
  'Abrindo a tabela MARCAS!

dim Sql0022
dim rs0022

Sql0022 = "SELECT * FROM imoveis where data_atualizacao like '%"&dia&"/"&mes&"/"&ano&"%' and captacao like '"&session("nome_id")&"' and imovel_em_negociacao like '"&"Vendido por outros"&"' ORDER BY cod_imovel ASC" 

Set rs0022 = Server.CreateObject("ADODB.RecordSet")

	rs0022.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs0022.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs0022.ActiveConnection = Conexao3
	
	
	rs0022.Open sql0022, Conexao3



if not rs0022.eof then 
                
				   
				   %>
  <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"></font></div> 
  <%
				   else
				  
	end if
'------------------------------------------------------------------------------
  
%>
  <%
dim rs002
Set rs002 = Server.CreateObject("ADODB.RecordSet")



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

dim SQL002
if session("permissao") <> 6 then

SQL002 = "select  compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  from compradores where "
do until instr(vDataAtual, " ") = 0
		SQL002 = SQL002 & "data_ligar_urgente like '%" _
			& left(vDataAtual, instr(vDataAtual," ") - 1) & "%' or "
		vDataAtual = Right(vDataAtual, len(vDataAtual) - instr(vDataAtual," "))
	loop
	if len(vDataAtual) > 1 then
		SQL002 = SQL002 & "data_ligar_urgente like '%" & vDataAtual & "%' and atendimento like '"& Session("nome_id") &"' and (standby like 'comprador a contatar' or standby like 'comprador OK') "&" ORDER  BY data_atualizacao DESC"
	else
		SQL002 = left(SQL002, len(SQL002) - 4)
		SQL002 = SQL002&" and atendimento like '"& Session("nome_id") &"' and (standby like 'comprador a contatar' or standby like 'comprador OK') ORDER  BY data_atualizacao DESC"
	end if




else




SQL002 = "select  compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  from compradores where "
do until instr(vDataAtual, " ") = 0
		SQL002 = SQL002 & "data_ligar_urgente like '%" _
			& left(vDataAtual, instr(vDataAtual," ") - 1) & "%' or "
		vDataAtual = Right(vDataAtual, len(vDataAtual) - instr(vDataAtual," "))
	loop
	if len(vDataAtual) > 1 then
		SQL002 = SQL002 & "data_ligar_urgente like '%" & vDataAtual & "%' "&" and (standby like 'comprador a contatar' or standby like 'comprador OK') ORDER  BY data_atualizacao DESC"
	else
		SQL002 = left(SQL002, len(SQL002) - 4)
		SQL002 = SQL002&" and (standby like 'comprador a contatar' or standby like 'comprador OK')  ORDER  BY data_atualizacao DESC"
	end if







end if



rs002.Open SQL002, conexao, 1, 3


if not rs002.eof then

if (UCase(session("nome_id")) = UCase(rs002("atendimento"))) then

%>
  <table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="300"><iframe src="archive_ligar_urgente_comprador.asp" name="meio" width="800px" height="300px" frameborder="0" scrolling="yes"></iframe></td>
  </tr>
</table>


<%
end if

else

end if

rs002.close
set rs002 = nothing
%>

  
  <form name="doublecombo"  method="GET" onSubmit="return isValidDigitNumber(this);" action="archive_compradores.asp">
    <table width="1100" height="26" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=claro%>">
      <tr>
    <td width="120" height="25" align="right"><select name="combo1" onChange="javascript:atualizacarros(this.form);" class="inputBox" style="HEIGHT: 16px; WIDTH: 115px; background:<%=medio%>" >
        <option value="cqualquer" selected>Cidade</option>
        <% if not rs3.eof then %>
        <% While NOT Rs3.EoF %>
        <option value="<% = Rs3("id_combo1") %>" <%if session("varCidade2") <> "cqualquer" then %><%if rs3("id_combo1") = int(session("varCidade2")) then response.write "selected" end if %><%end if%>> 
        <% = Rs3("nome_combo1") %>
        </option>
        <% Rs3.MoveNext %>
        <% Wend %>
        <option value="cqualquer">qualquer uma</option>
        <%else%>
        <option value=""></option>
        <%end if%>
      </select></td>
    <td width="120"><font color="#FFFFFF">
      <select name="combo2" class="inputBox" style="HEIGHT: 16px; WIDTH: 122px; background:<%=medio%>">
        <option value="bqualquer" >Bairro/Região</option>
        <% if not rs4.eof then%>
        <% While NOT Rs4.EoF %>
        <option value="<% = Rs4("id_combo2") %>" <%if session("varBairro2") <> "bqualquer" then %><%if rs4("id_combo2") = int(session("varBairro2")) then response.write "selected" end if %><%end if%>> 
        <% = Rs4("nome_combo2") %>
        </option>
        <% Rs4.MoveNext %>
        <% Wend %>
        <option value="bqualquer">qualquer um</option>
        <% else %>
        <option value=""></option>
        <% end if %>
      </select>
      </font></td>
        <td width="120"><select name="example2" size="1" class="inputBox" id="example2" onChange="redirect2(this.options.selectedIndex)" style="HEIGHT: 16px; WIDTH: 120px; background:<%=medio%>">
            <option value="nqualquer">Negociação </option>
            <option value="nqualquer" >Qualquer um </option>
            <option  value="Aluguel">Aluguel </option>
            <option value="Compra">Compra</option>
            <% if session("varNegociacao") <> "nqualquer" and session("varNegociacao") <> ""  then %>
            <option value="<%=session("varNegociacao")%>" selected><%=session("varNegociacao")%></option>
            <%end if%>
          </select></td>
    <td width="100"><select name="txt_tipo" size="1"  class="inputBox" style="HEIGHT: 16px; WIDTH: 100px; background:<%=medio%>">
                 <% if session("varTipo") <> "tqualquer" and session("varTipo") <> ""  then %>
			<option value="<%=session("varTipo")%>" selected><%=session("varTipo")%></option>
			<%end if%>		 
				 
				  <option value="tqualquer">Tipo</option>
				   <option value="tqualquer">Qualquer um</option>
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
                 </select></td>
    <td width="70"><select name="txt_quartos" id="txt_quartos" class="inputBox" style="HEIGHT: 16px; WIDTH: 70px; background:<%=medio%>" >
       <% if session("varQuartos") <> "qqualquer" and session("varQuartos") <> "" then %>
			<option value="<%=session("varQuartos")%>" selected><%=session("varQuartos")%></option>
			<%end if%>
	   
	    <option value="qqualquer">Quartos</option>
        <option value="01">01</option>
        <option value="02">02</option>
        <option value="03">03</option>
        <option value="04">04</option>
        <option value="05">05</option>
        <option value="06">06</option>
        <option value="07">07</option>
        <option value="08">08</option>
        <option value="09">09</option>
      </select></td>
	  <td width="70"><select name="txt_vagas" id="txt_vagas" class="inputBox" style="HEIGHT: 16px; WIDTH: 70px; background:<%=medio%>" >
       <% if session("varVagas") <> "vgqualquer" and session("varVagas") <> "" then %>
			<option value="<%=session("varVagas")%>" selected><%=session("varVagas")%></option>
			<%end if%>
	    <option value="vgqualquer">Vagas</option>
        <option value="01">01</option>
        <option value="02">02</option>
        <option value="03">03</option>
        <option value="04">04</option>
        <option value="05">05</option>
        <option value="06">06</option>
        <option value="07">07</option>
        <option value="08">08</option>
        <option value="09">09</option>
      </select></td>
	  <td width="120">
	  <select name="stage22" size="1" class="inputBox" id="stage22" style="HEIGHT: 16px; WIDTH: 160px; background:<%=medio%>">
                  <option value="vqualquer">Valor</option>
                  <option value="vqualquer">Qualquer um</option>
                  <% if session("varNegociacao") = "Aluguel" then %>
			 <option value="<%=session("varValor")%>" selected><% if session("varValor") <> "vqualquer" then response.write FormatNumber(session("varValor1"),2)&" até "&FormatNumber(session("varValor2"),2) else response.write "Valor" end if%></option>
			<option value="0000000000 0000000200">Até 200,00</option>
                  <option value="0000000000 0000000500"> até 500,00</option>
                  <option value="0000000000 0000000750"> até 750,00</option>
                  <option value="0000000000 0000001000"> até 1000,00</option>
                  <option value="0000000000 0000001500"> até 1500,00</option>
                  <option value="0000000000 0000002000"> até 2000,00</option>
                  <option value="0000000000 0000002500"> até 2500,00</option>
                  <option value="0000000000 0000003000"> até 3000,00</option>
                  <option value="0000000000 0000003500"> até 3500,00</option>
                  <option value="0000000000 0000004000"> até 4000,00</option>
                  <option value="0000004001 1000000000">Acima de 4000,00</option>
               <%else%>
			   
			   <option value="<%=session("varValor")%>" selected><% if session("varValor") <> "vqualquer" then response.write FormatNumber(session("varValor1"),2)&" até "&FormatNumber(session("varValor2"),2) else response.write "Valor" end if%></option>
			  
			   <option value="0000000000 0000020000">Até 20.000,00</option>
                  <option value="0000000000 0000050000"> até 50.000,00</option>
                  <option value="0000000000 0000080000"> até 80.000,00</option>
                  <option value="0000000000 0000110000"> até 110.000,00</option>
                  <option value="0000000000 0000150000"> até 150.000,00</option>
                  <option value="0000000000 0000200000"> até 200.000,00</option>
                  <option value="0000000000 0000250000"> até 250.000,00</option>
                  <option value="0000000000 0000300000"> até 300.000,00</option>
                  <option value="0000000000 0000350000"> até 350.000,00</option>
                  <option value="0000000000 0000400000"> até 400.000,00</option>
                  <option value="0000000000 1000000000">Acima de 400.000,00</option>
			   <%end if%>
			   
			    </select>
	  
	  </td>
	  <td width="130"><select name="txt_atendimento" id="txt_atendimento" class="inputBox" style="HEIGHT: 16px; WIDTH: 130px; background:<%=medio%>">
                     <% if session("varAtendimento") <> "atqualquer" and session("varAtendimento") <> "" then %>
			<option value="<%=session("varAtendimento")%>" selected><%=session("varAtendimento")%></option>
			<%end if%>
					 
					 
					  <option value="atqualquer">Atendimento</option>
                      <% if not rs444Atendimento.eof then %>
                      <% While NOT rs444Atendimento.EoF %>
                      <option value="<% = rs444Atendimento("List_Name") %>"> 
                      <% = rs444Atendimento("List_Name") %>
                      </option>
                      <% rs444Atendimento.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value="não informado">não informado</option>
                      <%end if%>
                    </select></td>
					
					
        <td width="200"><select name="txt_condicoes" id="txt_condicoes" class="inputBox" style="HEIGHT: 16px; WIDTH: 200px; background:<%=medio%>" >
          	     <option value="condiqualquer">Condições de pagamento</option> 
				  <% if session("varCondicoes") <> "condiqualquer" and session("varCondicoes") <> "" then %>
            <option value="<%=session("varCondicoes")%>" selected><%=session("varCondicoes")%></option>
            <%end if%>
				  
				 <option value="condiqualquer">não informado</option>
			    <option value="à vista">à vista</option>
                <option  value="à vista e parte parcelado">à vista e parte parcelado </option>
                <option value="à vista e parte FGTS" >à vista e parte FGTS </option>
                <option  value="à vista , FGTS e financiamento">à vista , FGTS e financiamento" </option>
                <option value="FGTS" >FGTS </option>
                <option  value="financiamento">financiamento</option>
                <option value="parcelado" >parcelado </option>
                <option value="parte à vista mais imóvel">parte à vista mais imóvel</option>
          </select></td>
	  
        <td width="65" bgcolor="#FFFFFF"> </td>
  </tr>
      <tr> 
        <td width="120" align="right" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="120" bgcolor="#FFFFFF"><font color="#FFFFFF">&nbsp; </font></td>
        <td width="120" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="100" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="70" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="70" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="120" bgcolor="#FFFFFF">&nbsp; </td>
	    <td width="130" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="130" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="65" bgcolor="#FFFFFF">&nbsp; </td>
  
  </tr>
  <tr>
  
  
        <td width="120" align="right" height="25"><select name="txt_suites" id="txt_suites" class="inputBox" style="HEIGHT: 16px; WIDTH: 120px; background:<%=medio%>" >
        
		 <% if session("varSuites") <> "suiqualquer" and session("varSuites") <> "" then %>
			<option value="<%=session("varSuites")%>" selected><%=session("varSuites")%></option>
			<%end if%>
		
		<option value="suiqualquer">Suítes</option>
        <option value="não">não</option>
        <option value="sim">sim</option>
        
      </select></td>
        <td width="120"><select name="txt_piscina" id="txt_piscina" class="inputBox" style="HEIGHT: 16px; WIDTH: 120px; background:<%=medio%>" >
            <% if session("varPiscina") <> "pisqualquer" and session("varPiscina") <> "" then %>
            <option value="<%=session("varPiscina")%>" selected><%=session("varPiscina")%></option>
            <%end if%>
            <option value="pisqualquer">Piscina</option>
            <option value="não">não</option>
            <option value="sim">sim</option>
            
          </select></td>
        <td width="120"><select name="txt_portaria" id="txt_portaria" class="inputBox" style="HEIGHT: 16px; WIDTH: 120px; background:<%=medio%>" >
            <% if session("varPortaria") <> "porqualquer" and session("varPortaria") <> "" then %>
            <option value="<%=session("varPortaria")%>" selected><%=session("varPortaria")%></option>
            <%end if%>
            <option value="porqualquer">Portaria</option>
            <option value="não">não</option>
            <option value="sim">sim</option>
          </select></td>
        <td width="100"><select name="txt_area_total" id="txt_area_total" class="inputBox" style="HEIGHT: 16px; WIDTH: 100px; background:<%=medio%>" >
            <% if session("varAreaTotal") <> "arequalquer" and session("varAreaTotal") <> "" then %>
            <option value="<%=session("varAreaTotal")%>"> 
            <% if session("varAreaTotal") <> "arequalquer" then response.write int(session("varAreaTotal1"))&"m² até "&int(session("varAreaTotal2"))&"m²" else response.write "Área Total" end if%>
            </option>
            <%end if%>
            <option value="arequalquer" >Área Útil</option>
            <option value="0000000025 0000000050">25m² até 50m²</option>
            <option value="0000000050 0000000075">50m² até 75m²</option>
            <option value="0000000075 0000000090">75m² até 90m²</option>
            <option value="0000000090 0000000110">90m² até 110m²</option>
            <option value="0000000110 0000000150">110m até 150m²</option>
            <option value="0000000150 0000000200">150m² até 200m²</option>
            <option value="0000000200 0000000250">200m² até 250m²</option>
            <option value="0000000250 0000000300">250m² até 300m²</option>
            <option value="0000000300 0000000350">300m² até 350m²</option>
            <option value="0000000350 0000000400">350m² até 400m²</option>
            <option value="0000000400 0000000450">400m² até 450m²</option>
            <option value="0000000450 0000000500">450m² até 500m²</option>
            <option value="0000000500 1000000000">Acima de 500m²</option>
          </select></td>
        <td width="70"><select name="txt_quintal" id="txt_quintal" class="inputBox" style="HEIGHT: 16px; WIDTH: 70px; background:<%=medio%>" >
            <% if session("varQuintal") <> "quiqualquer" and session("varQuintal") <> "" then %>
            <option value="<%=session("varQuintal")%>"><%=session("varQuintal")%></option>
            <%end if%>
            <option value="quiqualquer" >Quintal</option>
            <option value="não">não</option>
            <option value="sim">sim</option>
          </select></td>
	    <td width="70"><select name="txt_quadras" id="txt_quadras" class="inputBox" style="HEIGHT: 16px; WIDTH: 70px; background:<%=medio%>" >
            <% if session("varQuadras") <> "quaqualquer" and session("varQuadras") <> "" then %>
            <option value="<%=session("varQuadras")%>"><%=session("varQuadras")%></option>
            <%end if%>
            <option value="quaqualquer">Quadras</option>
            <option value="não">não</option>
            <option value="sim">sim</option>
          </select></td>
	    <td width="120"><select name="txt_condominio" size="1" class="inputBox" id="txt_Condomino" style="HEIGHT: 16px; WIDTH: 160px; background:<%=medio%>">
            <option value="conqualquer" >Condomínio</option>
            <option value="conqualquer">Qualquer um</option>
            <option value="<%=session("varCondominio")%>" selected> 
            <% if session("varCondominio") <> "conqualquer" then response.write FormatNumber(session("varCondominio1"),2)&" até "&FormatNumber(session("varCondominio2"),2) else response.write "Condomínio" end if%>
            </option>
            <option value="0000000000 0000000050">Até 50,00</option>
            <option value="0000000000 0000000100"> até 100,00</option>
            <option value="0000000000 0000000150"> até 150,00</option>
            <option value="0000000000 0000000200"> até 200,00</option>
            <option value="0000000000 0000000250"> até 250,00</option>
            <option value="0000000000 0000000300"> até 300,00</option>
            <option value="0000000000 0000000350"> até 350,00</option>
            <option value="0000000000 0000000400"> até 300,00</option>
            <option value="0000000000 0000000450"> até 350,00</option>
            <option value="0000000000 0000000500"> até 500,00</option>
            <option value="0000000000 0000000750"> até 750,00</option>
            <option value="0000000000 0000001000"> até 1000,00</option>
            <option value="0000000000 0000001500"> até 1500,00</option>
            <option value="0000000000 0000002000"> até 2000,00</option>
            <option value="0000000000 0000002500"> até 2500,00</option>
            <option value="0000000000 0000003000"> até 3000,00</option>
            <option value="0000000000 0000003500"> até 3500,00</option>
            <option value="0000000000 0000004000">0 até 4000,00</option>
            <option value="0000004001 1000000000">Acima de 4000,00</option>
          </select> </td>
	    <td width="130">
<select name="txt_edicula" id="txt_edicula" class="inputBox" style="HEIGHT: 16px; WIDTH: 130px; background:<%=medio%>" >
            <% if session("varEdicula") <> "ediqualquer" and session("varEdicula") <> "" then %>
            <option value="<%=session("varEdicula")%>"><%=session("varEdicula")%></option>
            <%end if%>
            <option value="ediqualquer" >Edícula</option>
            <option value="não">não</option>
            <option value="sim">sim</option>
          </select></td>
	    <td width="200" ><select name="txt_standby" id="txt_standby" class="inputBox" style="HEIGHT: 16px; WIDTH: 200px; background:<%=medio%>" >
            <option value="stanqualquer" selected>Situação do cliente</option>
			<% if session("varStandby") <> "stanqualquer" and session("varStandby") <> "" then %>
            <option value="<%=session("varStandby")%>" selected><%=session("varStandby")%></option>
            <%end if%>
            <option value="comprador OK" >comprador OK</option>
            
            <option value="comprou com a Veja">comprou com a Veja</option>
            <option value="comprou com outro">comprou com outro</option>
            <option value="suspenso">suspenso</option>
            <option value="cliente inexistente">cliente inexistente</option>
            <option value="cliente para permuta">cliente para permuta</option>
            <option value="cliente com proposta">cliente com proposta</option>
            <option value="comprador OK" >comprador OK</option>
            <option value="comprador a contatar" >comprador a contatar</option>
          </select></td>
        <td width="65" bgcolor="#FFFFFF">&nbsp;</td>
  
  </tr>
  
      <tr> 
        <td width="120" height="25" align="right" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="120" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="120" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="100" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="70" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="70" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="120" bgcolor="#FFFFFF">&nbsp; </td>
	    <td width="130" bgcolor="#FFFFFF">&nbsp; </td>
	    <td width="200" bgcolor="#FFFFFF" >&nbsp;</td>
        <td width="65" bgcolor="#FFFFFF" >&nbsp;</td>
  
  </tr>
  
      <tr> 
        <td width="120" height="25" align="right" bgcolor="<%=claro%>">
<select name="txt_ocupacao" id="txt_ocupacao" class="inputBox" style="HEIGHT: 16px; WIDTH: 120px; background:<%=medio%>" >
            <% if session("varOcupacao") <> "ocuqualquer" and session("varOcupacao") <> "" then %>
            <option value="<%=session("varOcupacao")%>"><%=session("varOcupacao")%></option>
            <%end if%>
            <option value="ocuqualquer" >Ocupação</option>
            <option value="vago">vago</option>
            <option value="alugado">alugado</option>
            <option value="ocupado por terceiros">ocupado por terceiros</option>
            <option value="ocupado pelo proprietário">Ocupado pelo proprietário</option>
          </select></td>
        <td width="120" bgcolor="#FFFFFF"><table width="120" height="25" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td width="65" bgcolor="<%=claro%>"> 
                <div align="center"> 
                  <input name="submit2" type="submit" class="inputSubmit" style="background:<%=medio%>;" value="Buscar" width="80">
                </div></td>
              <td width="55">&nbsp;</td>
            </tr>
          </table></td>
        <td width="120" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="100" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="70" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="70" bgcolor="#FFFFFF">&nbsp;</td>
	    <td width="120" bgcolor="#FFFFFF">&nbsp; </td>
	    <td width="130" bgcolor="#FFFFFF">&nbsp; </td>
	    <td width="200" bgcolor="#FFFFFF" >&nbsp;</td>
        <td width="65" bgcolor="#FFFFFF">&nbsp;</td>
  
  </tr>
  
  
</table>
</form>

</center>


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
   











       
if request.querystring("combo1") = "" and (request.querystring("SearchWhere")<> "" or request.querystring("SearchFor")<>"" ) then		

SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK') ORDER BY cod_compradores DESC"


if session("SearchFor") <>"" and session("SearchWhere") = "cod"  then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where  origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  cod_compradores like '"& session("SearchFor") &"'  ORDER BY Cod_compradores   DESC"

end if

if session("SearchFor") = "" and session("SearchWhere") = "cod"  then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where  origem_franquia like '"&session("vOrigem_Franquia")&"'   ORDER BY Cod_compradores   DESC"

end if




if session("SearchFor") = "" and session("SearchWhere") = "Data" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK')  ORDER  BY cod_compradores ASC"
end if

if session("SearchFor") <> "" and session("SearchWhere") = "Data" then
SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  (standby like 'comprador a contatar' or standby like 'comprador OK') and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "data like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "data like '%" & session("SearchFor") & "%'"&" ORDER  BY data DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" ORDER  BY data DESC"
	end if

end if


if session("SearchFor") <> "" and session("SearchWhere") = "Data" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK') and data like '"&session("SearchFor")&"%'  ORDER  BY data DESC"
end if


if session("SearchFor") = "" and session("SearchWhere") = "proprietario" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK')  ORDER  BY cod_compradores ASC"

end if






'---------------------------especial proprietário-----------------------------




if session("SearchFor") <> "" and session("SearchWhere") = "proprietario" then

if session("permissao") <> "6" then

SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and   "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "nome like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "nome like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if

else


SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"'  and   "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "nome like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "nome like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if

end if


end if




'-------------------------------------------------------------------------











if session("SearchFor") <> "" and session("SearchWhere") = "endereco" then





SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "endereco like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "endereco like '%" & session("SearchFor") & "%'"
	else
		SQL = left(SQL, len(SQL) - 4)
	end if






end if

if session("SearchFor") <> "" and session("SearchWhere") = "atendimento" then

SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK')  ORDER  BY Cod_compradores DESC"


end if

if session("SearchFor") = "" and session("SearchWhere") = "atendimento" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK') ORDER BY cod_compradores DESC"
end if




if session("SearchFor") ="" and session("SearchWhere") = "telefone" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento ,compradores.origem_franquia from compradores  where origem_franquia like '"&session("vOrigem_Franquia")&"'  and (standby like 'comprador a contatar' or standby like 'comprador OK')   ORDER BY telefone   ASC"
end if

if session("SearchFor") <>"" and session("SearchWhere") = "telefone" then

if session("permissao") <> "6" then

SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"'  and   "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "telefone like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "telefone like '%" & session("SearchFor") & "%'"
		SQL = SQL & " or telefone02 like '%" & session("SearchFor") & "%' or telefone03 like '%" & session("SearchFor") & "%'"
	else
		SQL = SQL & " or telefone02 like '%" & session("SearchFor") & "%' or telefone03 like '%" & session("SearchFor") & "%'"
		SQL = left(SQL, len(SQL) - 4)
	end if

else


SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"'  and   "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "telefone like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "telefone like '%" & session("SearchFor") & "%'"
		SQL = SQL & " or telefone02 like '%" & session("SearchFor") & "%' or telefone03 like '%" & session("SearchFor") & "%'"
	else
		SQL = SQL & " or telefone02 like '%" & session("SearchFor") & "%' or telefone03 like '%" & session("SearchFor") & "%'"
		SQL = left(SQL, len(SQL) - 4)
	end if


end if




end if




 'if session("SearchFor") ="" and session("SearchWhere") = "StandBy" then
'SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  from compradores where standby = '"&"incluido"&"'   ORDER BY Cod_compradores   DESC"
'end if


 if session("SearchFor") ="" and session("SearchWhere") = "comprou com a Veja" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and standby = '"&"comprou com a Veja"&"'   ORDER BY Cod_compradores   DESC"
end if

 if session("SearchFor") ="" and session("SearchWhere") = "comprou com outro" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and standby = '"&"comprou com outro"&"'   ORDER BY Cod_compradores   DESC"
end if


 if session("SearchFor") ="" and session("SearchWhere") = "suspenso" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and standby = '"&"suspenso"&"'   ORDER BY Cod_compradores   DESC"
end if


 if session("SearchFor") ="" and session("SearchWhere") = "cliente inexistente" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and standby = '"&"cliente inexistente"&"'   ORDER BY Cod_compradores   DESC"
end if

 if session("SearchFor") ="" and session("SearchWhere") = "cliente com proposta" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and standby = '"&"cliente com proposta"&"'   ORDER BY Cod_compradores   DESC"
end if


 if session("SearchFor") ="" and session("SearchWhere") = "cliente para permuta" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  standby = '"&"cliente para permuta"&"'   ORDER BY Cod_compradores   DESC"
end if






'-------------------------------------fidelizacão--------------------------------





if session("SearchFor") = "" and session("SearchWhere") = "fidelizacao" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK')  ORDER  BY data_atualizacao DESC"
end if

if session("SearchFor") <> "" and session("SearchWhere") = "fidelizacao" then
SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  (standby like 'comprador a contatar' or standby like 'comprador OK') and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "data_futuro_contato like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "data_futuro_contato like '%" & session("SearchFor") & "%'"&" ORDER  BY data_atualizacao DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" ORDER  BY data_atualizacao DESC"
	end if

end if






'-------------------------------------------------------------------------------

'-----------------------standby--------------------------------------------------
if session("SearchFor") <> "" and session("SearchWhere") = "standby" then
SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "standby like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "standby like '%" & session("SearchFor") & "%'"&" ORDER  BY data_atualizacao DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" ORDER  BY data_atualizacao DESC"
	end if

end if





'-------------------------------------origem do comprador--------------------------------







if session("SearchFor") <> "" and session("SearchWhere") = "origem do comprador" then
SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  (standby like 'comprador a contatar' or standby like 'comprador OK') and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "origem like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "origem like '%" & session("SearchFor") & "%'"&" ORDER  BY cod_compradores DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" ORDER  BY cod_compradores DESC"
	end if

end if





if  session("SearchFor") <> "" and session("SearchWhere") = "comprador a contatar" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  standby like '"&"comprador a contatar"&"'  and atendimento like '"&session("SearchFor")&"'  ORDER BY Cod_compradores   DESC"
end if




if  session("SearchFor") = "" and session("SearchWhere") = "comprador a contatar" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  standby like '"&"comprador a contatar"&"'  ORDER BY Cod_compradores   DESC"
end if


'------------------------------------------------------------------------------------


'------------------------------------Data de atualização--------------------------

if session("SearchFor") = "" and session("SearchWhere") = "data de atualização" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK')  ORDER  BY data DESC"
end if




if session("SearchFor") <> "" and session("SearchWhere") = "data de atualização" then
SQL = "select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and  (standby like 'comprador a contatar' or standby like 'comprador OK') and "
do until instr(session("SearchFor"), " ") = 0
		SQL = SQL & "data_atualizacao like '%" _
			& left(session("SearchFor"), instr(session("SearchFor")," ") - 1) & "%' or "
		session("SearchFor") = Right(session("SearchFor"), len(session("SearchFor")) - instr(session("SearchFor")," "))
	loop
	if len(session("SearchFor")) > 1 then
		SQL = SQL & "data_atualizacao like '%" & session("SearchFor") & "%'"&" ORDER  BY data_atualizacao DESC"
	else
		SQL = left(SQL, len(SQL) - 4)
		SQL = SQL&" ORDER  BY data_atualizacao DESC"
	end if

end if


if session("SearchFor") <> "" and session("SearchWhere") = "data de atualização" then
SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"' and (standby like 'comprador a contatar' or standby like 'comprador OK') and data_atualizacao like '"&session("SearchFor")&"%'  ORDER  BY data DESC"
end if


'------------------------------------------------------------------------------------


else







'-----------------------------Cidade-----------------------------------
stringIndex = " where cod_compradores <>"&"0"&"" 

if  session("varCidade") <> "cqualquer" and session("varCidade") <> "" then
stringCidade = " and cidade='"& session("varCidade")&"'"
else
stringCidade = ""
end if
 '--------------------------Bairro----------------------------

if session("varBairro") <> "bqualquer" and session("varBairro") <> "" then
stringBairro = " and Bairro like '%"&session("varBairro")&"%'"
else
stringBairro = ""
end if

 '------------------------------------------------------------- 


'-------------------Negociação---------------------------

if  session("varNegociacao") <> "nqualquer" and session("varNegociacao") <> "" then
stringNegociacao = " and negociacao='"&session("varNegociacao")&"'"
else
stringNegociacao = ""
end if



'---------------------------Tipo------------------------------
dim stringTipo


if  session("varTipo") <> "tqualquer" and  session("varTipo") <> "" then
stringTipo = " and tipo like '%"&session("varTipo")&"%'"
else

stringTipo = ""
end if

'-------------------------------------------------------------





'---------------------------Quartos------------------------------


if  session("varQuartos") <> "qqualquer" and session("varQuartos") <> "" then
stringQuartos = " and quartos <="&int(session("varQuartos"))&""
else
stringQuartos = ""
end if

'---------------------------------------------------------------------------

'---------------------------Vagas------------------------------


if  session("varVagas") <> "vgqualquer" and  session("varVagas") <> "" then
stringVagas = " and vagas <="&int(session("varVagas"))&""
else

stringVagas = ""
end if











'---------------------------------Valor-----------------------------------



dim stringValor





if session("varValor") <> "vqualquer" and session("varValor") <> ""   then
stringValor = " and Valor >="& session("varValor1") &" and Valor <="& session("varValor2") &""
else
stringValor = ""
end if



'------------------------atendimento----------------------------





dim stringAtendimento
if  session("varAtendimento") <> "atqualquer" and  session("varAtendimento") <> "" then
stringAtendimento = " and atendimento='"&session("varAtendimento")&"'"
else

stringAtendimento = ""
end if



'------------------------------------------------------------------------------


'---------------------------------Condominio-----------------------------------


dim stringCondominio2


if  session("varCondominio") <> "conqualquer" and session("varCondominio") <> "" then
stringCondominio2 = "  and Condominio >="& session("varCondominio1") &" and Condominio <="& session("varCondominio2") &" "
else
stringCondominio2 = ""
end if

'---------------------------------Área Total-----------------------------------





dim stringAreaTotal2
'foi trocado por área construida, mas as variáveis continuam área total

if  session("varAreaTotal") <> "arequalquer" and session("varAreaTotal") <> "" then
stringAreaTotal2 = "  and area_construida >="& session("varAreaTotal1") &" and area_construida <="& session("varAreaTotal2") &""
else
stringAreaTotal2 = ""
end if












'-------------------------------Suítes-----------------------------------------


dim stringSuites2
 
if  session("varSuites") <> "suiqualquer" and session("varSuites") <> "não" and session("varSuites") <> "0" and session("varSuites") <>"00" and session("varSuites") <>"" then
stringSuites2 = " and suites <> ''  and suites <>'"&"não informado"&"' and suites <>'"&"0"&"' and suites <>'"&"00"&"' and suites IS NOT NULL  "




else

stringSuites2 = ""
end if


'--------------------------Piscina--------------------------------------
dim stringPiscina2
 
if  session("varPiscina") <> "pisqualquer" and session("varPiscina") <> "não" and session("varPiscina") <> "0" and session("varPiscina") <>"00" and session("varPiscina") <>"" then
stringPiscina2 = " and piscina <> ''  and piscina <>'"&"não informado"&"' and piscina <>'"&"0"&"' and piscina <>'"&"00"&"' and piscina IS NOT NULL"




else

stringPiscina2 = ""
end if






'--------------------------------------------------------------------------------



'--------------------------Portaria--------------------------------------
dim stringPortaria2
 
if  session("varPortaria") <> "porqualquer" and session("varPortaria") <> "não" and session("varPortaria") <> "0" and session("varPortaria") <>"00" and session("varPortaria") <>"" then
stringPortaria2 = " and portaria <> ''  and portaria <>'"&"não informado"&"' and portaria <>'"&"0"&"' and portaria <>'"&"00"&"' and portaria IS NOT NULL"




else

stringPortaria2 = ""
end if



'--------------------------Quintal--------------------------------------
dim stringQuintal2
 
if  session("varQuintal") <> "quiqualquer" and session("varQuintal") <> "não" and session("varQuintal") <> "0" and session("varQuintal") <>"00" and session("varQuintal") <>"" then
stringQuintal2 = " and quintal <> ''  and quintal <>'"&"não informado"&"' and quintal <>'"&"0"&"' and quintal <>'"&"00"&"' and quintal IS NOT NULL"




else

stringQuintal2 = ""
end if


'--------------------------Quadras--------------------------------------
dim stringQuadras2
 
if  session("varQuadras") <> "quaqualquer" and session("varQuadras") <> "não" and session("varQuadras") <> "0" and session("varQuadras") <>"00" and session("varQuadras") <>"" then
stringQuadras2 = " and quadras <> ''  and quadras <>'"&"não informado"&"' and quadras <>'"&"0"&"' and quadras <>'"&"00"&"' and quadras IS NOT NULL"




else

stringQuadras2 = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Edícula--------------------------------------
dim stringEdicula2
 
if  session("varEdicula") <> "ediqualquer" and session("varEdicula") <> "não" and session("varEdicula") <> "0" and session("varEdicula") <>"00" and session("varEdicula") <>"" then
stringEdicula2 = " and edicula <> ''  and edicula <>'"&"não informado"&"' and edicula <>'"&"0"&"' and edicula <>'"&"00"&"' and edicula IS NOT NULL"




else

stringEdicula2 = ""
end if



'--------------------------------------------------------------------------------

'--------------------------Standby--------------------------------------
dim stringStandby2
 
if  session("varStandby") <> "stanqualquer" and session("varStandby") <> ""  then
stringStandby2 = " and standby <> ''  and standby ='"&session("varStandby")&"'  and standby IS NOT NULL"




else

stringStandby2 = ""
end if



'--------------------------------------------------------------------------------


'--------------------------Condições de pagamento--------------------------------------
dim stringCondicoes2
 
if  session("varCondicoes") <> "condiqualquer" and session("varCondicoes") <> ""  then
stringCondicoes2 = " and condicoes_pagamento <> ''  and condicoes_pagamento ='"&session("varCondicoes")&"'  and condicoes_pagamento IS NOT NULL"




else

stringCondicoes2 = ""
end if



'-------------------Ocupação---------------------------

dim stringOcupacao02

if  session("varOcupacao") <> "ocuqualquer" then
stringOcupacao02 = " and ocupacao <> '' and ocupacao='"&session("varOcupacao")&"'"
else

stringOcupacao02 = ""
end if
'------------------------------------------------------------------------------





if session("permissao") <> "6" then



SQL ="SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  FROM compradores"&stringIndex&stringCidade&stringBairro&stringNegociacao&stringTipo&stringQuartos&stringVagas&stringValor&stringAtendimento&stringCondominio2&stringAreaTotal2&stringSuites2&stringPiscina2&stringPortaria2&stringQuintal2&stringQuadras2&stringEdicula2&stringStandby2&stringCondicoes2&stringOcupacao02&" and origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '"&session("nome_id")&"'   order by cod_compradores DESC" 


else



SQL ="SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.origem_franquia  FROM compradores"&stringIndex&stringCidade&stringBairro&stringNegociacao&stringTipo&stringQuartos&stringVagas&stringValor&stringAtendimento&stringCondominio2&stringAreaTotal2&stringSuites2&stringPiscina2&stringPortaria2&stringQuintal2&stringQuadras2&stringEdicula2&stringStandby2&stringCondicoes2&stringOcupacao02&" and origem_franquia like '"&session("vOrigem_Franquia")&"'  order by cod_compradores DESC" 



end if

'---------------------------------------------------------------------------



if request.querystring("combo1") = "" and request.querystring("SearchWhere")="" and request.querystring("varCidade") = "" and session("SearchFor") = "" then

SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '%"& Session("Admin_ID") &"%'  ORDER  BY cod_compradores DESC"



'if session("permissao") = "6" then

'SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"'   ORDER  BY cod_compradores DESC"

'end if

if session("permissao") = "3" then

SQL = "Select compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento  from compradores where origem_franquia like '"&session("vOrigem_Franquia")&"' and atendimento like '%"& "internet" &"%'  ORDER  BY cod_compradores DESC"


end if



session("SearchWhere") = "atendimento"
session("SearchFor") = session("Admin_ID")
end if



 



end if  

	


























%>
<form action="archive_compradores.asp?SearchFor=<%=SearchFor%>" onSubmit="return isValidDigitNumber2(this);" Method="GET" name="b1" >

<table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
<td bgcolor="#DAE3F0">
<table border="0" align="center" cellpadding="2" cellspacing="0" bgcolor="<%=claro%>">
          <tr>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>Procurar</b></font></td>
            <td bgcolor="<%=claro%>">
<input type="text" name="SearchFor" class="inputBox" style="background:<%=medio%>;" value="<%=SearchFor%>">
            </td>
            <td bgcolor="<%=claro%>"><font face="Verdana, arial" size="1"  color="#FFFFFF"><b>em</b></font></td>
            <td bgcolor="<%=claro%>">
	
	
	<% 
	if SearchWhere = "" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background: <%=medio%>;">
<option value="proprietario" selected >Comprador</option>
<option value="telefone"  >Telefone</option>
<option value="Data" >Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao" >Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>

</select>

<%
end if
%>

	
			
			
	
<!-------------------------------------------------- -->



<!-------------------------------------------------- -->

<% 
	if SearchWhere = "Data" then
	
	%>		
			

<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario">Comprador</option>
<option value="endereco"  >Telefone</option>
<option value="Data" selected>Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao" >Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>

<%
end if
%>

<!-- --------------------------------------------------------- -->

<% 
	if SearchWhere = "proprietario" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" selected>Comprador</option>
<option value="telefone"  >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento">Atendimento</option>
<option value="fidelizacao" >Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>

<%
end if
%>

<% 
	if SearchWhere = "telefone" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" selected>Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao" >Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>

<%
end if
%>

<% 
	if SearchWhere = "cod" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod" selected>Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao" >Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>

<%
end if
%>

<% 
	if SearchWhere = "atendimento" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod" selected>Código do comprador</option>
<option value="atendimento" selected>Atendimento</option>
<option value="fidelizacao" >Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>

<%
end if
%>





<% 
	if SearchWhere = "fidelizacao" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao" selected>Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>


<% 
	if SearchWhere = "comprou com a Veja" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja" selected>comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>


<% 
	if SearchWhere = "comprou com outro" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro" selected>comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>

<% 
	if SearchWhere = "suspenso" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso" selected>suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>


<% 
	if SearchWhere = "cliente inexistente" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente" selected>cliente inexistente</option>

<option value="cliente para permuta">cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>


<% 
	if SearchWhere = "cliente para permuta" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta" selected>cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>



<% 
	if SearchWhere = "cliente com proposta" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta" >cliente para permuta</option>
<option value="cliente com proposta" selected>cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>


<% 
	if SearchWhere = "origem do comprador" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta" >cliente para permuta</option>
<option value="cliente com proposta" >cliente com proposta</option>
<option value="origem do comprador" selected>origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>


<% 
	if SearchWhere = "comprador a contatar" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta" >cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" selected>comprador a contatar</option>
<option value="data de atualização">data de atualização</option>
</select>


<%
end if
%>

<% 
	if SearchWhere = "data de atualização" then
	
	%>		
			
<select name="SearchWhere" class="inputBox" style="background:<%=medio%>;">
<option value="proprietario" >Comprador</option>
<option value="telefone" >Telefone</option>
<option value="Data">Data</option>
<option value="cod">Código do comprador</option>
<option value="atendimento" >Atendimento</option>
<option value="fidelizacao">Data de fidelização</option>
<option value="comprou com a Veja">comprou com a Veja</option>
<option value="comprou com outro">comprou com outro</option>
<option value="suspenso">suspenso</option>
<option value="cliente inexistente">cliente inexistente</option>

<option value="cliente para permuta" >cliente para permuta</option>
<option value="cliente com proposta">cliente com proposta</option>
<option value="origem do comprador">origem do comprador</option>
<option value="comprador a contatar" >comprador a contatar</option>
<option value="data de atualização" selected>data de atualização</option>
</select>


<%
end if
%>




            </td>
            <td bgcolor="<%=claro%>">
<input type="submit" value="Buscar" class="inputSubmit" style="background:<%=escuro%>;"></td>
</tr>
</table>
</td>
</tr>
</table>
</form>

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
	
'RS.MaxRecords = 100
	
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
%>
<center>
<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Foram encontrados  <%=intRecordCount%> registros na busca.</strong></font>
</center>
<br>


          
 <form  Method="Post" name="Formulario" action="multi_excluir_compradores.asp?varCodCompradores=<%=varCodCompradores%>&SearchFor=<%=SearchFor%>&SearchWhere=<%=SearchWhere%>&page=<%=cInt(intPage)%>" >
  <table width="900" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);">
      </td>
      <td width="26" height="18" bgcolor="<%=claro%>"> 
        <% if  session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %>
        <input name="image" type="image" src="bt_mini_excluir02.jpg" alt="Excluir" width="26" height="22" border="0">
        <%else%>
        <img src="bt_mini_excluir02.jpg" alt="Excluir" width="26" height="22" border="0"></img> 
        <%end if%>
      </td>
      <td width="26" height="18" bgcolor="<%=claro%>"> 
        <% if   session("permissao") = "1" or session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %>
        <a href="javascript:newWindow2('verifica_tudo01.asp')"><img src="bt_mini_incluir01.jpg" alt="Incluir" width="26" height="22" border="0"></a> 
        <%else%>
        <img src="bt_mini_incluir01.jpg" alt="Incluir" width="26" height="22" border="0"> 
        <%end if%>
      </td>
      
	  
	  
	
 <td width="30" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;">&nbsp;</td>
	 
	  
	 
	  
	  
	  <td width="30" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;">&nbsp;</td>
	 
	  <td width="30" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;">&nbsp;</td>
	 
	  <td width="30" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;">&nbsp;</td>
	 
	 
<td width="80" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Eu 
          quero</strong></font></div></td>
	 
	 
	 
	  <td width="80" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Eu 
          tenho </strong></font></div></td>
	 
	 
	  <td width="43" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cod</strong></font></div></td>
	 
	 
	  <td width="40" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Indica</strong></font></div></td>
	 
	 
	  <td width="80" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Permuta</strong></font></div></td>
     
	 
	 
	 <td width="200" bgcolor="#000000" height="18" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></div></td>
	 
	 
	  <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div>
        <div align="center"></div></td>
      
	  
	  
	 
	 
	 
	 
	 <td width="20" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
     
	 
	  
      
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


if rs("StandBy") = "comprou com a Veja" then
color1 = "#4d4343"
end if

if rs("StandBy") = "comprou com outro" then
color1 = "#9a9090"
end if

if rs("StandBy") = "suspenso" then
color1 = "#f7e302"
end if

if rs("StandBy") = "cliente inexistente" then
color1 = "#0fb5ab"
end if

if rs("StandBy") = "cliente para permuta" then
color1 = "#a753f6"
end if

if rs("StandBy") = "cliente com proposta" then
color1 = "#1956C6"
end if
'cliente com proposta
'cliente para permuta

if rs("StandBy") = "comprador a contatar" then
color1 = "#61b4e8"
end if




dim varCodCompradores
%>




	<% session("page")=intPage%>
	<% varCodCompradores = rs("COD_compradores") %>
	<tr> 
      <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("COD_compradores")%>"></td>
      <td width="26" height="18" bgcolor="<%=claro%>"> 
        <% if  session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6"  then %>
        <a href="excluir_compradores.asp?varCodCompradores=<%=varCodCompradores%>&page=<%=cInt(intPage)%>&varCidade=<%=session("varCidade")%>&varCidade2=<%=session("varCidade2")%>&varBairro=<%=session("varBairro")%>&varBairro2=<%=session("varBairro2")%>&varNegociacao=<%=session("varNegociacao")%>&varTipo=<%=session("varTipo")%>&varQuartos=<%=session("varQuartos")%>&varVagas=<%=session("varVagas")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>&varAtendimento=<%=session("varAtendimento")%>&varOcupacao=<%=session("varOcupacao")%>&varCondominio=<%=session("varCondominio")%>&varCondominio1=<%=session("varCondominio1")%>&varCondominio2=<%=session("varCondominio2")%>&varAreaTotal=<%=session("varAreaTotal")%>&varAreaTotal1=<%=session("varAreaTotal1")%>&varAreaTotal2=<%=session("varAreaTotal2")%>&varSuites=<%=session("varSuites")%>&varPiscina=<%=session("varPiscina")%>&varPortaria=<%=session("varPortaria")%>&varQuintal=<%=session("varQuintal")%>&varQuadras=<%=session("varQuadras")%>&varEdicula=<%=session("varEdicula")%>&varStandby=<%=session("varStandby")%>&varCondicoes=<%=session("varCondicoes")%>" onclick="return confirmacao();"><img src="bt_mini_excluir01.jpg" alt="Excluir" width="26" height="22" border="0"></img></a> 
        <%else%>
        <img src="bt_mini_excluir01.jpg" alt="Excluir" width="26" height="22" border="0"></img> 
        <%end if%>
      </td>
      <td width="26" height="18" bgcolor="<%=claro%>"> 
        <% if    session("permissao") <> "3" and session("permissao") <> "4" and session("permissao") <> "5"  and session("permissao") <> "6"   then %>
        <% if  UCase(rs("atendimento")) <> UCase(Session("nome_id")) then%>
        <img src="bt_mini_atualizar01.jpg" alt="Atualizar" width="26" height="22" border="0"></img> 
        <%else%>
        <a href="javascript:newWindow2('visualizar_compradores33.asp?varCodCompradores=<%=varCodCompradores%>')"><img src="bt_mini_atualizar01.jpg" alt="Atualizar" width="26" height="22" border="0"></img></a> 
        <%end if%>
        <%else%>
        <% if  (UCase(rs("atendimento")) <> UCase("Internet") and session("permissao") = "3") then%>
        <img src="bt_mini_atualizar01.jpg" alt="Atualizar" width="26" height="22" border="0"></img> 
        <%else%>
        <a href="javascript:newWindow2('visualizar_compradores33.asp?varCodCompradores=<%=varCodCompradores%>')"><img src="bt_mini_atualizar01.jpg" alt="Atualizar" width="26" height="22" border="0"></img></a> 
        <% end if%>
        <%end if%>
      </td>
      
	   
      <td width="30" bgcolor="<% if rs("standby") = "comprador OK" then response.write "green" else response.write color1 end if%>" height="18" style="border:1px solid #FFFFFF;"></td>
	 
	  
	
	 
	 
	  <td width="30" valign="middle" align="center" height="18"  bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
       
	    <%
				' if rs("atendimento") <> session("nome_id") and session("permissao") <> "6" then 
				 
				  dim vPermissao004
	   vPermissao004 = "sim"
	   ' if rs("atendimento") <> session("nome_id") and session("permissao") <> "6" then 
		
		if vPermissao004 <> "sim" then
				 %>
	   
	   
	   <% else %>
	   
	   
	    <%







'---------------------------------------------------------------------------
dim rs444P
dim strSQL444P

    Set rs444P = Server.CreateObject("ADODB.RecordSet")
'se no cliente ou no servidor.


	strSQL444P = "SELECT imoveis.cod_imovel,imoveis.telefone,imoveis.telefone02,imoveis.telefone03 FROM imoveis where (telefone like '%"&rs("telefone")&"%' or telefone02 like '%"&rs("telefone")&"%' or telefone03 like '%"&rs("telefone")&"%')"
	
	
	
	
 
	
	
	
	
	
	 
	
	 
		
Rs444P.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
Rs444P.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.
 
	 
	 rs444P.Open strSQL444P,Conexao 
	 
	 
	 
	
	   
	   
     %>
        <%
	 
 if not rs444P.eof then

 %>
      
	    <% 
				   dim rs444Permuta,SQL444Permuta
 Set rs444Permuta = Server.CreateObject("ADODB.RecordSet")
 SQL444Permuta = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone like'"& rs("telefone")&"' order by cod_permuta DESC" 
	
	
	rs444Permuta.CursorLocation = 3
         rs444Permuta.CursorType = 3
           rs444Permuta.ActiveConnection = Conexao
	
	
	rs444Permuta.open SQL444Permuta,Conexao,2,1  
	
	dim varCod_Permuta006
	
			
	if  not rs444Permuta.eof and session("permissao") = "" then
	
	varCod_Permuta006 = rs444Permuta("cod_permuta") 
	
	vPerguntaPermuta = "sim"
				  
				 'while not rs444Permuta.eof 
				  %>
              
        <div align="left"><a href="javascript:newWindow5555('visualizar_permuta33.asp?varCodPermuta=<%=rs444Permuta("cod_permuta")%>')"><img src="icone_permuta01.jpg" width="26" height="22" border="0" align="middle" ID="info_icon_SAC3834"  onMouseOver="show_info_popup(this,'<%=rs444Permuta("cod_permuta")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs444Permuta("cod_permuta")%>')"></img></a> 
          <DIV STYLE="border: #000000 0px solid;  width: 270px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 1px; right: 1px;" CLASS="smalltext" ID="<%=rs444Permuta("cod_permuta")%>">
	   
	   <table width="580" border="0" cellspacing="0" cellpadding="0">
               <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Última atualização</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444Permuta("data_atualizacao")%></strong></font></td>
              </tr>
			   
			   
			   
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>C&oacute;digo 
                      da permuta</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444Permuta("cod_permuta")%></strong></font></td>
              </tr>
			 
			 
			 
			 
			    <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Meu 
                      nome &eacute;:</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<input name="txt_nome" value="<%=rs444Permuta("nome")%>" type="text" id="txt_nome" size="38" maxlength="200" align="left" class="inputBox" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>; "></td>
              </tr>
              
			 
              
			  
			  
             
			  
			  
			  
                
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atualmente 
                      tenho um im&oacute;vel na cidade de:</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("cidade_vend") = "cqualquer" then response.write "não informado" else response.write rs444Permuta("cidade_vend") end if %>
                    </strong></font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>No 
                      bairro: </strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("bairro_vend") = "bqualquer" then response.write "não informado" else response.write rs444Permuta("bairro_vend") end if %>
                    </strong></font></td>
              </tr>
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Na 
                      vila: </strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("vila_vend") = "vlqualquer" then response.write "não informado" else response.write rs444Permuta("vila_vend") end if %>
                    </strong></font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>do 
                      tipo: </strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("tipo_vend") = "tqualquer" then response.write "não informado" else response.write rs444Permuta("tipo_vend") end if %>
                    </strong></font></td>
              </tr>
			  
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("quartos_vend") = "qqualquer" then response.write "não informado" else response.write rs444Permuta("quartos_vend") end if %>
                    </strong></font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Com 
                      o seguinte n&uacute;mero de vagas na garagem</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("vagas_vend") = "vgqualquer" then response.write "não informado" else response.write rs444Permuta("vagas_vend") end if %>
                    </strong></font></td>
              </tr>
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>No 
                      valor de</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("valor_vend") = "vqualquer" then response.write "não informado" else response.write FormatNumber(rs444Permuta("valor_vend"),2) end if %>
                    </strong></font></td>
              </tr>
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Meu 
                      im&oacute;vel tem a seguinte descri&ccedil;&atilde;o</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> <textarea name="textarea" class="inputBox" id="textarea"  style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Permuta("descricao_vend")%></textarea></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pretendo 
                      morar na cidade de:</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("cidade_comp") = "cqualquer" then response.write "não informado" else  response.write rs444Permuta("cidade_comp") end if %>
                    </strong></font></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>No 
                      bairro: </strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("bairro_comp") = "bqualquer" then response.write "não informado" else  response.write rs444Permuta("bairro_comp") end if %>
                    </strong></font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Na 
                     vila: </strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">&nbsp;<font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"></strong> 
                    <%if rs444Permuta("vila_comp") = "vlqualquer" then response.write "não informado" else  response.write rs444Permuta("vila_comp") end if %>
                    </strong></font></td>
              </tr>
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quero 
                      trocar por um im&oacute;vel do tipo:</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("tipo_comp") = "tqualquer" then response.write "não informado" else  response.write rs444Permuta("tipo_comp") end if %>
                    </strong></font></td>
              </tr>
                
				
				 
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Com 
                      o seguinte n&uacute;mero de dormit&oacute;rios</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("quartos_comp") = "qqualquer" then response.write "não informado" else  response.write rs444Permuta("quartos_comp") end if %>
                    </strong></font></td>
              </tr>
			  
			    <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Com 
                      o seguinte n&uacute;mero de vagas na garagem</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("vagas_comp") = "vgqualquer" then response.write "não informado" else  response.write rs444Permuta("vagas_comp") end if %>
                    </strong></font></td>
              </tr>
			  
			  
			  
			    <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>No 
                      valor de</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    <%if rs444Permuta("valor_comp") = "vqualquer" then response.write "não informado" else  response.write FormatNumber(rs444Permuta("valor_comp"),2) end if %>
                    </strong></font></td>
              </tr>
				
				
				
              <tr>
                  <td width="290" bgcolor="<%=medio%>" height="100" style="border:1px solid #FFFFFF;" >
<div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Que 
                      tenha a seguinte descri&ccedil;&atilde;o</strong></font></div></td>
                  <td width="290" height="100" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao2" class="inputBox" id="txt_descricao2"  style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=medio%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Permuta("descricao_comp")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                       
                      </tr>
                    </table></td>
              </tr>
            </table>
	   
	   
	   
	   
	   
	   
	   
	   
	   </DIV>
	   
	   
	   
			   
			   
			   
			    <%
				  ' rs444Permuta.movenext
				  ' wend
				   %>
                <%else%>
				
				<%
				varCod_Permuta006 = "0"
				%>
              </div>
              <%end if%>
		
		
		 
        <%  end if
		 rs444P.close 
		  set rs444P = nothing %>
		  
		  
		  <% end if%>
      </td>
	  
	  
	  <td width="30" bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;">
	  
	   <%
	   dim vPermissao003
	   vPermissao003 = "sim"
	   ' if rs("atendimento") <> session("nome_id") and session("permissao") <> "6" then 
		
		if vPermissao003 <> "sim" then
		
		%>
	   
	   
	   <% else %>
	  
	  <%
	  
	    dim rs444Imovel,SQL444Imovel
 Set rs444Imovel = Server.CreateObject("ADODB.RecordSet")
 SQL444Imovel = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta  FROM imoveis where telefone like'"& rs("telefone")&"' or telefone02 like'"& rs("telefone")&"' or telefone03 like'"& rs("telefone")&"'  order by cod_imovel DESC"
	
	
	
	 rs444Imovel.CursorLocation = 3
     rs444Imovel.CursorType = 3

     rs444Imovel.ActiveConnection = Conexao3
	
	
	rs444Imovel.open SQL444Imovel,Conexao3,2,1  
	
	dim CidadeImovelFicha01
	dim BairroImovelFicha01
	
	
	dim semImovel
	
		dim varCod_imovel006
		
		semImovel = "notNothing"
		
			
	if  not rs444Imovel.eof then
	
	
	varCod_imovel006 = rs444Imovel("cod_imovel")
	
	
	vPerguntaImovel = "sim"
	
	
	
	            CidadeImovelFicha01 = rs444Imovel("cidade")
				BairroImovelFicha01 = rs444Imovel("bairro")
				  
				'While not rs444Imovel.eof  
				  %>
                
				
	   <% if rs("atendimento") <> session("nome_id") and session("permissao") <> "6" then %>
	   
	   <% else %>
				
				
				
        <div align="right"><a href="javascript:newWindow3333('visualizar_imovel33.asp?varCod_imovel=<%=rs444Imovel("cod_imovel")%>')"><img src="icone_imovel01.jpg" width="26" height="22" border="0"  align="middle" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rs444Imovel("cod_imovel")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs444Imovel("cod_imovel")%>')"></img></a> 
          <DIV STYLE="border: #000000 0px solid;  width: 570px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 200px; right: 200px;" CLASS="smalltext" ID="<%=rs444Imovel("cod_imovel")%>">
		
		
		<table width="580" border="0" cellspacing="0" cellpadding="0">
                
           
		     <tr>
                        
                <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                    de atualiza&ccedil;&atilde;o</strong></font></div></td>
                        
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444Imovel("data_atualizacao")%></strong></font></td>
              </tr>
		    <tr>
                        
                <td height="30" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Situação 
                    do imóvel</strong></font></div></td>
                        
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><% if rs444Imovel("imovel_em_negociacao") <> ""then response.write rs444Imovel("imovel_em_negociacao") else response.write "não informado" end if  %></strong></font></td>
              </tr>
		   
		     <tr>
                        
                <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Captação</strong></font></div></td>
                        
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><% if rs444Imovel("captacao") <> ""then response.write rs444Imovel("captacao") else response.write "não informado" end if  %></strong></font></td>
              </tr>
			 
			  <tr>
                        
                <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                  <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;ltimo 
                    email enviado</strong></font></div></td>
			 <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><% if rs444Imovel("dataLastEmail") <> "" then response.write rs444Imovel("dataLastEmail") else response.write "Nenhum email enviado" end if %></strong></font></td>
			 
			 <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Código 
                    do imóvel</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("cod_imovel")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                    de inclus&atilde;o</strong> </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("data")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			 
			 
			 
			 
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Proprietário</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("proprietario")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong> 
                    </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("telefone")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                
			  
			 
              
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email 
                    </strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("email")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
              
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("endereco")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_cidade" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("cidade")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro 
                    </strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_bairro" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("bairro")%>" size="38" maxlength="33" align="left">
                  </td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
                    <input name="txt_tipo" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%if rs444Imovel("Tipo") <> "tqualquer" then response.write rs444Imovel("Tipo") else response.write "qualquer um" end if  %>" size="38" maxlength="33" align="left">
                    </td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quartos</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("quartos")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vagas</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("vagas")%>" size="38" maxlength="33" align="left"></td>
              </tr>
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negociação</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("negociacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
			  
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Área 
                    Total</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("area_total")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Área 
                    Útil</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs444Imovel("area_construida")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Condomínio</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs444Imovel("condominio")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Suítes</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="" type="text" class="inputBox" id="" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<% if rs444Imovel("suites") <> "" then response.write rs444Imovel("suites") else response.write "0" end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
			  
			  
			  
			  
			  
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value=" <%if vValor <> "vqualquer" then%><%=FormatNumber(rs444Imovel("Valor"),2)%> <%else%>não informado<%end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
			 
			                           
			 
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Descrição 
                          do imóvel </strong></font></div></td>
                    </tr>
                    <tr> 
                        <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; " onKeyPress="return limitfield(this, 800)"><%=rs444Imovel("obs_imovel")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        
                      </tr>
                    </table></td>
              </tr>
            </table>
		
</DIV>
		<% end if%>		 
				 
				 
				 
				  <%

              ' rs444Imovel.movenext
			  ' wend

             %>
                  <%else%>
				  
				 <% 
				varCod_imovel006 = "0"
				
				semImovel = "nothing"
				
				
				CidadeImovelFicha01 = "não informado"
				BairroImovelFicha01 = "não informado"
				  
				  %>
				  
                </div>
                <%end if%></div>
	  
	  <% end if %>
       </td>
	 
	  <td width="30" bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;"> 
	  
	   <% if rs("atendimento") <> session("nome_id") and session("permissao") <> "6" then %>
	   
	   
	   <% else %>
        <a href="javascript:newWindow30303('visualizar_compradores33.asp?varCodCompradores=<%=rs("cod_compradores")%>')"><img src="icone_comprador01.jpg" width="26" height="22" border="0" align="middle" ID="info_icon_SAC3834" onMouseOver="show_info_popup(this,'<%=rs("cod_compradores")%>',35)" onMouseOut="hide_info_popup(this,'<%=rs("cod_compradores")%>')"></img></a> 
        <DIV STYLE="border: #000000 0px solid;  width: 570px; background-image: url(imovel10001.jpg); visibility: hidden; position: absolute; left: 200px; right: 200px;" CLASS="smalltext" ID="<%=rs("cod_compradores")%>">
		
		
		<table width="580" border="0" cellspacing="0" cellpadding="0">
                
             
			  <tr>
                        
              <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                  de atualiza&ccedil;&atilde;o</strong></font></div></td>
                        <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("data_atualizacao")%></strong></font></td>
              </tr>
			 
			 
			 <tr>
                        <td height="30" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atendente</strong></font></div></td>
                        <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("atendimento")%></strong></font></td>
              </tr>
			  
			  
			   <tr>
                        
              <td height="30" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Situa&ccedil;&atilde;o 
                  do cliente</strong></font></div></td>
                        <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs("standby")%></strong></font></td>
              </tr>
			 
			 
			  
			 
			 
			  <tr>
                        <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                  de inclus&atilde;o</strong></font></div></td>
                        
              <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
                <%=rs("data")%>
                </strong></font></td>
              </tr>
			  
			  
			   <tr>
                        
              <td height="20" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Acessos</strong></font></div></td>
                        
              <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
                <%if rs("acessos") <> "" then response.write rs("acessos") else response.write "0" end if %>
                </strong></font></td>
              </tr>
			  
			  
			  
			   <tr>
                        <td height="20" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                          <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;ltimo 
                  email enviado</strong></font></div></td>
                        <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6"  then %>
          <%if  UCase(rs("atendimento")) <> UCase(Session("nome_id")) then %>
          &nbsp; 
          <%else%>
          <a href="javascript:newWindow2('visualizar_lastemail.asp?varCodCompradores=<%=varCodCompradores%>')"> 
          <%if rs("dataLastEmail") <> "" then response.write rs("dataLastEmail") else response.write "nenhum email enviado" end if %>
          </a> 
          <%end if%>
          <%else%>
          <a href="javascript:newWindow2('visualizar_lastemail.asp?varCodCompradores=<%=varCodCompradores%>')"> 
          <%if rs("dataLastEmail") <> "" then response.write rs("dataLastEmail") else response.write "nenhum email enviado" end if %>
          </a> 
          <%end if%></strong></font></td>
              </tr>
			 
			 <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>C&oacute;digo 
                  do comprador</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs("cod_compradores")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Uacute;tima 
                  atualiza&ccedil;&atilde;o</strong> </font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs("data_atualizacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			 
			 
			 
			 
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Meu 
                  nome</strong> </font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario" type="text" class="inputBox" id="txt_proprietario" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs("nome")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
             
			  <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_telefone" type="text" class="inputBox" id="txt_telefone" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs("telefone")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                
			  
			 
              
			   <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Email</strong></font></div></td>
                <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_email" type="text" class="inputBox" id="txt_email" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs("email")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
              
			   <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o</strong></font></div></td>
                <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_endereco" type="text" class="inputBox" id="txt_endereco" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs("endereco")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  
			  
			  
               
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade 
                  que estou interessado</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_cidade" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs("cidade")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro 
                  que estou interessado</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <input name="txt_bairro" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs("bairro")%>" size="38" maxlength="33" align="left">
                  </td>
              </tr>
              <tr>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo 
                  do im&oacute;vel desejado</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
                    <input name="txt_tipo" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%if rs("Tipo") <> "tqualquer" then response.write rs("Tipo") else response.write "qualquer um" end if  %>" size="38" maxlength="33" align="left">
                    </td>
              </tr>
               
             
                <tr> 
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;meros 
                  de quartos do im&oacute;vel desejado</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs("quartos")%>" size="38" maxlength="33" align="left"></td>
              </tr>
			  
			  <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;meros 
                  de vagas do im&oacute;vel desejado</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_quartos" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value="<%=rs("vagas")%>" size="38" maxlength="33" align="left"></td>
              </tr>
              
               
              <tr>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o 
                  que eu quero</strong></font></div></td>
                  <td bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txt_negociacao" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=claro%>;" value="<%=rs("negociacao")%>" size="38" maxlength="33" align="left"></td>
              </tr>
                <tr> 
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Faixa 
                  de pre&ccedil;o que eu quero</strong></font></div></td>
                  <td bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor" type="text" class="inputBox" id="txt_proprietario3" style="font-weight: bold;font-size:12;border-color : <%=medio%>;color:#FFFFFF;HEIGHT: 18px; WIDTH: 290px; background:<%=medio%>;" value=" <%if vValor <> "vqualquer" then%><%=FormatNumber(rs("Valor"),2)%> <%else%>não informado<%end if%>" size="38" maxlength="33" align="left"></td>
              </tr>
             
			 
			                           
			 
              <tr>
                <td width="290" height="100" style="border:1px solid #FFFFFF;" ><table width="290" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                        
                    <td width="290" height="18" bgcolor="<%=claro%>" style="border-bottom: 2px solid #FFFFFF;"> 
                      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Aqui 
                        tem a descri&ccedil;&atilde;o do im&oacute;vel que eu 
                        quero</strong></font></div></td>
                    </tr>
                    <tr> 
                        
                    <td width="290" height="82" bgcolor="<%=medio%>" >&nbsp;</td>
                    </tr>
                  </table></td>
                  <td width="290" height="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;" ><textarea name="txt_descricao" class="inputBox" id="txt_descricao" style="font-weight: bold;font-size:12;border-color : <%=claro%>;color:#FFFFFF;HEIGHT: 100px; WIDTH: 290px; background:<%=claro%>; " onKeyPress="return limitfield(this, 800)"><%=rs("descricao")%></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                       
                      </tr>
                    </table></td>
              </tr>
            </table>
		
</DIV>
	  
	<% end if%>  
	  
	  </td>
	 
	  


<td width="80" bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%if rs("tipo") <> "tqualquer" then response.write  rs("tipo") else response.write "qualquer um" end if%></font></div></td>
	 
	 
	 <td width="80" bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;"><%
 '---------Selecionar permutante pelo telefone------------------------------------------------
		   
		     dim rs202Imoveis,SQL444Imoveis202
 Set rs202Imoveis = Server.CreateObject("ADODB.RecordSet")
 SQL444Imoveis202 = "SELECT imoveis.cod_imovel,imoveis.tipo,imoveis.telefone,imoveis.telefone02,imoveis.telefone03 FROM imoveis where (telefone like '"&rs("telefone")&"' or telefone02 like '"&rs("telefone")&"' or telefone03 like '"&rs("telefone")&"')"
	
	
	rs202Imoveis.CursorLocation = 3
         rs202Imoveis.CursorType = 3
           rs202Imoveis.ActiveConnection = Conexao3
	
	
	rs202Imoveis.open SQL444Imoveis202,Conexao3,2,1  
	
			
	if  not rs202Imoveis.eof then
		%>
		
		 <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rs202Imoveis("tipo") <> "" then response.write rs202Imoveis("tipo") else response.write "não informado" end if %></font>
		
		<%   
		   
end if

 rs202Imoveis.close
  
  set rs202Imoveis = nothing 

%></td>
	  
	  <td width="43"  bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("cod_compradores")%></font></div></td>
	  
	   
      <td width="40" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
         <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%



'------------------------Cidade---------------------------









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



dim stringCondominio3


Porcentual02 = int(rs("condominio"))*10/100
   


   vCondominioMenor = int(rs("condominio")) - int(Porcentual02)
   vCondominioMaior = int(rs("condominio")) + int(Porcentual02)




if  int(rs("condominio")) <> 0 and rs("condominio") <> ""  then

'stringCondominio2 = " and Condominio >="& vCondominioMenor &" and Condominio <="& vCondominioMaior &""
'stringCondominio3 = "  and Condominio <="& vCondominioMaior &""
'stringCondominio2 = ""
stringCondominio3 = ""
else
stringCondominio3 = ""
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

'stringStandby = "  and (imovel_em_negociacao like  '"&"Suspenso"&"' or imovel_em_negociacao like  '"&"imóvel OK"&"' or imovel_em_negociacao like  '"&"Imóvel a recaptar"&"') "

stringStandby = "  and ( imovel_em_negociacao like  '"&"imóvel OK"&"' ) "




'---------------------------------------------------------------------------

    Set rs444 = Server.CreateObject("ADODB.RecordSet")
'se no cliente ou no servidor.


	strSQL444 = "SELECT imoveis.cod_imovel FROM imoveis"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringCondominio3&stringAreaConstruida&stringSuites&stringPiscina&stringPortaria&stringQuintal&stringQuadras&stringEdicula&stringOcupacao&stringStandby
	
	
	
	
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
 
 

'response.write varIndicacaoAreaConstruida







%> 
      </font></div></td>
	  
	   
      <td width="80" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
        <%
		   
		   '---------Selecionar permutante pelo telefone------------------------------------------------
		   
		     dim rs202,SQL444Permuta202
 Set rs202 = Server.CreateObject("ADODB.RecordSet")
 SQL444Permuta202 = "SELECT permuta.cod_permuta,permuta.nome,permuta.telefone,permuta.email,permuta.cidade_vend,permuta.bairro_vend,permuta.endereco_vend,permuta.tipo_vend,permuta.descricao_vend,permuta.cidade_comp,permuta.bairro_comp,permuta.tipo_comp,permuta.descricao_comp,permuta.cod_imovel,permuta.link_imovel,permuta.foto_imovel,permuta.data,permuta.quartos_comp,permuta.quartos_vend,permuta.valor_comp,permuta.valor_vend,permuta.atendimento,permuta.data_atualizacao,permuta.vila_vend,permuta.vila_comp,permuta.vagas_vend,permuta.vagas_comp,permuta.cod_comprador,permuta.standby,permuta.datalastemail,permuta.textolastemail,permuta.dados_confidenciais  FROM permuta where telefone='"& rs("telefone")&"' order by cod_permuta DESC" 
	
	
	rs202.CursorLocation = 3
         rs202.CursorType = 3
           rs202.ActiveConnection = Conexao3
	
	
	rs202.open SQL444Permuta202,Conexao3,2,1  
	
			
	if  not rs202.eof then
		   
		   
		   
		   
		   
		   
'------------------------Sua Cidade--------------------------

stringIndex202 = " where cod_permuta<>"&"0"&""
 
 
 
  if   rs202("cidade_vend") = "não informado" or rs202("cidade_vend") = "" or rs202("cidade_vend") = "cqualquer" or  rs202("cidade_vend") = "qualquer um" then
	stringCidadeVend202 = ""
 else

stringCidadeVend202 = " and (Cidade_comp='"&rs202("cidade_vend")&"' or Cidade_comp='"&"não informado"&"' or Cidade_comp='"&"cqualquer"&"' or Cidade_comp='"&"qualquer um"&"')"

end if	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend202

 if   rs202("bairro_vend") = "não informado" or rs202("bairro_vend") = "" or rs202("bairro_vend") = "bqualquer" or  rs202("bairro_vend") = "qualquer um" then
	stringBairroVend202 = ""
 else
'stringBairroVend = ""
stringBairroVend202 = " and (Bairro_comp like'%"&rs202("bairro_vend")&"%' or Bairro_comp like'%"&"não informado"&"%' or Bairro_comp like '%"&"bqualquer"&"%'  or Bairro_comp like'%"&"qualquer um"&"%')"

end if


'--------------------------Sua Vila--------------------------------

dim stringVilaVend202

'" and (Vila_comp='"&rs("vila_vend")&"' or Vila_comp='"&"não informado"&"' or Vila_comp='"&"vlqualquer"&"' or Vila_comp='"&"qualquer um"&"' )"

 if   rs202("vila_vend") = "não informado" or rs202("vila_vend") = "" or rs202("vila_vend") = "vlqualquer" or rs202("vila_vend") = "qualquer um" then
	stringVilaVend202 =  ""
 else

stringVilaVend202 = ""

end if






 '--------------------------Tipo do seu imóvel------------------------
 
 
 dim stringTipoVend202
 
 
 if rs202("tipo_vend") = "não informado" or rs202("tipo_vend") = "" or rs202("tipo_vend") = "tqualquer" or rs202("tipo_vend") = "qualquer um"  then

stringTipoVend202 = ""

else
stringTipoVend202 = " and Tipo_comp like '%"&rs202("tipo_vend")&"%'"
 
 end if


 
 '-----------------------Número de quartos do seu imóvel-----------------
 
 
 
 
 dim stringQuartosVend202
 
 
 

stringQuartosVend202 = " and Quartos_comp <="&int(rs202("quartos_vend"))&""

 


 '-----------------------Número de Vagas do seu imóvel-----------------
 
 
 
 
 dim stringVagasVend202
 
 
 



stringVagasVend202 = " and vagas_comp <="&int(rs202("vagas_vend"))&""

 




 
 
 
 
 '-----------------------------Valor de venda do seu imóvel----------------
 
 
 
dim PorcentualVend202

dim vValorMenorVend202
dim vValorMaiorVend202

PorcentualVend202 = int(rs202("valor_vend"))*20/100

   


   vValorMenorVend202 = int(rs202("valor_vend")) - int(PorcentualVend202)
   vValorMaiorVend202 = int(rs202("valor_vend")) + int(PorcentualVend202)

 
 
 
 
 
	 dim stringValorVend202
  
	
	
	
	'stringValorVend202 = " and Valor_comp >="&  vValorMenorVend202 &" and Valor_comp <="& vValorMaiorVend202&""
 
     stringValorVend202 = " and Valor_comp >="&  int(vValorMenorVend202) &" "
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp202
  if rs202("cidade_comp")="não informado" or rs202("cidade_comp")="" or rs202("cidade_comp")="cqualquer" or rs202("cidade_comp") = "qualquer um" then
	stringCidadeComp202 = ""
	else
	
	stringCidadeComp202 = " and Cidade_vend ='"& rs202("cidade_comp") &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp202

	if rs202("bairro_comp") = "não informado" or  rs202("bairro_comp") = "" or  rs202("bairro_comp") = "bqualquer" or rs202("bairro_comp") = "qualquer um" then
	
	
	
	
	
	stringBairroComp202 = ""
	
	
	
	
	else
	
	
	
	'stringBairroComp = " and Bairro_vend ='"& rs("bairro_comp") &"'"
	
	
	
	
 
dim Numero_Indicacoes202
dim Numero_Indicacoes02202




Numero_Indicacoes202 = 0
Numero_Indicacoes02202 = 0


dim soma02202
dim soma202

soma202 = 0
soma02202 = 0

dim Variavel202
dim Retorno202
dim contar202
Variavel202 = rs202("bairro_comp")
Retorno202 = Split(rs202("bairro_comp"),", ")

contar202=0

dim stringBairro3202
dim stringBairro4202
dim stringBairro5202

for contar202=0 to UBound(Retorno202)

stringBairro3202 = "and ( "
stringBairro4202 = " Bairro_vend='"&Retorno202(contar202)&"'or  " &stringBairro4202

stringBairro5202 = " cod_permuta=0)"


stringBairroComp202 = stringBairro3202&stringBairro4202&stringBairro5202



next


stringBairro3202 = ""
stringBairro4202 = ""
stringBairro5202 = ""

	
	
	

	
	
	end if
	
	
	
	
	 '----------------------------Vila pretendida---------------------------------
 
 'and Vila_vend ='"& rs("vila_comp") &"'
	 dim stringVilaComp202

	if rs202("vila_comp") <> "não informado" and rs202("vila_comp") <> "" and rs202("vila_comp") <> "vlqualquer" and rs202("vila_comp") <> "qualquer um" then
	stringVilaComp202 = ""
	else
	
	stringVilaComp202 = ""
	end if
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	' dim stringTipoComp
  'if rs("tipo_comp")="não informado" or rs("tipo_comp")="" or rs("tipo_comp")="tqualquer" or rs("tipo_comp") = "qualquer um" then
	'stringTipoComp = ""
	'else
	
	
	'stringTipoComp = " and Tipo_vend ='"& rs("tipo_comp")&"'"
	'end if
	
	
	
	'--------------------------Tipo----------------------------

if rs202("tipo_comp") <> "qualquer um" and rs202("tipo_comp") <> "não informado" then




 
dim Numero_IndicacoesTipoComp202
dim Numero_Indicacoes02TipoComp202




Numero_IndicacoesTipoComp202 = 0
Numero_Indicacoes02TipoComp202 = 0


dim soma02TipoComp202
dim somaTipoComp202

somaTipoComp202 = 0
soma02TipoComp202 = 0

dim VariavelTipoComp202
dim RetornoTipoComp202
dim contarTipoComp202
VariavelTipoComp202 =  rs202("tipo_comp")
RetornoTipoComp202 = Split(rs202("tipo_comp"),", ")

contarTipoComp202=0

dim stringTipo3Comp202
dim stringTipo4Comp202
dim stringTipo5Comp202

for contarTipoComp202=0 to UBound(RetornoTipoComp202)

stringTipo3Comp202 = "and ( "
stringTipo4Comp202 = " tipo_vend='"&RetornoTipoComp202(contarTipoComp202)&"'or  " &stringTipo4Comp202

stringTipo5Comp202 = " cod_permuta=0)"


stringTipo2Comp202 = stringTipo3Comp202&stringTipo4Comp202&stringTipo5Comp202







next

stringTipo3Comp202 = ""
stringTipo4Comp202 = ""
stringTipo5Comp202 = ""


else
stringTipo2Comp202 = ""
end if

	
	
	
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp202
  
	
	stringQuartosComp202 = " and Quartos_vend >="& int(rs202("quartos_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 '-----------------------------------Vagas Pretendidas---------------------------------
 
 
 
 
	 dim stringVagasComp202
 
	
	stringVagasComp202 = " and vagas_vend >="& int(rs202("vagas_comp")) &""
	
	
 
 
 '-----------------------------------------------------------------------
 
 
 
 
 
 
 
 '----------------------------Valor pretendido----------------------------

dim PorcentualComp202

dim vValorMenorComp202
dim vValorMaiorComp202

PorcentualComp202 = int(rs202("valor_comp"))*20/100

   


   vValorMenorComp202 = int(rs202("valor_comp")) - int(PorcentualComp202)
   vValorMaiorComp202 = int(rs202("valor_comp")) + int(PorcentualComp202)


	 dim stringValorComp202
  
	
	
	'stringValorComp202 = " and Valor_vend >="& vValorMenorComp202 &" and Valor_vend <="& vValorMaiorComp202 &""
	
	stringValorComp202 = "  and Valor_vend <="& int(vValorMaiorComp202) &""
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
	'strSQL444 = "SELECT * FROM permuta"
	'&stringVilaVend
	'&stringVilaComp
	varIndicacaoCodigo202=rs202("cod_permuta")
	
	strSQL444202 = "SELECT permuta.cod_permuta   FROM permuta"&stringIndex202&stringCidadeVend202&stringBairroVend202&stringVilaVend202&stringTipoVend202&stringQuartosVend202&stringVagasVend202&stringValorVend202&stringCidadeComp202&stringBairroComp202&stringVilaComp202&stringTipo2Comp202&stringQuartosComp202&stringVagasComp202&stringValorComp202&" and standby <> 'incluido' and cod_permuta not like "&varIndicacaoCodigo202
	
	
	 varIndicacaoCidadeVend202=rs202("cidade_vend")
 varIndicacaoBairroVend202=rs202("bairro_vend")
 varIndicacaoVilaVend202=rs202("vila_vend")
 varIndicacaoQuartosVend202=rs202("quartos_vend")
 varIndicacaoVagasVend202=rs202("vagas_vend")
 varIndicacaoValorVend202=rs202("valor_vend")
 varIndicacaoTipoVend202=rs202("tipo_vend")


 varIndicacaoCidadeComp202=rs202("cidade_comp")
 varIndicacaoBairroComp202=rs202("bairro_comp")
 varIndicacaoVilaComp202=rs202("vila_comp")
 varIndicacaoQuartosComp202=rs202("quartos_comp")
 varIndicacaoVagasComp202=rs202("vagas_comp")
 varIndicacaoValorComp202=rs202("valor_comp")
 varIndicacaoTipoComp202=rs202("tipo_comp")
	
	
 dim rs444202
 Set rs444202 = Server.CreateObject("ADODB.RecordSet")	
	
	 
rs444202.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444202.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444202.ActiveConnection = Conexao3
	 
	 rs444202.Open strSQL444202,Conexao3 
	   
     %>
        <% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6" then %>
        <div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow2('indicacao_permuta22.asp?varIndicacaoCidadeVend=<%=varIndicacaoCidadeVend202%>&varIndicacaoBairroVend=<%=varIndicacaoBairroVend202%>&varIndicacaoVilaVend=<%=varIndicacaoVilaVend202%>&varIndicacaoTipoVend=<%=varIndicacaoTipoVend202%>&varIndicacaoQuartosVend=<%=varIndicacaoQuartosVend202%>&varIndicacaoVagasVend=<%=varIndicacaoVagasVend202%>&varIndicacaoValorVend=<%=varIndicacaoValorVend202%>&varIndicacaoCidadeComp=<%=varIndicacaoCidadeComp202%>&varIndicacaoBairroComp=<%=varIndicacaoBairroComp202%>&varIndicacaoVilaComp=<%=varIndicacaoVilaComp202%>&varIndicacaoTipoComp=<%=varIndicacaoTipoComp202%>&varIndicacaoQuartosComp=<%=varIndicacaoQuartosComp202%>&varIndicacaoVagasComp=<%=varIndicacaoVagasComp202%>&varIndicacaoValorComp=<%=varIndicacaoValorComp202%>&varIndicacaoCodigo=<%=varIndicacaoCodigo202%>')"><strong><%=rs444202.RecordCount%></strong></a></font></div>
          <%else%>
          <div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=rs444202.RecordCount%></strong></font></div>
          <%end if%>
          <%
	 
 do while not rs444202.eof 

 
 
 rs444202.movenext
loop
 
 rs444202.close
  
 
 
else
%>
<div align="center"><font size="2" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif"><strong>0</strong></font></div>
<%



end if



%>
   <%'strSQL444202%>     </td>
	  
	  <td width="200"  bgcolor="<%=color1%>" height="18" style="border:1px solid #FFFFFF;"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%


 
 response.write rs("nome")
 

  



%>
          </font></div></td>
	  
	  
	  
	  
	  <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6" then %>
          <% if  UCase(rs("atendimento")) <> UCase(Session("nome_id")) then response.write "Não informado" else response.write rs("telefone") end if %>
          <%else%>
          <%response.write rs("telefone") end if %>
          </font> </div></td>
		 
	
	 
	 
	 
		  
		   
      
     
		  
       <td width="20" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"> 
        <div align="center"> <%if session("permissao") <> "4" and  session("permissao") <> "5" and  session("permissao") <> "6"  then %><%if  UCase(rs("atendimento")) <> UCase(Session("nome_id")) then %>&nbsp;<%else%><a href="javascript:newWindow2('form_enviar_email22.asp?varCodCompradores=<%=varCodCompradores%>')"><img src="bt_email22.jpg" width="25" height="18" border="0"></a><%end if%><%else%><a href="javascript:newWindow2('form_enviar_email22.asp?varCodCompradores=<%=varCodCompradores%>')"><img src="bt_email22.jpg" width="25" height="18" border="0"></a><%end if%></div></td>
	  
	 
      
      
    
	
	
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
      <div align="center"><font face="Verdana, arial" size="1" color="#000000"> 
        <%If cInt(intPage) > 1 Then%>
        <!-- se a página atual for maior que "1" então o link anteriro é colocado na 
			  na tela .-->
        <a href="?page=<%=intPage - 1%>&varCidade=<%=session("varCidade")%>&varCidade2=<%=session("varCidade2")%>&varBairro=<%=session("varBairro")%>&varBairro2=<%=session("varBairro2")%>&varNegociacao=<%=session("varNegociacao")%>&varTipo=<%=session("varTipo")%>&varQuartos=<%=session("varQuartos")%>&varVagas=<%=session("varVagas")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>&varAtendimento=<%=session("varAtendimento")%>&varOcupacao=<%=session("varOcupacao")%>&varCondominio=<%=session("varCondominio")%>&varAreaTotal=<%=session("varAreaTotal")%>&varSuites=<%=session("varSuites")%>&varPiscina=<%=session("varPiscina")%>&varPortaria=<%=session("varPortaria")%>&varQuintal=<%=session("varQuintal")%>&varQuadras=<%=session("varQuadras")%>&varEdicula=<%=session("varEdicula")%>&varStandby=<%=session("varStandby")%>&varCondicoes=<%=session("varCondicoes")%>" style="color:#000000"> 
        <font face="Verdana, arial" size="1" color="#000000"><b>Anterior</b></font></a> 
        <%End If%>
        </font></div></td>
          
    <td bgcolor="#FFFFFF"> <div align="center"><font face="Verdana, arial" color="#000000" size="1" > 
        <%If cInt(intPage) < cInt(intPageCount) and cInt(intPage) > 1 Then%>
        <!-- se página atual é menor que o total de páginas e intPage maior que um
			  ou seja, se não estiver na primeira página e nem na última então. -->
       
	   <%dim cont,cont2,i %>
	 
	 
	 <%if int(intPageCount) > 1 then%>
<%
If int(intPage)-5 > 1 then
cont=int(intPage)-5
else
cont=1
end if
%>
<%if cint(cont+10) > cint(intPageCount) then 
cont2=int(intPageCount)
else
cont2=int(cont)+10
end if
%>
<%for i=int(cont) to int(cont2)%>
<%

%>
<a href="?page=<%=i%>&varCidade=<%=session("varCidade")%>&varCidade2=<%=session("varCidade2")%>&varBairro=<%=session("varBairro")%>&varBairro2=<%=session("varBairro2")%>&varNegociacao=<%=session("varNegociacao")%>&varTipo=<%=session("varTipo")%>&varQuartos=<%=session("varQuartos")%>&varVagas=<%=session("varVagas")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>&varAtendimento=<%=session("varAtendimento")%>&varOcupacao=<%=session("varOcupacao")%>&varCondominio=<%=session("varCondominio")%>&varAreaTotal=<%=session("varAreaTotal")%>&varSuites=<%=session("varSuites")%>&varPiscina=<%=session("varPiscina")%>&varPortaria=<%=session("varPortaria")%>&varQuintal=<%=session("varQuintal")%>&varQuadras=<%=session("varQuadras")%>&varEdicula=<%=session("varEdicula")%>&varStandby=<%=session("varStandby")%>&varCondicoes=<%=session("varCondicoes")%>"><%if int(intPage) = int(i) then %><font color="#FF0000"><%else%><font color="#000000"><%end if%><%=i%></font>
</a> 
<%next%>
<%end if%>

	 
	   
	   
	   
	   
        <%End If%></font>
        </div></td>
          
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, arial" size="1" color="#000000" > 
        <%If cInt(intPage) < cInt(intPageCount)  Then%>
        <!-- se intPage é menor que o número de páginas então colocar o botão próximo -->
        <a href="?page=<%=intPage + 1%>&varCidade=<%=session("varCidade")%>&varCidade2=<%=session("varCidade2")%>&varBairro=<%=session("varBairro")%>&varBairro2=<%=session("varBairro2")%>&varNegociacao=<%=session("varNegociacao")%>&varTipo=<%=session("varTipo")%>&varQuartos=<%=session("varQuartos")%>&varVagas=<%=session("varVagas")%>&SearchFor=<%=session("SearchFor")%>&SearchWhere=<%=session("SearchWhere")%>&varValor=<%=session("varValor")%>&varValor1=<%=session("varValor1")%>&varValor2=<%=session("varValor2")%>&varAtendimento=<%=session("varAtendimento")%>&varOcupacao=<%=session("varOcupacao")%>&varCondominio=<%=session("varCondominio")%>&varAreaTotal=<%=session("varAreaTotal")%>&varSuites=<%=session("varSuites")%>&varPiscina=<%=session("varPiscina")%>&varPortaria=<%=session("varPortaria")%>&varQuintal=<%=session("varQuintal")%>&varQuadras=<%=session("varQuadras")%>&varEdicula=<%=session("varEdicula")%>&varStandby=<%=session("varStandby")%>&varCondicoes=<%=session("varCondicoes")%>" style="color:#000000"><font face="Verdana, arial" size="1" color="#000000"><b>Próximo</b></font> 
        </a> 
        <%End If%>
        </font></div></td>
        </tr>
      </table>










 
<%'else%>
<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
<div align="center"I></div>
</font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font> 
<%else%>
<%
End if
%>

  
<%
  rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
           %>
  <% response.flush%>
  <%response.clear%>
  <%else%>
  <table width="95" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      
      
    <td width="95" height="18">
      <% if session("permissao") = "2" or session("permissao") = "3" or session("permissao") = "4" or session("permissao") = "5" or session("permissao") = "6"  then %>
      <a href="javascript:newWindow2('verifica_tudo01.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a>
      <%else%>
      <img src="bt_incluir001.jpg" width="95" height="18" border="0">
      <%end if%>
    </td>
      
    </tr>
 </table>
 
 
 
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I><font color="<%=escuro%>">Comprador </font><font color="<%=escuro%>"> 
  n&atilde;o encontrado</font></div>
</font> 
<br>



 
  <%end if%>
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
group2[2][2]=new Option("até 200,00","0000000000 0000000200")
group2[2][3]=new Option(" até 500,00","0000000000 0000000500")
group2[2][4]=new Option(" até 750,00","0000000000 0000000750")
group2[2][5]=new Option(" até 1000,00","0000000000 0000001000")
group2[2][6]=new Option(" até 1500,00","0000000000 0000001500")
group2[2][7]=new Option(" até 2000,00","0000000000 0000002000")
group2[2][8]=new Option(" até 2500,00","0000000000 0000002500")
group2[2][9]=new Option(" até 3000,00","0000000000 0000003000")
group2[2][10]=new Option(" até 3500,00","0000000000 0000003500")
group2[2][11]=new Option(" até 4000,00","0000000000 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("Até  20.000,00","0000000000 0000020000")
group2[3][3]=new Option(" até 50.000,00","0000000000 0000050000")
group2[3][4]=new Option(" até 80.000,00","0000000000 0000080000")
group2[3][5]=new Option(" até 110.000,00","0000000000 0000110000")
group2[3][6]=new Option(" até 150.000,00","0000000000 0000150000")
group2[3][7]=new Option(" até 200.000,00","0000000000 0000200000")
group2[3][8]=new Option(" até 250.000,00","0000000000 0000250000")
group2[3][9]=new Option(" até 300.000,00","0000000000 0000300000")
group2[3][10]=new Option(" até 350.000,00","0000000000 0000350000")
group2[3][11]=new Option(" até 400.000,00","0000000000 0000400000")
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





'-----------------------------
           rs3.Close           
		   
           Set rs3 = Nothing
		   
'---------------------------------






'-----------------------------
           rs4.Close           
		   
           Set rs4 = Nothing
		   
'---------------------------------






'-----------------------------
           rs444atendimento.Close           
		   
           Set rs444atendimento = Nothing
		   
'---------------------------------


'-----------------------------
           rs444Tipo22.Close           
		   
           Set rs444Tipo22 = Nothing
		   
'---------------------------------








%>
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
SqlMarcas3 = "SELECT combo1.id_combo1,combo1.nome_combo1,combo1.data_combo1  FROM combo1 ORDER BY nome_combo1" 



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
'indica o tipo de cursor utilizado, se somente  de leitura ou se de leitura e gravação.

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
<%  EscreveFuncaoJavaScript ( Conexao3 ) %>



<%
conexao3.close
set conexao3 = nothing


set conexao = nothing


'response.write SQL

%> 

</body>
</html>

