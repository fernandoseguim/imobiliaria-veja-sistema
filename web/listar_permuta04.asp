<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->

<% response.buffer=True %>


<html>
<head>
<title></title>



</head>

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow2(abrejanela) {
   openWindow = window.open(abrejanela,'openWin','width=590,height=510,resizable=no,scrollbars=yes')
   openWindow.focus( )
   }

</SCRIPT>







<body bgcolor="#FFFFFF" vlink="#48576C" link="#48576C" alink="#000000" topmargin="3">
<p><font size="1" color="#CC6600" face="Verdana, Arial, Helvetica, sans-serif"><strong>Propriet&aacute;rios 
  que querem vender:</strong></font><br>
  
<td width="350" height="18" bgcolor="6497D0"><table width="350" border="0" cellspacing="0" cellpadding="0">
        <tr> 

          
      <td width="170" height="18" bgcolor="<%=escuro%>"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></div></td>
          
      <td width="170" height="18" bgcolor="<%=escuro%>"> 
        <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro</strong></font></div></td>
          
     
         
        </tr>
      </table></td>


  <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2
dim varNotFind
dim negrito,negrito2


  

 
 
'Codificar para receber "qualquer um" vcidade +2 pois vBairro est� conectado a vcidade
'----------------------------------------------------------
if session("vCidade_vend") = "" then
session("vCidade_vend") = request.querystring("vCidade")

end if

if session("vCidade_comp") = "" then
session("vCidade_comp") = request.querystring("vCidade4")

end if


if session("vBairro_vend") = "" then
session("vBairro_vend") = request.querystring("vBairro")
end if

if session("vBairro_comp") = "" then
session("vBairro_comp") = request.querystring("vBairro4")
end if




'------------------------recebendo tipos----------------


if session("vTipo_vend") = "" then
session("vTipo_vend") = request.querystring("vTipo_vend")
end if

if session("vTipo_comp") = "" then
session("vTipo_comp") = request.querystring("vTipo_comp")
end if





 '---------------------------------------------------  
 
 
 '--------------------------recebendo quartos------------------------
 
 
if session("vQuartos_vend") = "" then
session("vQuartos_vend") = request.querystring("vQuartos_vend")
end if

if session("vQuartos_comp") = "" then
session("vQuartos_comp") = request.querystring("vQuartos_comp")
end if

 '---------------------------------------------------------------------
 
 
 '-------------------------------recebendo valor------------------------
 
 
if session("vValor_vend") = "" then
session("vValor_vend") = request.querystring("vValor_vend")
end if

if session("vValor_comp") = "" then
session("vValor_comp") = request.querystring("vValor_comp")

end if



   
  
  session("vValor_vend1")=left(session("vValor_vend"),10)
   session("vValor_vend2")=right(session("vValor_vend"),10)
  
  
  
 
  
   session("vValor_comp1")=left(session("vValor_comp"),10)
   session("vValor_comp2")=right(session("vValor_comp"),10)



 
 '---------------------------------------------------------------------
 
 
 
 
 
 
 '------------------------Sua Cidade--------------------------


 
 
 
 stringCidadeVend = " where Cidade_vend='"&session("vCidade_vend")&"'"	
 
 
 
 
'--------------------------Seu bairro--------------------------------

dim stringBairroVend

 if   session("vBairro_vend") = "bqualquer" then
	stringBairroVend = ""
 else

stringBairroVend = " and Bairro_vend='"&session("vBairro_vend")&"'"

end if



 '--------------------------Tipo do seu im�vel------------------------
 
 
 dim stringTipoVend
 
 
 if session("vTipo_vend") = "tqualquer" then

stringTipoVend = ""

else
stringTipoVend = " and Tipo_vend='"&session("vTipo_vend")&"'"
 
 end if


 
 '-----------------------N�mero de quartos do seu im�vel-----------------
 
 
 
 
 dim stringQuartosVend
 
 
 if session("vQuartos_vend") = "qqualquer" then

stringQuartosVend = ""
else

stringQuartosVend = " and Quartos_vend='"&session("vQuartos_vend")&"'"
 end if
 


 
 
 
 
 '-----------------------------Valor de venda do seu im�vel----------------
 
 
 
	 dim stringValorVend
  if session("vValor_vend")="vqualquer" then
	stringValorVend = ""
	else
	
	stringValorVend = " and Valor_vend >="& session("vValor_vend1") &" and Valor_vend <="& session("vValor_vend2") &""
  end if
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '-------------------Cidade Pretendida-----------------------------------
 
 
 
	 dim stringCidadeComp
  if session("vCidade_comp")="cqualquer" then
	stringCidadeComp = ""
	else
	
	stringCidadeComp = " and Cidade_comp ='"& session("vCidade_comp") &"'"
	end if
	
 
 
 '----------------------------Bairro pretendido---------------------------------
 
 
	 dim stringBairroComp

	if session("vBairro_comp") = "bqualquer" then
	stringBairroComp = ""
	else
	
	stringBairroComp = " and Bairro_comp ='"& session("vBairro_comp") &"'"
	end if
	
	
	

	
	
	
 
 
 
 
 
 
 
 
 
 
 
 '-------------------------------------------------------------------------
 
 
 '------------------------------Tipo Pretendido---------------------------------
 
 
 
 
 
	 dim stringTipoComp
  if session("vTipo_comp")="tqualquer" then
	stringTipoComp = ""
	else
	
	
	stringTipoComp = " and Tipo_comp ='"& session("vTipo_comp") &"'"
	end if
	
	
 
 
 '-----------------------------------Quartos Pretendidos---------------------------------
 
 
 
 
	 dim stringQuartosComp
  if session("vQuartos_comp")="qqualquer" then
	stringQuartosComp = ""
	else
	
	stringQuartosComp = " and Quartos_comp ='"& session("vQuartos_comp") &"'"
	end if
	
 
 
 '-----------------------------------------------------------------------
 
 
 '----------------------------Valor pretendido----------------------------



	 dim stringValorComp
  if session("vValor_comp")="vqualquer" then
	stringValorComp = ""
	else
	
	
	stringValorComp = " and Valor_comp >="& session("vValor_comp1") &" and Valor_comp <="& session("vValor_comp2") &""
	end if
	
 
 
 
 
 
 
 
	
'---------------------------------------------------------------	
	
 
 strSQL = "SELECT * FROM permuta"&stringCidadeVend&stringBairroVend&stringTipoVend&stringQuartosVend&stringValorVend&stringCidadeComp&stringBairroComp&stringTipoComp&stringQuartosComp&stringValorComp
 
 
 
 
  'Aqui a vari�vel strSQL � defenida para depois ser usada no record set.
  
  
  
  
  
  
  
  
  '------------------------------------------------------
  
  
 
  
  
  
  
  
  
  
  
  
  
  
  '---------------------------------------------------------
  
  
  
   
 



Set RS = Server.CreateObject("ADODB.Recordset")
'um objeto recordset � inst�nciado.

Dim LinkTemp
'essa vari�vel vai ser usada como contador

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio
color2 = claro
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
	
RS.Open strSQL, Conn, 1, 3
'o recordset � aberto
	
RS.PageSize = 10
'Aqui configura-se o recordset para 20 registros por p�gina.

RS.CacheSize = RS.PageSize
'o Cache tamb�m conter� 20 registros por p�gina.

intPageCount = RS.PageCount
'A vari�vel intPageCount recebe o valor do n�mero de p�gina do recordset retornado.

intRecordCount = RS.RecordCount
'A vari�vel intRecordCount recebe o valor do n�mero de registros retornados no recordset.

If NOT (RS.BOF AND RS.EOF) Then
'verifica se existem registros retornados.
%>
<table width="537" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="537" height="36"><table width="537" border="0" cellspacing="0" cellpadding="0">
        
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

dim varCodPermuta



dim Conexao2,rs7
 Set Conexao2 = Server.CreateObject("ADODB.Connection")
	Set rs7 = Server.CreateObject("ADODB.RecordSet")
	Conexao2.Open dsn
	dim strSQL7
	
	
	



%>
<% varCodPermuta =RS("cod_permuta") %>
 <tr>
 <td>
 
 <table width="350" border="0" cellspacing="0" cellpadding="0">
  <tr>
               
  </tr>
  <tr>
    
  </tr>
  <tr>
    <td width="350" height="18" bgcolor="4780C5">
	<table width="350" border="0" cellspacing="0" cellpadding="0">
        <tr> 
                      <td width="170" height="18" bgcolor="<%=color1%>"> 
                        <div align="center"><a href="javascript:newWindow2('visualizar_permuta01.asp?varCodPermuta=<%=varCodPermuta%>')"><font face="Verdana, arial" size="1" color="white"><%=RS("nome")%></font></a></div></td>
                      <td width="170" height="18" bgcolor="<%=color1%>"> 
                        <div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Bairro_vend")%></font></div></td>
          
          
        </tr>
      </table>
	
	</td>
  
        
  
</table>
 
 
       
		
		 <%
RS.MoveNext


	  




If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If

if corfonte = "black" then
 corfonte = "white"
 
 else
 
 corfonte = "white"
 end if
 'acima � feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
		
		
		
		
      </table></td>
  </tr>
  
  <tr>
    <td width="537" height="18"><table width="537" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><div align="center"><font face="Verdana, arial" size="1"> 
              <%If cInt(intPage) > 1 Then%>
			  <!-- se a p�gina atual for maior que "1" ent�o o link anteriro � colocado na 
			  na tela .-->
              <a href="?page=<%=intPage - 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vTipo=<%=session("vTipo")%>">
              <b>Anterior</b></a> 
              <%End If%>
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" >
              
			  <!-- se p�gina atual � menor que o total de p�ginas e intPage maior que um
			  ou seja, se n�o estiver na primeira p�gina e nem na �ltima ent�o. -->
			  
             
              </font></div></td>
          <td><div align="center"><font face="Verdana, arial" size="1" > 
              <%If cInt(intPage) < cInt(intPageCount)  Then%>
			  <!-- se intPage � menor que o n�mero de p�ginas ent�o colocar o bot�o pr�ximo -->
              <a href="?page=<%=intPage + 1%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vTipo=<%=session("vTipo")%>"><b>Pr�ximo</b> 
              </a> 
              <%End If%>
              </font></div></td>
        </tr>
      </table></td>
  </tr>
</table>




<%End If


Else

%>
 <% 
    response.redirect "not_permuta01.html" 
  
  
  %>
  <% end if %>
  <%


RS.close
Set RS = Nothing



%>
  <% response.flush%>
  <%response.clear%>
  <!--#include file="dsn2.asp"-->


</body>
</html>
