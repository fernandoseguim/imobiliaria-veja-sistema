<!--#include file="dsn.asp"-->

<%response.Buffer = true %>

<%

dim stringCidade2,stringBairro2,stringNegociacao2,stringQuartos2,stringValor2,stringTipo2
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


varIndicacaoCidade = request.querystring("varIndicacaoCidade")
varIndicacaoBairro = request.querystring("varIndicacaoBairro")
varIndicacaoNegociacao = request.querystring("varIndicacaoNegociacao")
varIndicacaoQuartos = request.querystring("varIndicacaoQuartos")
varIndicacaoVagas = request.querystring("varIndicacaoVagas")
varIndicacaoTipo = request.querystring("varIndicacaoTipo")
varIndicacaoValor = request.querystring("varIndicacaoValor")
varIndicacaoValor = int(varIndicacaoValor)
vValorMenor = int("0")
vValorMaior = int("0")

dim porcentual



%>

  <%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2

dim negrito,negrito2,varCodComprador



	  Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
 
 dim stringIndex2
stringIndex2 = " where cod_compradores <>"&"0"&""


if varIndicacaoCidade <> "qualquer um" and varIndicacaoCidade <> "n�o informado"  then
stringCidade2 = " and (cidade='"&varIndicacaoCidade&"' or cidade='"&"n�o informado"&"')"
else
stringCidade2 = ""
end if

 
 


 '--------------------------Bairro----------------------------

if varIndicacaoBairro <> "qualquer um" and varIndicacaoBairro <> "n�o informado" then
stringBairro2 = " and (Bairro like'%"&varIndicacaoBairro&"%' or Bairro like'%"&"n�o informado"&"%')"
else
stringBairro2 = ""
end if

 '------------------------------------------------------------- 

'--------------------------Tipo----------------------------

if varIndicacaoTipo <> "qualquer um" and varIndicacaoTipo <> "n�o informado" and varIndicacaoTipo <> "tqualquer" then
stringTipo2 = " and Tipo like'%"&varIndicacaoTipo&"%'"
else
stringTipo2 = ""
end if

 '------------------------------------------------------------- 







'-------------------Negocia��o---------------------------
if varIndicacaoNegociacao = "venda" then
vNegocio = "compra"
end if

if varIndicacaoNegociacao = "aluguel" then
vNegocio = "aluguel"
end if

if  varIndicacaoNegociacao <> "qualquer um" then
stringNegociacao2 = " and negociacao='"&vNegocio&"'"
else
stringNegociacao2 = ""
end if


'---------------------------Quartos------------------------------


if  varIndicacaoQuartos <> 0 then
stringQuartos2 = " and quartos <="&int(varIndicacaoQuartos)&""
else
stringQuartos2 = ""
end if

'---------------------------------------------------------------------------

'---------------------------Vagas------------------------------


if  varIndicacaoVagas <> 0 then
stringVagas2 = " and vagas <="&int(varIndicacaoVagas)&""
else
stringVagas2 = ""
end if

'---------------------------------------------------------------------------





'---------------------------------Valor-----------------------------------




 porcentual = int(varIndicacaoValor)*10/100

   vValorMenor = int(varIndicacaoValor)-int(porcentual)
   vValorMaior = int(varIndicacaoValor)+int(porcentual)
   






stringValor2 = " and Valor >="& vValorMenor &" and Valor <="& vValorMaior &""


dim stringStandby

stringStandby = " and standby = '"&"excluido"&"'"



 
 
 
 
 
 
 
 '------------------------------------------------------------------
 
 
 
 




strSQL ="SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores"&stringIndex2&stringCidade2&stringBairro2&stringTipo2&stringNegociacao2&stringQuartos2&stringVagas2&stringValor2&stringStandby
	


'----------------------------------------------------Fim da instru��o SQL---------------------------------
  
  
  dim EnderecoIP , vData
  vData = now()
  
 
 EnderecoIP = request.ServerVariables("REMOTE_ADDR")
 
  
  
  
  
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
color1 = "#537497"
color2 = "#94ADC8"
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

    <html>
	<head>
	<title>Indica��o de compradores</title>
	<STYLE>BODY {
SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-BASE-COLOR: #DF700F;}
</STYLE>


<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow3(abrejanela3) {
   openWindow3 = window.open(abrejanela3,'openWin3','width=600,height=510,resizable=no,scrollbars=yes')
   openWindow3.focus( )
   }

</SCRIPT>


	
	</head>
	<body bgcolor="EAA813">
	
<center>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atualmente 
  esse im&oacute;vel tem <%=intRecordCount%> indica��es.</strong><br>
  </font> 
</center>
<br>
<table width="708" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
  
  

  
  
    <td width="518" height="155">
	
	  
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
<% varCodComprador = rs("cod_compradores") %>
  
  
  
  
	
	
	
	
	
	
	  <table width="568" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="708" height="11"><img src="top_display2.jpg" width="708" height="11"></td>
  </tr>
  <tr> 
          <td width="708" height="153">
<table width="708" border="0" cellspacing="0" cellpadding="0">
              <tr> 
          <td width="8" height="153"><img src="left_display2.jpg" width="8" height="153"></td>
          <td><table width="692" height="153" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                <td width="552" height="16" bgcolor="FE9225"><table width="692" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="130"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade</strong></font></div></td>
                            <td width="140"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro/Regi�o</strong></font></div></td>
                           <td width="140"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Vila</strong></font></div></td>
						    <td width="99"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo</strong></font></div></td>
                            <td width="80"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Negocia&ccedil;&atilde;o</strong></font></div></td>
                            <td width="103"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Valor</strong></font></div></td>
                          </tr>
                        </table></td>
              </tr>
              <tr> 
                <td width="552" height="16" bgcolor="E17508">
				<table width="692" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="130"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Cidade")%></font></div></td>
                            <td width="140"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Bairro")%></font></div></td>
                            <td width="140"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Vila")%></font></div></td>
						    <td width="99"><div align="center"><font face="Verdana, arial" size="1" color="white"><%if rs("tipo") = "terreno" then response.Write"Terreno/�rea" else response.Write RS("Tipo") end if%></font></div></td>
                            <td width="80"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=RS("Negociacao")%></font></div></td>
                            <td width="103"><div align="center"><font face="Verdana, arial" size="1" color="white"><%=FormatNumber(RS("Valor"),2)%></font></div></td>
                          </tr>
                        </table>
				
				
				
				</td>
              </tr>
              <tr> 
                      <td bgcolor="FE9225"><div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:newWindow3('indicacao_display03.asp?varCodComprador=<%=varCodComprador%>')" style="color:#FFFFFF;text-decoration:none;" >
                          Ol� , meu nome � <strong><%=rs("nome")%></strong>, o 
                          sitema veja analizou os dados do seu im�vel e o que 
                          eu desejo comprar, e detectou a possibilidade de neg�cio 
                          entre n�s. Lique j� para <strong>4123-72-44</strong> 
                          e fale com o meu atendente o sr(a) <strong><%if rs("atendimento") = "Spirity" or rs("atendimento") = "internet" then response.write "Wanderlei" else response.write rs("atendimento") end if%></strong>, 
                          para que o mesmo agende uma visita minha ao seu im�vel, 
                          <strong>clique aqui</strong> e saiba mais sobre meus 
                          interesses e condi��es de pagamento. Muito Obrigado. 
                          </a></font></div></td>
              </tr>
            </table></td>
          <td width="8" height="153"><img src="right_display2.jpg" width="8" height="153"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
          <td width="708" height="11"><img src="bottom_display2.jpg" width="708" height="11"></td>
  </tr>
</table>
<br>
	
	
	 <%
RS.MoveNext


	  





 'acima � feito a troca de cores das tabelas e do texto dos recordsets.

If RS.EOF Then Exit for
Next	
%>
		
	
	
	
  </tr>
</table>
	
	</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><table width="518" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><div align="center"><font face="Verdana, arial" size="1"> 
            <%If cInt(intPage) > 1 Then%>
            <!-- se a p�gina atual for maior que "1" ent�o o link anteriro � colocado na 
			  na tela .-->
            <a href="?page=<%=intPage - 1%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoValor=<%=varIndicacaoValor%>"> 
            <b><font color="#000000">Anterior</font></b></a> 
            <%End If%>
            </font></div></td>
          <td width="250"><div align="center"><font face="Verdana, arial" size="1" > 
            </font></div></td>
          
        <td> 
          <%If cInt(intPage) < cInt(intPageCount)  Then%>
          <!-- se intPage � menor que o n�mero de p�ginas ent�o colocar o bot�o pr�ximo -->
          <a href="?page=<%=intPage + 1%>&varIndicacaoCidade=<%=varIndicacaoCidade%>&varIndicacaoBairro=<%=varIndicacaoBairro%>&varIndicacaoTipo=<%=varIndicacaoTipo%>&varIndicacaoNegociacao=<%=varIndicacaoNegociacao%>&varIndicacaoQuartos=<%=varIndicacaoQuartos%>&varIndicacaoValor=<%=varIndicacaoValor%>"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pr�ximo</strong></font></a></td><% end if%>
        </tr>
      </table></td>
  </tr>
</table>


<%End If


Else

%>

  <% 
response.write "<html><body bgcolor='EAA813'><br><br><br><center><font size='1' face='Verdana, Arial, Helvetica, sans-serif'><strong>Indica��o n�o encontrada!</strong></font></center></body></html>"

%>
  <% end if %>
  
  <script>
<!--

/*
Double Combo Script Credit
By JavaScript Kit (www.javascriptkit.com)
Over 200+ free JavaScripts here!
*/

var groups2=document.doublecombo.example2.options.length
/* Aqui � criada uma vari�vel "groups" que receber� os valores 
do combo example. */



var group2=new Array(groups2)
/* aqui a vari�vel group recebe os valores do "array(groups)" que contem os valores
do combo example.*/

for (i2=0; i2<groups2; i2++)
/* aqui temos um contador de zero at� o n�mero de elementos do array "groups" */

group2[i2]=new Array()
/* aqui � criado o array "group" que receber� valores conforme o n�mero de elementos
do array "groups". */

group2[0][0]=new Option("Qualquer Valor","vqualquer")


/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


group2[1][0]=new Option("Qualquer Valor","vqualquer")




/* aqui temos um array bidimensional "group" que receber� valores de op��es. */

group2[2][0]=new Option("Valor","vqualquer")
group2[2][1]=new Option("Qualquer Valor","vqualquer")
group2[2][2]=new Option("Menos de 200,00","0000000000 0000000200")
group2[2][3]=new Option("201,00 at� 500,00","0000000201 0000000500")
group2[2][4]=new Option("501,00 at� 750,00","0000000501 0000000750")
group2[2][5]=new Option("751,00 at� 1000,00","0000000751 0000001000")
group2[2][6]=new Option("1001,00 at� 1500,00","0000001001 0000001500")
group2[2][7]=new Option("1501,00 at� 2000,00","0000001501 0000002000")
group2[2][8]=new Option("2001,00 at� 2500,00","0000002001 0000002500")
group2[2][9]=new Option("2501,00 at� 3000,00","0000002501 0000003000")
group2[2][10]=new Option("3001,00 at� 3500,00","0000003001 0000003500")
group2[2][11]=new Option("3501,00 at� 4000,00","0000003501 0000004000")
group2[2][12]=new Option("Mais de 4000,00","0000004001 1000000000")







group2[3][0]=new Option("Valor","vqualquer")
group2[3][1]=new Option("Qualquer Valor","vqualquer")
group2[3][2]=new Option("At�  20.000,00","0000000000 0000020000")
group2[3][3]=new Option("20.001,00 at� 50.000,00","0000020001 0000050000")
group2[3][4]=new Option("50.001,00 at� 80.000,00","0000050001 0000080000")
group2[3][5]=new Option("80.001,00 at� 110.000,00","0000080001 0000110000")
group2[3][6]=new Option("110.001,00 at� 150.000,00","0000110001 0000150000")
group2[3][7]=new Option("150.001,00 at� 200.000,00","0000150001 0000200000")
group2[3][8]=new Option("200.001,00 at� 250.000,00","0000200001 0000250000")
group2[3][9]=new Option("250.001,00 at� 300.000,00","0000250001 0000300000")
group2[3][10]=new Option("300.001,00 at� 350.000,00","0000300001 0000350000")
group2[3][11]=new Option("350.001,00 at� 400.000,00","0000350001 0000400000")
group2[3][12]=new Option("Acima de 400.000,00","0000400001 1000000000")









/* aqui temos um array bidimensional "group" que receber� valores de op��es. */


var temp2=document.doublecombo.stage22
/* aqui a vari�vel "temp" recebe os valores do segundo combo o "stage2" */

function redirect2(x2){
/* aqui � criada a fun��o "redirect" que comanda o carregamento do combo "stage2" */

for (m2=temp2.options.length-1;m2>0;m2--)
temp2.options[m2]=null
/* aqui temos um contador "m" que d� um valor nulo para o combo "stage2" para que 
posteriormente esse combo possa receber os valores determinados. */

for (i2=0;i2<group2[x2].length;i2++){
temp2.options[i2]=new Option(group2[x2][i2].text,group2[x2][i2].value)
/* aqui o combo "stage2" recebe os valores do array "group" dependendo do que � escolhido no
primeiro combo "example".*/

}
temp2.options[0].selected=true
}
/* aqui o array "temp.options[0]" ser� o valor inicial selecionado ele corresponde ao array
"stage2".*/

function go(){
location2=temp2.options[temp2.selectedIndex].value
}

/* aqui  a vari�vel "location" recebe os valores de "stage2" que corresponde ao endere�o de
link para o carregamento de p�gina. */


//-->
</script>
  <%


RS.close
Set RS = Nothing



%>
  <% response.flush%>
  <%response.clear%>
  
  
  <%
  
  conexao.close
  
  set conexao = nothing
  
  %>
  <!--#include file="dsn2.asp"-->
 <!--#include file="cores.asp"-->
<br>


</body>
</html>
