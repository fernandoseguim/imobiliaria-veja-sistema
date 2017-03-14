<%
Function EscreveFuncaoJavaScript ( Conexao33 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo1.options[form.combo1.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas33 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas33 = Conexao33.Execute ( SqlMarcas33 )

While NOT rsMarcas33.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas33("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo2.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros33 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas33("id_combo1")&" order by nome_combo2"
Set rsCarros33 = Conexao33.Execute ( SqlCarros33 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
While NOT rsCarros33.EoF

Response.Write "form.combo2.options[" & i & "] = new Option('" & rsCarros33("nome_combo2") & "','" & rsCarros33("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros33.MoveNext
Wend
Response.Write "form.combo2.options[" & i & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" & vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas33.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao33 = Server.CreateObject("ADODB.Connection")
Conexao33.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql33 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Sql55 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs55 = Conexao33.Execute ( Sql55 )
Set Rs33 = Conexao33.Execute ( Sql33 ) 
%> 










<%
Function EscreveFuncaoJavaScript22 ( Conexao44 )
'O parametro conexao receberá uma conexao aberta!
'Em funcoes, geralmente não criamos objetos do tipo conexões!
'Opte por sempre deixar sua função o mais compatível possível com qualquer projeto!

'Primeiro vamos escrever o cabecalho de qualquer script javascript!
Response.Write "<script language=""JavaScript"">" & vbcrlf 
Response.Write "function atualizacarros2 (form) {" & vbcrlf

'Essa função JavaScript recebe o form em que estão os campos a serem atualizados!
'Veja na chamada da função no método OnChange em que se passa o this.form!

'Vamos criar um switch para ele verificar qual opção foi selecionada!! 
Response.Write "switch (form.combo3.options[form.combo3.selectedIndex].value) {" & vbcrlf 

'Agora entramos com o banco de dados! Temos que jogar aqui todas as opções de carro!
SqlMarcas44 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set rsMarcas44 = Conexao44.Execute ( SqlMarcas44 )

While NOT rsMarcas44.EOF

'Caso tenha sido essa marca selecionada... 
Response.Write "case '" & rsMarcas44("id_combo1") & "':" & vbcrlf

'Apagamos tudo o que tem na caixa dos carros!
Response.Write "form.combo4.length=0;" & vbcrlf 

'Abrimos todos os carros relativos a essa marca!
SqlCarros44 = "SELECT * FROM combo2 WHERE id_combo1 =" & rsMarcas44("id_combo1")&" order by nome_combo2"
Set rsCarros44 = Conexao44.Execute ( SqlCarros44 )

'Fazemos um loop por todos os carros, criando uma nova opção no SELECT! 
i = 0 
While NOT rsCarros44.EoF

Response.Write "form.combo4.options[" & i & "] = new Option('" & rsCarros44("nome_combo2") & "','" & rsCarros44("id_combo2") & "');" & vbcrlf 
i=i+1

rsCarros44.MoveNext
Wend
Response.Write "form.combo4.options[" & i & "] = new Option('" & "qualquer um" & "','" & "bqualquer" & "');" & vbcrlf
'Imprimos um break! (Verifique tutoriais de JavaScript, se tiverem alguma dúvida da sua utilização! 
Response.Write "break;" & vbcrlf

'Próxima marca! 
rsMarcas44.MoveNext 
Wend 

'Fecha chaves do switch e da função! E fecha o script! 
Response.Write "}}" & vbcrlf & "</script>" & vbcrlf 

End Function
%> 




















<%
'Criando conexão com o banco de dados! 
Set Conexao44 = Server.CreateObject("ADODB.Connection")
Conexao44.Open "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & Server.MapPath("bd_araquio.mdb")

'Abrindo a tabela MARCAS!
Sql44 = "SELECT * FROM combo1 ORDER BY nome_combo1" 
Set Rs55 = Conexao44.Execute ( Sql44 ) 
%> 













<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<% response.buffer=True%>

<%
Dim Conexaoo,strSQLL,rss,intRecordCountt,varCod_imovell,varSucesso_imovell
varCod_imovell = request.QueryString("varCod_imovell")
varSucesso_imovell = request.QueryString("varSucesso_imovell")
   
   Set rss = Server.CreateObject("ADODB.RecordSet")
   dim rs44,strSQL44,strSQL66,rs66
   
    Set rs44 = Server.CreateObject("ADODB.RecordSet")
	Set rs66 = Server.CreateObject("ADODB.RecordSet")
	strSQL44 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	strSQL66 = "SELECT * FROM combo2 where id_combo1 = 3  ORDER BY nome_combo2" 
	
    Set Conexaoo = Server.CreateObject("ADODB.Connection")
	strSQLL = "SELECT * FROM imoveis where Foto_Grande not like 'imovel00000.jpg' order by cod_imovel DESC "
	 
   Conexaoo.Open dsn
   
RSS.CursorLocation = 3
RSS.CursorType = 3

        rss.Open strSQLL, Conexaoo 
		rs44.Open strSQL44, Conexaoo
		rs66.Open strSQL66, Conexaoo
		
		
%>		


<%





dim strSQL,rs,varCodimovel,coratual,corfonte,rsMax,regAtual,NRecords,pagina,paginas, NumReg, resto, vTipo, vBairro, vNegociacao, vValor, page,SQL, vCidade,vValor1,vValor2
dim varNotFind
dim negrito,negrito2
dim vTipo_comp,vTipo_vend
dim vQuartos_vend,vQuartos_comp
dim vValor_vend,vValor_comp

vQuartos_vend=request.querystring("txt_Quartos_vend")
  session("vQuartos_vend") = vQuartos_vend

vQuartos_comp=request.querystring("txt_Quartos_comp")
  session("vQuartos_comp") = vQuartos_comp
  
  
  
'--------------recebe valores-----------------------------


vValor_vend=request.querystring("txt_Valor_vend")
  session("vValor_vend") = vValor_vend

vValor_comp=request.querystring("txt_Valor_comp")
  session("vValor_comp") = vValor_comp





 
  vTipo_vend=request.querystring("txt_tipo_vend")
  session("vTipo_vend") = vTipo_vend
  
  
   dim vCidade2
   
    vCidade2=request.querystring("combo1")
   session("vCidade2") = vCidade2
 
   
   
   
	
	
	
	
	 if session("vCidade2") = "" then
session("vCidade2") = request.querystring("vCidade2")

end if
	  Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	if session("vCidade2") <> "cqualquer" then
	
	dim rs2,SQL2
 Set rs2 = Server.CreateObject("ADODB.RecordSet")
 SQL2 = "select * from combo1 where id_combo1 ="&session("vCidade2")
 
 rs2.open SQL2,Conexao,2,1
 
 vCidade = rs2("nome_combo1")
 else
 vCidade = vCidade2
 end if

	session("vCidade_vend")= vCidade
	
	
	
	dim vBairro2
	 vBairro2=request.querystring("combo2")
	 session("vBairro2") = vBairro2
	 if session("vBairro2") = "" then
session("vBairro2") = request.querystring("vBairro2")

end if
	 
	 if session("vBairro2") <> "bqualquer" then
	  dim rs3,SQL3
 Set rs3 = Server.CreateObject("ADODB.RecordSet")
 SQL3 = "select * from combo2 where id_combo2 ="& session("vBairro2")
 
 rs3.open SQL3,Conexao,2,1

 vBairro = rs3("nome_combo2")
 else
 vBairro = vBairro2
	end if                                      
									
	 
	 
	 
	 session("vBairro_vend")= vBairro
	  
	  
	 if session("vCidade")="sao bernardo" then
	  session("vCidade")="São Bernardo"
	  end if
	 
	   if session("vCidade")="santo andre" then
	  session("vCidade")="Santo André"
	  end if
	  
	   if session("vCidade")="sao caetano" then
	  session("vCidade")="São Caetano"
	  end if
	  
	  
	  if session("vBairro")="bairro assuncao" then
	  session("vBairro")="Bairro Assunção"
	  end if
	  
	   if session("vBairro")="ceramica" then
	  session("vBairro")="Cerâmica"
	  end if
	  
	  
	  if session("vBairro")="jd sao caetano" then
	  session("vBairro")="JD São Caetano"
	  end if 
	  'Acima as variáveis recebem os valores dos formulários para fazer busca.
	
	'-----------------------------------------------------------------------------------
	
	
	
  vTipo_comp=request.querystring("txt_tipo_comp")
  session("vTipo_comp") = vTipo_comp
  
  if session("vTipo_comp") = "" then
  session("vTipo_comp") = request.querystring("vTipo_comp")
  end if
  
   dim vCidade3
   
   vCidade3=request.querystring("combo3")
   
   session("vCidade3") = vCidade3
   
   if session("vCidade3") = "" then
session("vCidade3") = request.querystring("vCidade3")
end if
   
   
    
	
	
	
	
	  Set Conexao2 = Server.CreateObject("ADODB.Connection")
	
	Conexao2.Open dsn
	
	if session("vCidade3") <> "cqualquer" then
	
	dim rs4,SQL4
 Set rs4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "select * from combo1 where id_combo1 ="&session("vCidade3")
 
 rs4.open SQL4,Conexao2,2,1
 
 dim vCidade4
 
 vCidade4 = rs4("nome_combo1")
 else
 vCidade4 = vCidade3
 end if

	session("vCidade_comp")= vCidade4
	
	
	
	dim vBairro3
	 vBairro3=request.querystring("combo4")
	 session("vBairro3") = vBairro3
	 if session("vBairro3") = "" then
session("vBairro3") = request.querystring("vBairro3")

end if
	 
	 if session("vBairro3") <> "bqualquer" then
	  dim rs5,SQL5
 Set rs5 = Server.CreateObject("ADODB.RecordSet")
 SQL5 = "select * from combo2 where id_combo2 ="& session("vBairro3")
 
 rs5.open SQL5,Conexao2,2,1

 vBairro4 = rs5("nome_combo2")
 else
 vBairro4 = vBairro3
	end if                                      
									
	 
	 
	 
	 session("vBairro_comp")= vBairro4
	  
	  
	 if session("vCidade4")="sao bernardo" then
	  session("vCidade4")="São Bernardo"
	  end if
	 
	   if session("vCidade4")="santo andre" then
	  session("vCidade4")="Santo André"
	  end if
	  
	   if session("vCidade4")="sao caetano" then
	  session("vCidade4")="São Caetano"
	  end if
	  
	  
	  if session("vBairro4")="bairro assuncao" then
	  session("vBairro4")="Bairro Assunção"
	  end if
	  
	   if session("vBairro4")="ceramica" then
	  session("vBairro4")="Cerâmica"
	  end if
	  
	  
	  if session("vBairro4")="jd sao caetano" then
	  session("vBairro4")="JD São Caetano"
	  end if 
	  'Acima as variáveis recebem os valores dos formulários para fazer busca.
	
	
	
	
	%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%  EscreveFuncaoJavaScript ( Conexao33 ) %>
<%  EscreveFuncaoJavaScript22 ( Conexao33 ) %>

<script>
function isValidDigitNumber (doublecombo)
{





























	
	if (doublecombo.combo1.value == "cqualquer") {
        alert("Escolha uma cidade!");
        doublecombo.combo1.focus();
		
        return false;
    }
	
	
	
	
	if (doublecombo.combo3.value == "cqualquer") {
        alert("Escolha uma cidade pretendida!");
        doublecombo.combo3.focus();
		
        return false;
    }
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}

//-----------------------------------------------










</script>



</head>
<!--#include file="style_imoveis.asp"-->
<body>
<table width="790" border="1" cellspacing="0" cellpadding="0" bordercolor="#FFFFFF">
  <tr bgcolor="<%=claro%>"> 
   
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_imoveis_inicial.asp" style="color: FFFFFF">Im&oacute;veis</a></strong></font></div></td>
    
	  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_proposta_inicial.asp" style="color: FFFFFF">Proposta</a></strong></font></div></td>
     
    
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>" > 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_email_inicial.asp" style="color: FFFFFF">Email</a></strong></font></div></td>
   
   
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_cidade.asp" style="color: FFFFFF">Cidades</a></strong></font></div></td>
  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_bairro_inicial.asp" style="color: FFFFFF">Bairros</a></strong></font></div></td>
  
  <td height="18" bordercolor="#FFFFFF" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_compradores_inicial.asp" style="color: FFFFFF">Compradores</a></strong></font></div></td>
 
  
  
    <td height="18" bordercolor="#FFFFFF" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_permuta_inicial.asp" style="color: FFFFFF">Permuta</a></strong></font></div></td>
 
  
  </tr>
</table>


<br>
<form name="doublecombo"  onSubmit="return isValidDigitNumber(this);" method="get" action="listar_permuta02.asp">

<center><font color="#CC6600" size="5" face="Times New Roman, Times, serif"><strong>Pesquisa por Permuta</strong></font> </center>
<br>




<%
	if varSucesso_imovel = "" then
	response.Write varSucesso_imovel
	%>
          <%else%>
          <font color="000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso_imovel%> 
          foi incluido com sucesso.</font> 
          <% end if %>



  <table width="590" border="0" cellspacing="0" cellpadding="0">
    
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
          atual</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> <a href="javascript:newWindow3('form_incluir_cidade.asp')"></img></a><font color="#FFFFFF"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <select name="combo1" class="inputBox" id="select2" style="HEIGHT: 18px; WIDTH: 150px; background:white;color:black;" onChange="javascript:atualizacarros(this.form);">
                     <option value="cqualquer" selected>Cidade</option>
					  <% if not rs33.eof then %>
                      <% While NOT Rs33.EoF %>
                      <option value="<% = Rs33("id_combo1") %>" <% if rs33("nome_combo1") = "Santo André" then%><%else%><%end if%>> 
                      <% = Rs33("nome_combo1") %>
                      </option>
                      <% Rs33.MoveNext %>
                      <% Wend %>
                      <%else%>
                      <option value=""></option>
                      <%end if%>
                    </select>
                    </font></font></td>
              </tr>
                <tr> 
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      atual</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;">
<select name="combo2" class="inputBox" id="select4" style="HEIGHT: 18px; WIDTH: 150px; background:white;color:black;">
                      <option value="bqualquer" selected>Bairro</option>
					  <% if not rs44.eof then%>
                      <% While NOT Rs44.EoF %>
                      <option value="<% = Rs44("id_combo2") %>"<%if rs44("nome_combo2") = "Bairro Campestre" then%><%end if%>> 
                      <% = Rs44("nome_combo2") %>
                      </option>
                      <% Rs44.MoveNext %>
                      <% Wend %>
                      <% else %>
                      <option value=""></option>
                      <% end if %>
                    </select> </td>
              </tr>
			  
			   <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      de im&oacute;vel atual</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
          <select name="txt_tipo_vend" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: white;color:black;">
            <option value="tqualquer" selected>Tipo</option>
					<option value="casa" >Casa</option>
                    <option value="apartamento">Apartamento </option>
                    <option value="flat">Flat</option>
                    <option value="terreno">Terreno</option>
                    <option value="rural">Rural</option>
                    <option value="comercial">Comercial</option>
                  </select>
                    </font></td>
              </tr>
			  
			  
			  
			  
              <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
            de quartos</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
          <select name="txt_quartos_vend" size="1" id="select3" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: white;color:black;">
            <option value="qqualquer" selected>Quartos</option>
					<option value="01" >01</option>
                    <option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                    </font></td>
              </tr>
               
			   
			  
			  <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
            do im&oacute;vel</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
          <select name="txt_valor_vend" size="1" id="txt_valor_vend" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: white;color:black;">
            		<option value="vqualquer">Valor</option>
                      <option value="vqualquer">Qualquer um</option>
                      <option value="0000000000 0000020000">menos de 20.000,00</option>
                      <option value="0000020000 0000050000">20.000,00 até 50.000,00</option>
                      <option value="0000050000 0000100000">50.000,00 até 100.000,00</option>
                      <option value="0000100000 0000200000">100.000,00 até 200.000,00</option>
                      <option value="0000200000 1000000000">acima de 200.000,00</option>
                  </select>
                    </font></td>
              </tr> 
			   
			   
			   
			   
			   
			   
              <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cidade 
                      pretendida </font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> <font color="#FFFFFF"> 
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <select name="combo3" class="inputBox" id="combo3" style="HEIGHT: 18px; WIDTH: 150px; background:white;color:black;" onChange="javascript:atualizacarros2(this.form);">
                      <option value="cqualquer" selected>Cidade</option>
					  <% if not rs55.eof then %>
                      <% While NOT Rs55.EoF %>
                      <option value="<% = Rs55("id_combo1") %>" <% if rs55("nome_combo1") = "Santo André" then%><%else%><%end if%>>
                    <% = Rs55("nome_combo1") %>
                    </option>
                    <% Rs55.MoveNext %>
                    <% Wend %>
					<%else%>
					<option value=""></option>
					<%end if%>
                  </select>
                    </font></font> </td>
              </tr>
                <tr> 
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Bairro 
                      pretendido </font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;">
<select name="combo4" class="inputBox" id="combo4" style="HEIGHT: 18px; WIDTH: 150px; background:white;color:black;">
                     <option value="bqualquer" selected>Bairro</option>
					  <% if not rs66.eof then%>
                      <% While NOT Rs66.EoF %>
                      <option value="<% = Rs66("id_combo2") %>"<%if rs66("nome_combo2") = "Bairro Campestre" then%><%end if%>>
                    <% = Rs66("nome_combo2") %>
                    </option>
                    <% Rs66.MoveNext %>
                    <% Wend %>
					<% else %>
					<option value=""></option>
					<% end if %>
					
					
                  </select></td>
              </tr>
              <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo 
                      de im&oacute;vel pretendido</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
          <select name="txt_tipo_comp" size="1" id="select" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: white;color:black;">
            <option value="tqualquer" selected>Tipo</option>
					  <option value="casa" >Casa</option>
                      <option value="apartamento">Apartamento </option>
                      <option value="flat">Flat</option>
                      <option value="terreno">Terreno</option>
                      <option value="rural">Rural</option>
                      <option value="comercial">Comercial</option>
                    </select>
                    </font> </td>
              </tr>
              <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
            de quartos</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
          <select name="txt_quartos_comp" size="1" id="txt_quartos_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: white;color:black;">
            <option value="qqualquer" selected>Quartos</option>
					<option value="01" >01</option>
                    <option value="02">02 </option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
					<option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                  </select>
                    </font></td>
              </tr>
			  
			  <tr>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"> 
        <div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
            do im&oacute;vel</font></div></td>
                  
      <td bgcolor="#FFFFFF" style="border:1px solid #FFFFFF;"><font color="#FFFFFF"> 
          <select name="txt_valor_comp" size="1" id="txt_valor_comp" class="inputBox" style="HEIGHT: 18px; WIDTH: 150px; background: white;color:black;">
            		<option value="vqualquer">Valor</option>
                      <option value="vqualquer">Qualquer um</option>
                      <option value="0000000000 0000020000">menos de 20.000,00</option>
                      <option value="0000020000 0000050000">20.000,00 até 50.000,00</option>
                      <option value="0000050000 0000100000">50.000,00 até 100.000,00</option>
                      <option value="0000100000 0000200000">100.000,00 até 200.000,00</option>
                      <option value="0000200000 1000000000">acima de 200.000,00</option>
                  </select>
                    </font></td>
              </tr> 
			  
			  
			  
			  
              <tr>
                <td>&nbsp;</td>
                  <td><table width="290" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><input name="image" type="image"  src="bt_procurar001.jpg" width="149" height="18" border="0"></td>
                        
            <td>&nbsp;</td>
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


<table width="700" height="400" border="0" cellpadding="0" cellspacing="0">
  <tr>   
  <td><iframe src="listar_permuta03.asp?&vCidade3=<%=session("vCidade3")%>&vCidade4=<%=session("vCidade4")%>&vBairro3=<%=session("vBairro3")%>&vBairro4=<%=session("vBairro4")%>&vTipo_comp=<%=session("vTipo_comp")%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vTipo_vend=<%=session("vTipo_vend")%>&vQuartos_vend=<%=session("vQuartos_vend")%>&vQuartos_comp=<%=session("vQuartos_comp")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_comp=<%=session("vValor_comp")%>" name="meio" width="350px" height="400px" frameborder="0" scrolling="no"></iframe></td>
  <td><iframe src="listar_permuta04.asp?&vCidade3=<%=session("vCidade3")%>&vCidade4=<%=session("vCidade4")%>&vBairro3=<%=session("vBairro3")%>&vBairro4=<%=session("vBairro4")%>&vTipo_comp=<%=session("vTipo_comp")%>&vCidade=<%=session("vCidade")%>&vCidade2=<%=session("vCidade2")%>&vBairro=<%=session("vBairro")%>&vBairro2=<%=session("vBairro2")%>&vTipo_vend=<%=session("vTipo_vend")%>&vQuartos_vend=<%=session("vQuartos_vend")%>&vQuartos_comp=<%=session("vQuartos_comp")%>&vValor_vend=<%=session("vValor_vend")%>&vValor_comp=<%=session("vValor_comp")%>" name="meio" width="350px" height="400px" frameborder="0" scrolling="no"></iframe></td>
  </tr>
</table>
</body>
</html>
