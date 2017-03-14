<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin02.asp"-->

<!--#include file="cores.asp"-->
<html>
<head>
<title></title>

<!--#include file="style6_imoveis.asp"-->
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
   openWindow = window.open(abrejanela,'openWin','width=345,height=180,resizable=yes')
   openWindow.focus( )
   }

</SCRIPT>


</head>
<body  topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">
<table width="600" height="18" border="0" cellpadding="0" cellspacing="0">
  <tr> 
		  
      
    <td width="200" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_senha.asp">Senhas 
        do sistema</a></strong></font></div></td>
		   
    <td width="200" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ip.asp">IPs 
        do sistema</a></strong></font></div></td>
   
    <td width="200" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_telefone_acesso.asp">Telefone 
        de acesso</a></strong></font></div></td>
   
    </tr>
	

	
  </table>
<br>
<%
Dim orderBy
orderBy = request.querystring("orderby")
dim total
dim SQL
dim SearchFor
dim SearchWhere
SearchWhere = request("SearchWhere")
SearchFor = request.querystring("SearchFor")




dim varCodIP

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQL = "Select ip.id_ip,ip.ip,ip.origem_franquia,ip.quem_incluiu from ip  ORDER BY id_ip DESC"

rs.Open SQL, Conexao

%>
<br>
<center>
  <a href="archive_ip.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Carregar 
  Página</strong></font></a> 
</center>
<div align="left"><br>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Permissão 
  1: Apenas visualiza as informa&ccedil;&otilde;es<br>
  Permiss&atilde;o 2</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">: 
  Visualiza e atualiza as informa&ccedil;&otilde;es<br>
  Permiss&atilde;o 3: Visualiza , atualiza e exclui as informa&ccedil;&otilde;es<br>
  Permiss&atilde;o 4: Somente o administrador</font><br>
  <br>
  <%
If NOT (rs.BOF AND rs.EOF) Then
%>
</div>
<center>
<form  Method="Post" name="Formulario" action="multi_excluir_ip.asp?SearchFor=<%=SearchFor%>" >           
<table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
   
		
            <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);"></td>
            <td width="95" height="18"><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"></td>
             
            <td width="95" height="18"><a href="javascript:newWindow2('form_incluir_ip.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a></td>
           
		    <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Franquia</strong></font></div></td>
		    <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quem incluiu</strong></font></div></td>
		    <td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>IP</strong></font></div></td>
           
		</tr>
      
	   <%








Do While not rs.eof

'------------------------------------------------

%>
	 
	 
	      
				 
				<tr>
				<% varCodIP = rs("id_ip") %>
                <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("id_ip")%>"></td>
                 		 
				  <td width="95" height="18" bgcolor="<%=color1%>"><a href="excluir_ip.asp?varCodIP=<%=varCodIP%>&SearchFor=<%=SearchFor%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a></td>
                  <td width="95" height="18" bgcolor="<%=color1%>"><% if  session("permissao") = "6"  then %><a href="javascript:newWindow2('visualizar_ip.asp?varCodIP=<%=varCodIP%>')"><img src="bt_visualizar002.jpg" width="95" height="18" border="0"></img></a><%else%><img src="bt_visualizar002.jpg" width="95" height="18" border="0"></img><%end if%></img></td>
                  <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("origem_franquia")%></font></div></td>
				   <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("quem_incluiu")%></font></div></td>
				 
				  <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("ip")%></font></div></td>
               
			  </tr>
            
          <%
'-----------------------------------------------






rs.movenext


If colorchanger = 1 Then
	colorchanger = 0
	color1 = medio
	color2 = claro
Else
	colorchanger = 1
	color1 = claro
	color2 = medio
End If




loop
%>
          
          
       
  
</table>
</form>

</center>






 <%else%>
 
 
 
 
 
 
 <br>
 <center><a href="javascript:newWindow2('form_incluir_ip.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a></center>
 <br>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I>Ip não encontrado</div>
</font>            
 
           
            <%
End If

%>
        
<%
  rs.Close
           'fecha a conexão
           Conexao.Close
           Set rs = Nothing
		   set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>

</body>
</html>
