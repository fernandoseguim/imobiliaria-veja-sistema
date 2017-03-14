<% response.Buffer = true %>
<!--#include file="dsn.asp"-->


<!--#include file="cores.asp"-->
<html>
<head>
<title></title>

<!--#include file="style6_imoveis.asp"-->


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
		   
    <td width="200" bgcolor="<%=claro%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_ip.asp">IPs 
        do sistema</a></strong></font></div></td>
   
    <td width="200" bgcolor="<%=medio%>"> 
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




dim varCod_Telefone_acesso

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQL = "Select telefone_acesso.cod_telefone_acesso,telefone_acesso.telefone_acesso,telefone_acesso.origem_franquia,telefone_acesso.quem_incluiu from telefone_acesso  ORDER BY cod_telefone_acesso DESC"

rs.Open SQL, Conexao

%>
<br>
<center>
  <a href="archive_telefone_acesso.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Carregar 
  Página</strong></font></a> 
</center>
<div align="left"><br>
  <br>
  <br>
  <%
If NOT (rs.BOF AND rs.EOF) and ucase(session("Admin_id")) = ucase("WSBRAGA") Then
%>
</div>
<center>
<form  Method="Post" name="Formulario" action="multi_excluir_telefone_acesso.asp?SearchFor=<%=SearchFor%>" >           
<table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
   
		
            <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);"></td>
            <td width="95" height="18"><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"></td>
             
            <td width="95" height="18"><a href="javascript:newWindow2('form_incluir_telefone_acesso.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a></td>
            <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Franquia</strong></font></div></td>
            <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Quem incluiu</strong></font></div></td>
           
		    <td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefone de acesso</strong></font></div></td>
           
		</tr>
      
	   <%








Do While not rs.eof

'------------------------------------------------

%>
	 
	 
	     
			    <tr>
				 <% varCod_telefone_acesso = rs("cod_telefone_acesso") %>
                <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("cod_telefone_acesso")%>"></td>
                  <td width="95" height="18" bgcolor="<%=color1%>"><a href="excluir_telefone_acesso.asp?varCod_telefone_acesso=<%=varCod_telefone_acesso%>&SearchFor=<%=SearchFor%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a></td>
                  <td width="95" height="18" bgcolor="<%=color1%>"><a href="javascript:newWindow2('visualizar_telefone_acesso.asp?varCod_telefone_acesso=<%=varCod_telefone_acesso%>')"><img src="bt_visualizar002.jpg" width="95" height="18" border="0"></img></a></td>
                 <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("origem_franquia")%></font></div></td>
				 <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("quem_incluiu")%></font></div></td>
				  <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("telefone_acesso")%></font></div></td>
               
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
 
 
 
 
 <% if ucase(session("Admin_id")) = ucase("WSBRAGA") then %>
 
 <br>
 <center><a href="javascript:newWindow2('form_incluir_telefone_acesso.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a></center>
 <br>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I>Telefone de acesso não encontrado</div>
</font>  
<%else%>

<%end if%>          
 
           
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
