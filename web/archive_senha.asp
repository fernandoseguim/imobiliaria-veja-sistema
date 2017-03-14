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
   openWindow = window.open(abrejanela,'openWin','width=345,height=250,resizable=yes')
   openWindow.focus( )
   }

</SCRIPT>






</head>
<body  topmargin="0" bgcolor="FFFFFF" vlink="#FFFFFF" link="#FFFFFF" alink="#FFFFFF">


<table width="600" height="18" border="0" cellpadding="0" cellspacing="0">
  <tr> 
		  
      
    <td width="200" bgcolor="<%=medio%>"> 
      <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="archive_senha.asp">Senhas 
        do sistema</a></strong></font></div></td>
		   
    <td width="200" bgcolor="<%=claro%>"> 
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
dim Admin_ID
dim Password

Admin_ID = request.form("Admin_ID")
Password = request.form("Password")


session("Password") = Password

SearchWhere = request("SearchWhere")
SearchFor = request.querystring("SearchFor")




dim varCodSenha

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   
 '-------------------------pegar nome de usuário--------------  
  
   dim rs_usuario,SQL_usuario
   
Set rs_usuario = Server.CreateObject("ADODB.RecordSet")

 
 SQL_usuario = "Select senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id,senha.origem_franquia  from senha where admin_id like '"&session("Admin_ID")&"' and admin_pass like '"&session("Password")&"'  ORDER BY ID DESC"      
	
rs_usuario.Open SQL_usuario, Conexao	
	
	dim vOrigem_usuario
	
	if not  rs_usuario.eof then
	vOrigem_usuario = rs_usuario("origem_franquia")
	else
	vOrigem_usuario = "Sao Bernardo"
	end if
		


SQL = "Select senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id,senha.origem_franquia  from senha where origem_franquia like '"&vOrigem_usuario&"' ORDER BY ID DESC"

rs.Open SQL, Conexao


%>
<br>
<center>
  <a href="archive_senha.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Carregar 
  Página</strong></font></a> 
</center>
<div align="left"><br>
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">Permissão 
  1: Apenas visualiza as informa&ccedil;&otilde;es<br>
  Permiss&atilde;o 2</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">: 
  Visualiza e atualiza as informa&ccedil;&otilde;es<br>
  Permiss&atilde;o 3: Visualiza , atualiza e verifica estat&iacute;stica<br>
  Permiss&atilde;o 4:Visualiza, atualiza, verifica estat&iacute;sticas e excluir 
  informa&ccedil;&otilde;es <br>
  Permiss&atilde;o 5: Somente o administrador</font><br>
  <br>
  <%
If NOT (rs.BOF AND rs.EOF) Then
%>
</div>
<form  Method="Post" name="Formulario" action="multi_excluir_senha.asp?SearchFor=<%=SearchFor%>" >           
<table width="880" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="880" height="18"><table width="880" height="18" border="0" cellpadding="0" cellspacing="0">
        <tr>
		
            <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);"></td>
            <td width="95" height="18"><% if  session("permissao") = "6"  then %><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"><%else%><img src="bt_excluir002.jpg" width="95" height="18" border="0"></img><%end if%></td>
             
            <td width="95" height="18"><% if  session("permissao") = "6"  then %><a href="javascript:newWindow2('form_incluir_senha.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a><%else%><img src="bt_incluir001.jpg" width="95" height="18" border="0"><%end if%></td>
           <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Franquia </strong></font></div></td>
		    <td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome </strong></font></div></td>
            <td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>ID</strong></font></div></td>
             <td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Senha</strong></font></div></td>
		      <td width="130" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Permissão</strong></font></div></td>
		
  </tr>
  
    
     
	   <%








Do While not rs.eof

'------------------------------------------------

%>
	 
	 
	   
			   <% varCodSenha = rs("ID") %>
			  <tr>
                <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("ID")%>"></td>
                  <td width="95" height="18" bgcolor="<%=color1%>"><% if  session("permissao") = "6"  then %><a href="excluir_senha.asp?varCodSenha=<%=varCodSenha%>&SearchFor=<%=SearchFor%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a><%else%><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img><%end if%></td>
                  <td width="95" height="18" bgcolor="<%=color1%>"><% if  session("permissao") = "6"  then %><a href="javascript:newWindow2('visualizar_senha.asp?varCodSenha=<%=varCodSenha%>')"><img src="bt_visualizar002.jpg" width="95" height="18" border="0"></img></a><%else%><img src="bt_visualizar002.jpg" width="95" height="18" border="0"></img><%end if%></td>
                  <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("origem_franquia")%></font></div></td>
				  <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("List_Name")%></font></div></td>
                <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Admin_ID")%></font></div></td>
                <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Admin_Pass")%></font></div></td>
			    <td width="130" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("Permissao")%></font></div></td>
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
</table>
</form>








 <%else%>
 
 
 
 
 <div align="center"><a href="javascript:newWindow2('form_incluir_senha.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a></div>
 
 
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I>Senha não encontrada</div>
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
  

  
<!--#include file="dsn2.asp"-->
</body>
</html>
