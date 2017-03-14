<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin03.asp"-->

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
<br>
<center><a href="archive_franquia.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000"><strong>Carregar 
  Página</strong></font></a> 
</center>
<div align="left"></div>

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




dim varCodFranquia

Dim colorchanger
Dim color1
Dim color2
colorchanger = 0
color1 = medio    
color2 = claro





Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	 
   Conexao.Open dsn
   

       
		


SQL = "Select franquia.id_franquia,franquia.nome_franquia,franquia.data_franquia from franquia  ORDER BY id_franquia DESC"

rs.Open SQL, Conexao

%>

  <br>
  <br>
  <%
If NOT (rs.BOF AND rs.EOF) Then
%>
</center>
<center>
<form  Method="Post" name="Formulario" action="multi_excluir_franquia.asp?SearchFor=<%=SearchFor%>" >           
<table width="510" border="0" cellspacing="0" cellpadding="0">
  <tr>
   
		
            <td width="20" height="18" bgcolor="<%=claro%>">
<input type="checkbox" name="selTodos" onclick="check(true);"></td>
            <td width="95" height="18"><input name="image" type="image" src="bt_excluir002.jpg" width="95" height="18" border="0"></td>
             
            <td width="95" height="18"><a href="javascript:newWindow2('form_incluir_franquia.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a></td>
           
		    <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Franquia</strong></font></div></td>
		    <td width="150" height="18" bgcolor="#000000" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data</strong></font></div></td>
		   
           
		</tr>
      
	   <%








Do While not rs.eof

'------------------------------------------------

%>
	 
	 
	      
				 
				<tr>
				<% varCodFranquia = rs("id_franquia") %>
                <td width="20" height="18" bgcolor="<%=color1%>"><input type="checkbox" name="check01" value="<%=rs("id_franquia")%>"></td>
                 		 
				  <td width="95" height="18" bgcolor="<%=color1%>"><a href="excluir_franquia.asp?varCodFranquia=<%=varCodFranquia%>&SearchFor=<%=SearchFor%>"><img src="bt_excluir001.jpg" width="95" height="18" border="0"></img></a></td>
                  <td width="95" height="18" bgcolor="<%=color1%>"><a href="javascript:newWindow2('visualizar_franquia.asp?varCodFranquia=<%=varCodFranquia%>&SearchFor=<%=SearchFor%>')"><img src="bt_visualizar002.jpg" width="95" height="18" border="0"></img></a></td>
                  <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("nome_franquia")%></font></div></td>
				   <td width="150" height="18" bgcolor="<%=color1%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("data_franquia")%></font></div></td>
				 
				 
               
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
 <center>
  <a href="javascript:newWindow2('form_incluir_franquia.asp')"><img src="bt_incluir001.jpg" width="95" height="18" border="0"></a>
</center>
 <br>
 <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
<div align="center"I>Franquia não encontrada</div>
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
