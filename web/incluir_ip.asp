<%
Option Explicit
%>
<!--#include file="dsn.asp"-->


<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEmail,vTelefone,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vVagas,vValor,vNegociacao,vFoto
Dim vdata2
dim vVila

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
 
 
 
  
	
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
    
	
   
	dim vNome,vIP,vPermissao
	dim vQuem_Incluiu
	dim vOrigem_Franquia
	dim vSenha
			
	
	vIP = request.form("txt_ip")
	 vQuem_Incluiu = request.form("txt_quem_incluiu")
	 vOrigem_Franquia = request.form("txt_origem_franquia")
	 vSenha = request.form("txt_senha")
	 
	
	Conexao.execute"Insert into ip(ip,quem_incluiu,origem_franquia,senha_incluiu) values('"& vIP &"','"& vQuem_Incluiu &"','"& session("vOrigem_Franquia") &"','"& vSenha &"')"
	 
	 dim varCidade
	 response.Redirect "form_incluir_ip.asp?varSucesso_bairro="&vIP&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>













 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
