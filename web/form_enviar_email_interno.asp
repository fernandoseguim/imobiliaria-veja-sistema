<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->

<%






response.buffer=true


dim  Var_Responder_Para

Var_Responder_Para = request.QueryString("Var_Responder_Para")

 dim varSucesso_email
 dim varExistente
   
   dim rs
   Set rs = Server.CreateObject("ADODB.RecordSet")
    
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open dsn
		
%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Enviar email</title>
<script>

// Verifica se somente n�meros foram digitados no campo
function isValidDigitNumber (nform) 



{




{
if (nform.txtDe.value == "") {
		alert("Digite quem est� mandando o email.");
		nform.txtDe.focus();
		nform.txtDe.select();
		return false;
}
}





{
if (nform.txtPara.value == "") {
		alert("Digite para quem deseja mandar o email.");
		nform.txtPara.focus();
		nform.txtPara.select();
		return false;
}
}







{
if (nform.txtAssunto.value == "") {
		alert("Digite o assunto.");
		nform.txtAssunto.focus();
		nform.txtAssunto.select();
		return false;
}
}

{
if (nform.txtMensagem.value == "") {
		alert("Digite sua mensagem.");
		nform.txtMensagem.focus();
		nform.txtMensagem.select();
		return false;
}
}





//-----------------------------------------------------------------------------



//-----------------------------------------------------------------------------

//-----------------------------------------------------------------------------






var elem=nform.elements;

for (nCount=0; nCount < elem.length; nCount++)
  

	if(elem[nCount].type.indexOf("text")==0)	{
	var strValidNumber12_1="'";
	for (nCount2=0; nCount2 < elem[nCount].value.length; nCount2++) 
	{
	strTempChar12_1=elem[nCount].value.substring(nCount2,nCount2+1);
if (strValidNumber12_1.indexOf(strTempChar12_1,0)!=-1) {



alert("Este campo  n�o pode conter aspas");
elem[nCount].focus();
elem[nCount].select();
return false;
}
}
}














	
}












</script>

<script>
function limitfield(what,limit){
if (what.value.length>=limit)
return false
}
</script>



</head>

<!--#include file="style2_sugestoes.asp"-->
<body onload=b2.txtNome.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">



<form method="post" onSubmit="return isValidDigitNumber(this);"  action="incluir_email_interno.asp?"  name="nform">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
    </tr>
    <tr>
      <td width="590" height="30"> 
        <%
	  dim varSucesso
	  
	  varSucesso = request.QueryString("varSucesso")
	   if varSucesso <> "" then
	   %>
        <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucesso%></font></div>
	   <%
	   
	   end if
	  
	  
	  %>
	  
	  
	  
	  
	  </td>
    </tr>
    <tr>
      <td width="590" height="54"><table width="590" height="54" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5" height="54">&nbsp;</td>
            <td width="580" height="54"><table width="580" height="54" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;" >
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">D&ecirc;:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center"> 
                      <select name="txtDe" id="txtDe" class="inputBox" style="HEIGHT: 16px; WIDTH: 290px; background: <%=medio%>">
                        
						<%
						dim SQL_DE
						
						SQL_DE = "Select senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id,senha.origem_franquia  from senha where admin_id like '"&session("Admin_ID")&"' and admin_pass like '"&session("Password")&"'  ORDER BY ID DESC" 
						
						dim rs_DE
						Set rs_DE = Server.CreateObject("ADODB.RecordSet")

	rs_DE.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs_DE.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs_DE.ActiveConnection = Conexao
	
	
	rs_DE.Open SQL_DE, Conexao



						
						if  not rs_DE.eof then
						
						%>
						
						
						
						
						
						<option value="<%=rs_DE("admin_id")%>" selected><%=rs_DE("admin_id")%></option>
                      
					  <%
					   end if
					  
					  
					  
					  rs_DE.close

                       set rs_DE = nothing
					  
					  
					  
					  %>
					  
					  </select>
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <select name="txtPara" id="txtPara" class="inputBox" style="HEIGHT: 16px; WIDTH: 290px; background: <%=medio%>">
                    
					<%
						dim SQL_PARA
						
						SQL_PARA = "Select senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id,senha.origem_franquia  from senha   ORDER BY ID DESC" 
						
						dim rs_PARA
						Set rs_PARA = Server.CreateObject("ADODB.RecordSet")

	rs_PARA.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor � criado
'se no cliente ou no servidor.

rs_PARA.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e grava��o.

rs_PARA.ActiveConnection = Conexao
	
	
	rs_PARA.Open SQL_PARA, Conexao



						
						
						
						%>
						
						
						
						<% if not rs_PARA.eof then %>
                      <% While NOT rs_PARA.EoF %>
						
						<option value="<%=rs_PARA("admin_id")%>" <% if rs_PARA("admin_id") = Var_Responder_Para then response.write "selected" end if %>><%=rs_PARA("admin_id")%></option>
                      
					  <% rs_PARA.MoveNext %>
                      <% Wend %>
					  <%end if%>
					  
					  
					  <%
					  rs_PARA.close

                       set rs_PARA = nothing
					  
					  
					  
					  %>
 </select>
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" id="txtAssunto" size="38" maxlength="50" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>">
                    </div></td>
                </tr>
              </table></td>
            <td width="5" height="54">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td width="590" height="140"><table width="590" height="300" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td width="580"><table width="580" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Mensagem</font></div></td>
                </tr>
                <tr>
                  <td height="240"><textarea name="txtMensagem" cols="162" rows="160" class="inputBox" id="txtMensagem" style="HEIGHT: 240px; WIDTH: 580px; background: <%=medio%>; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 8000)"></textarea></td>
                </tr>
				<tr>
                  <td height="20"><table width="580" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>&nbsp;</td>
                        <td width="145"><input name="image" type="image" src="bt_enviar0011.jpg" width="145" height="18"></td>
                        <td width="145"><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar0011.jpg" width="145" height="18" border="0"></a></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>
 <%
  
           Set rs = Nothing
		   
		   Set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>

<%
 
 
					


%>


<%'strSQL444%>
<%'Var_Responder_Para%>
<%'session("Admin_id")%>
</body>
</html>
