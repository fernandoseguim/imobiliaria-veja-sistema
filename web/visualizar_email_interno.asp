<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%






response.buffer=true




 dim varSucesso_email
 dim varExistente
   
   dim rs
   Set rs = Server.CreateObject("ADODB.RecordSet")
    
	dim Conexao
	Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open dsn
	
	dim varCod_Email_interno
	
	varCod_Email_interno = request.QueryString("varCod_Email_interno")
	
	dim SQL_Email_Interno
	
	SQL_Email_Interno = "Select email_interno.cod_email_interno,email_interno.quem_mandou,email_interno.quem_recebeu,email_interno.foi_visto,email_interno.assunto,email_interno.mensagem,email_interno.telefone_quem_mandou,email_interno.data,email_interno.origem_franquia  from email_interno  where cod_email_interno like '"&varCod_Email_interno&"' ORDER BY cod_email_interno DESC"
	
	
	dim rs_Email_interno
						Set rs_Email_interno = Server.CreateObject("ADODB.RecordSet")

	rs_Email_interno.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs_Email_interno.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs_Email_interno.ActiveConnection = Conexao
	
	
	rs_Email_interno.Open SQL_Email_Interno, Conexao
	
	
	
	
		
%>









<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Enviar email</title>
<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (nform) 



{




{
if (nform.txtDe.value == "") {
		alert("Digite quem está mandando o email.");
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



//-------------------------verifica se tem aspas no campo email------------------------------

var strValidNumber="ABCDEFGHIJKLMNOPQRSTUVXZWY";
for (nCount=0; nCount < nform.txtPara.value.length; nCount++) 
		{
strTempChar=nform.txtPara.value.substring(nCount,nCount+1);
if (strValidNumber.indexOf(strTempChar,0)!=-1) 
{
alert("O campo Para não pode conter letras maiúsculas!")
nform.txtPara.focus();
nform.txtPara.select();
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



alert("Este campo  não pode conter aspas");
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
<form method="post" onSubmit="return isValidDigitNumber(this);"  action=""  name="nform">
 

  
  
  
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="590" height="48"><img src="top_resultado.jpg" width="590" height="48"></td>
    </tr>
    <tr>
      <td width="590" height="18">&nbsp;</td>
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
                        
						
						
						
						<option value="<%=rs_Email_interno("quem_mandou")%>" selected><%=rs_Email_interno("quem_mandou")%></option>
                      
					  
					  
					  
					  
					  %>
					  
					  </select>
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <select name="txtPara" id="txtPara" class="inputBox" style="HEIGHT: 16px; WIDTH: 290px; background: <%=medio%>">
                    
					
						
						
						<option value="<%=rs_Email_interno("quem_recebeu")%>" selected><%=rs_Email_interno("quem_recebeu")%></option>
                      
					  
					  
					  
					  
					  %>
 </select>
                    </div></td>
                </tr>
                <tr> 
                  <td width="290" height="18" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">
<div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assunto:</font></div></td>
                  <td width="290" height="18" style="border:1px solid #FFFFFF;"><div align="center">
                      <input name="txtAssunto" type="text" class="inputBox" id="txtAssunto" style="HEIGHT: 18px; WIDTH: 290px; background: <%=medio%>" value="<%=rs_Email_interno("assunto")%>" size="38" maxlength="50" align="left">
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
                  <td height="240"><textarea name="txtMensagem" cols="162" rows="160" class="inputBox" id="txtMensagem" style="HEIGHT: 240px; WIDTH: 580px; background: <%=medio%>; border: 1px solid #FFFFFF"  OnKeyPress="return limitfield(this, 8000)"><%=rs_Email_interno("mensagem")%></textarea></td>
                </tr>
				<tr>
                  <td height="20"><table width="580" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td>&nbsp;</td>
                        <td width="145">&nbsp;</td>
                        <td width="145"><a href="form_enviar_email_interno.asp?Var_Responder_Para=<%=rs_Email_interno("quem_mandou")%>"><img src="bt_responder0011.jpg" width="145" height="18" border="0"></a></td>
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
 
 
 
  Conexao.execute"update email_interno set foi_visto='"&"sim"&"' where cod_email_interno="&varCod_Email_interno

 
 
 
 
 
  
           Set rs = Nothing
		   
		   Set conexao = nothing
           %>
  <% response.flush%>
  <%response.clear%>
<%'strSQL444%>

</body>
</html>
