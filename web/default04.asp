<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->
<% response.buffer=True%>


<html>
<head>
<title>Imobiliária Veja</title>

<!--#include file="style_imoveis02.asp"-->

<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (b2) 








{
if (b2.txt_pergunta.value == "") {
		alert("Você precisa escolher uma opção para que o atendimento seja preferencial.");
		
		b2.txt_pergunta.focus();
		return false;
}
}








	



</script>



</head>
<body onload=document.forms.senha.Admin_ID.focus(); bgcolor="#FFFFFF" vlink="#48576C" link="#48576C" alink="#000000">
<div align="center">

<br>
<img src="default_img01.jpg" border="0" align="middle"></img>
  <p><br>
    
    
    <strong><font color="#9d9249">Para melhor atend&ecirc;-lo, informe se o senhor 
    tem</font></strong></p>
  <p><strong><font color="#9d9249">um im&oacute;vel para vender, alugar ou permutar 
    ?<br>
    <font face="Verdana" size="1"></font></font></strong> </p>
</div>
<form action="resposta_avaliacao.asp"  onSubmit="return isValidDigitNumber(this);" Method="Post" name="b2">
  

  
  
  
  <table border="0" cellpadding="0" cellspacing="0" align="center">
<tr>
      <td bgcolor="<%=claro%>" style="border:1px solid <%=claro%>;"> 
        <table border="0" cellpadding="2" cellspacing="1" align="center" style="border:1px solid #FFFFFF;">

<tr>
            <td bgcolor="#f7ecbf" align="center"> 
              <table border="0" cellpadding="2" cellspacing="0">
                <tr bgcolor="#f7ecbf"> 
                  <td align="right" bgcolor="#f7ecbf">&nbsp;</td>
                  <td bgcolor="#f7ecbf"><input type="radio" name="txt_pergunta" value="sim"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sim, eu tenho  imóvel.</strong></font></td>
</tr>


<tr bgcolor="#f7ecbf"> 
                  <td align="right" bgcolor="#f7ecbf">&nbsp;</td>
                  <td bgcolor="#f7ecbf"><input type="radio" name="txt_pergunta" value="nao"><font color="#9d9249" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Não, eu não tenho imóvel.</strong></font></td>
</tr>
                
                <tr bgcolor="#f7ecbf"> 
                  <td bgcolor="#f7ecbf">&nbsp;</td>
                  <td align="right" bgcolor="#f7ecbf"> <div align="center">
                      <input  type="submit" class="inputSubmit" value="Entrar" size="20" style="HEIGHT: 18px; WIDTH: 120px; background:#9d9249" >
                    </div></td>
</tr>
</table>


</tr>
</table>
</td>
</tr>
</table>
</form>
<center>
</center>


</body>
</html>
<% response.flush%>
  <%response.clear%>
 
<!--#include file="dsn2.asp"-->