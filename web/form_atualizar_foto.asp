<% response.buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="style_imoveis.asp"-->
<!--#include file="loggedin.asp"-->
<!--#include file="cores.asp"-->
<%
dim varSucesso_cidade,varExistente,vProprietario,varCodImovel
dim varResultado
varResultado = request.QueryString("varResultado")

varSucesso_cidade = request.querystring("varSucesso_cidade")
varExistente = request.querystring("varExistente")

varCodImovel = request.QueryString("varCodImovel")

dim varNumFoto

varNumFoto = request.querystring("varNumFoto")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Atualizar foto</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">


<script>

// Verifica se somente números foram digitados no campo
function isValidDigitNumber (nform) 
{
{






if (nform.txtAnuncio.value == "") {
		alert("Digite o endereço da imagem da foto.");
		nform.txtAnuncio.focus();
		nform.txtAnuncio.select();
		return false;
}






 vfile = b2.txtAnuncio.value;
    tfile = vfile.length;
    
    if (vfile.substr(tfile - 4, 4) != ".jpg" && vfile.substr(tfile - 4, 4) != ".gif" && vfile.substr(tfile - 4, 4) != ".JPG" && vfile.substr(tfile - 4, 4) != ".GIF") {
        alert("O arquivo do formulário Foto deverá possuir o formato (.jpg) ou (.gif)!");
        b2.txtAnuncio.value == vfile.substr(tfile - 4, 4);
		b2.txtAnuncio.focus();
		b2.txtAnuncio.select();
		
		
        return false;
    }
	
	

var strVerif2 = b2.txtAnuncio.value;
var	strVerif_n2 = strVerif2.length;
if (strVerif2.substring(strVerif_n2 - 15,strVerif_n2) == "imovel00000.jpg" ){

       alert("Você escolheu o arquivo errado, imovel00000.jpg não pode ser enviado.");
       b2.txtAnuncio.focus();
		b2.txtAnuncio.select();
		
		
		
return false;

}











}

{






//------------- Verifica se é numérico---------------------



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
//-----------------------------------------------

}


}






</script>



</head>

<body onload=doublecombo.txt_cidade.focus(); bgcolor="<%=escuro%>" bottommargin="0" topmargin="0" leftmargin="0" rightmargin="0" marginheight="0" marginwidth="0">
<form method="post" ENCTYPE="multipart/form-data" onSubmit="return isValidDigitNumber(this);" action="outFile009.asp?varCodImovel=<%=varCodImovel%>&varNumFoto=<%=varNumFoto%>"  name="b2">
  <table width="345" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="345" height="48"><img src="top_resultado.jpg" width="345" height="48"></td>
    </tr>
    <tr>
      <td width="345" height="18"><div align="center">
        
		  
		  
		  
         <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><% response.write varResultado%></font></div></td>
    </tr>
    <tr>
      <td><table width="345" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="5">&nbsp;</td>
            <td><table width="335" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="100" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
                    <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Foto</font></div></td>
                  <td width="235" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"><input name="txtAnuncio" type="file"  size="30" maxlength="23" align="left" class="inputBox" style="HEIGHT: 18px; WIDTH: 235px; background: <%=claro%>; "></td>
                </tr>
                <tr>
                  <td width="100">&nbsp;</td>
                  <td width="235"><table width="235" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="117"><input name="image" type="image"  src="bt_enviar003.jpg" width="117" height="18" border="0"></td>
                        <td><a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar003.jpg" width="118" height="18" border="0"></a></td>
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
</body>
</html>
<% response.flush%>
  <%response.clear%>
