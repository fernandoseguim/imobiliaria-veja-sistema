

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,var1,var2,rs,vtxt_Nome,vtxt_Admin_id, vtxt_Admin_pass, vtxt_Ano ,vtxt_1_1 ,vtxt_1_2,vtxt_1_3 ,vtxt_1_4,vtxt_1_5 ,vtxt_1_6,vtxt_1_7 ,vtxt_2_1 ,vtxt_2_2,vtxt_2_3 ,vtxt_2_4,vtxt_2_5 ,vtxt_2_6,vtxt_2_7 ,vtxt_3_1 ,vtxt_3_2,vtxt_3_3 ,vtxt_3_4, vtxt_3_5 ,vtxt_3_6,vtxt_3_7 ,vtxt_4_1 ,vtxt_4_2,vtxt_4_3 ,vtxt_4_4, vtxt_4_5 ,vtxt_4_6,vtxt_4_7 ,vtxt_5_1 ,vtxt_5_2,vtxt_5_3 ,vtxt_5_4, vtxt_5_5 ,vtxt_5_6,vtxt_5_7 ,vtxt_6_1 ,vtxt_6_2,vtxt_6_3 ,vtxt_6_4, vtxt_6_5 ,vtxt_6_6,vtxt_6_7 ,vtxt_7_1 ,vtxt_7_2,vtxt_7_3 ,vtxt_7_4, vtxt_7_5 ,vtxt_7_6,vtxt_7_7 ,vtxt_8_1 ,vtxt_8_2,vtxt_8_3 ,vtxt_8_4, vtxt_8_5 ,vtxt_8_6,vtxt_8_7 ,vtxt_9_1 ,vtxt_9_2,vtxt_9_3 ,vtxt_9_4, vtxt_9_5 ,vtxt_9_6,vtxt_9_7 ,vtxt_10_1 ,vtxt_10_2,vtxt_10_3 ,vtxt_10_4,vtxt_10_5 ,vtxt_10_6,vtxt_10_7 
Dim vtxt_11_1 ,vtxt_11_2,vtxt_11_3 ,vtxt_11_4,vtxt_11_5 ,vtxt_11_6,vtxt_11_7 ,vtxt_12_1 ,vtxt_12_2,vtxt_12_3 ,vtxt_12_4,vtxt_12_5 ,vtxt_12_6,vtxt_12_7 ,vtxt_13_2 ,vtxt_13_3,vtxt_13_4 ,vtxt_13_5,vtxt_13_6 ,vtxt_13_7, vtxt_endereco, vtxt_telefone, vtxt_venc_1, vtxt_venc_2, vtxt_venc_3, vtxt_venc_4, vtxt_venc_5, vtxt_venc_6 , vtxt_venc_7, vtxt_venc_8, vtxt_venc_9, vtxt_venc_10, vtxt_venc_11, vtxt_venc_12 , vtxt_rep_1, vtxt_rep_2, vtxt_rep_3, vtxt_rep_4, vtxt_rep_5, vtxt_rep_6 , vtxt_rep_7, vtxt_rep_8, vtxt_rep_9, vtxt_rep_10, vtxt_rep_11, vtxt_rep_12

 
 
   
  
      
	
	  vtxt_nome = request.form("txt_nome")  
	  vtxt_admin_id = request.form("txt_admin_id")  
	  vtxt_admin_pass = request.form("txt_admin_pass")  
	  vtxt_ano = request.form("txt_ano")  
	  
	   vtxt_1_1 = request.form("txt_1_1")
	  vtxt_1_2 = request.form("txt_1_2") 
	  vtxt_1_3 = request.form("txt_1_3") 
	  vtxt_1_4 = request.form("txt_1_4") 
	  vtxt_1_5 = request.form("txt_1_5") 
	  vtxt_1_6 = request.form("txt_1_6") 
	  vtxt_1_7 = request.form("txt_1_7") 
	  
	 
	 
	  vtxt_2_1 = request.form("txt_2_1")
	  vtxt_2_2 = request.form("txt_2_2") 
	  vtxt_2_3 = request.form("txt_2_3") 
	  vtxt_2_4 = request.form("txt_2_4") 
	  vtxt_2_5 = request.form("txt_2_5") 
	  vtxt_2_6 = request.form("txt_2_6") 
	  vtxt_2_7 = request.form("txt_2_7") 
	
	  
	  
	   vtxt_3_1 = request.form("txt_3_1")
	  vtxt_3_2 = request.form("txt_3_2") 
	  vtxt_3_3 = request.form("txt_3_3") 
	  vtxt_3_4 = request.form("txt_3_4") 
	  vtxt_3_5 = request.form("txt_3_5") 
	  vtxt_3_6 = request.form("txt_3_6") 
	  vtxt_3_7 = request.form("txt_3_7") 
	  
	  
	  
	   vtxt_4_1 = request.form("txt_4_1")
	  vtxt_4_2 = request.form("txt_4_2") 
	  vtxt_4_3 = request.form("txt_4_3") 
	  vtxt_4_4 = request.form("txt_4_4") 
	  vtxt_4_5 = request.form("txt_4_5") 
	  vtxt_4_6 = request.form("txt_4_6") 
	  vtxt_4_7 = request.form("txt_4_7") 
	 
	  
	   vtxt_5_1 = request.form("txt_5_1")
	  vtxt_5_2 = request.form("txt_5_2") 
	  vtxt_5_3 = request.form("txt_5_3") 
	  vtxt_5_4 = request.form("txt_5_4") 
	  vtxt_5_5 = request.form("txt_5_5") 
	  vtxt_5_6 = request.form("txt_5_6") 
	  vtxt_5_7 = request.form("txt_5_7") 
	
	  
	  
	  
	    vtxt_6_1 = request.form("txt_6_1")
	  vtxt_6_2 = request.form("txt_6_2") 
	  vtxt_6_3 = request.form("txt_6_3") 
	  vtxt_6_4 = request.form("txt_6_4") 
	  vtxt_6_5 = request.form("txt_6_5") 
	  vtxt_6_6 = request.form("txt_6_6") 
	  vtxt_6_7 = request.form("txt_6_7") 
	
	  
	    vtxt_7_1 = request.form("txt_7_1")
	  vtxt_7_2 = request.form("txt_7_2") 
	  vtxt_7_3 = request.form("txt_7_3") 
	  vtxt_7_4 = request.form("txt_7_4") 
	  vtxt_7_5 = request.form("txt_7_5") 
	  vtxt_7_6 = request.form("txt_7_6") 
	  vtxt_7_7 = request.form("txt_7_7") 
	 
	  
	  
	   vtxt_8_1 = request.form("txt_8_1")
	  vtxt_8_2 = request.form("txt_8_2") 
	  vtxt_8_3 = request.form("txt_8_3") 
	  vtxt_8_4 = request.form("txt_8_4") 
	  vtxt_8_5 = request.form("txt_8_5") 
	  vtxt_8_6 = request.form("txt_8_6") 
	  vtxt_8_7 = request.form("txt_8_7") 
	 
	  
	  
	  
	   vtxt_9_1 = request.form("txt_9_1")
	  vtxt_9_2 = request.form("txt_9_2") 
	  vtxt_9_3 = request.form("txt_9_3") 
	  vtxt_9_4 = request.form("txt_9_4") 
	  vtxt_9_5 = request.form("txt_9_5") 
	  vtxt_9_6 = request.form("txt_9_6") 
	  vtxt_9_7 = request.form("txt_9_7") 
	 
	  
	  vtxt_10_1 = request.form("txt_10_1")
	  vtxt_10_2 = request.form("txt_10_2") 
	  vtxt_10_3 = request.form("txt_10_3") 
	  vtxt_10_4 = request.form("txt_10_4") 
	  vtxt_10_5 = request.form("txt_10_5") 
	  vtxt_10_6 = request.form("txt_10_6") 
	  vtxt_10_7 = request.form("txt_10_7") 
	  
	
	
	vtxt_11_1 = request.form("txt_11_1")
	  vtxt_11_2 = request.form("txt_11_2") 
	  vtxt_11_3 = request.form("txt_11_3") 
	  vtxt_11_4 = request.form("txt_11_4") 
	  vtxt_11_5 = request.form("txt_11_5") 
	  vtxt_11_6 = request.form("txt_11_6") 
	  vtxt_11_7 = request.form("txt_11_7") 
	 
	  
	  
	  
	  vtxt_12_1 = request.form("txt_12_1")
	  vtxt_12_2 = request.form("txt_12_2") 
	  vtxt_12_3 = request.form("txt_12_3") 
	  vtxt_12_4 = request.form("txt_12_4") 
	  vtxt_12_5 = request.form("txt_12_5") 
	  vtxt_12_6 = request.form("txt_12_6") 
	  vtxt_12_7 = request.form("txt_12_7") 
	 
	  
	  
	  
	  vtxt_13_2 = request.form("txt_13_2") 
	  vtxt_13_3 = request.form("txt_13_3") 
	  vtxt_13_4 = request.form("txt_13_4") 
	  vtxt_13_5 = request.form("txt_13_5") 
	  vtxt_13_6 = request.form("txt_13_6") 
	 
	  '---------------------------------------------------
	  
	   vtxt_endereco = request.form("txt_endereco") 
	  vtxt_telefone = request.form("txt_telefone") 
	 
	  vtxt_venc_1 = request.form("txt_venc_1") 
	  vtxt_venc_2 = request.form("txt_venc_2") 
	  vtxt_venc_3 = request.form("txt_venc_3") 
	  
	   vtxt_venc_4 = request.form("txt_venc_4") 
	  vtxt_venc_5 = request.form("txt_venc_5") 
	  vtxt_venc_6 = request.form("txt_venc_6")
	  
	   vtxt_venc_7 = request.form("txt_venc_7") 
	  vtxt_venc_8 = request.form("txt_venc_8") 
	  vtxt_venc_9 = request.form("txt_venc_9")
	  
	   vtxt_venc_10 = request.form("txt_venc_10") 
	  vtxt_venc_11 = request.form("txt_venc_11") 
	  vtxt_venc_12 = request.form("txt_venc_12")
	  
	  vtxt_rep_1 = request.form("txt_rep_1") 
	  vtxt_rep_2 = request.form("txt_rep_2") 
	  vtxt_rep_3 = request.form("txt_rep_3") 
	  
	   vtxt_rep_4 = request.form("txt_rep_4") 
	  vtxt_rep_5 = request.form("txt_rep_5") 
	  vtxt_rep_6 = request.form("txt_rep_6")
	  
	   vtxt_rep_7 = request.form("txt_rep_7") 
	  vtxt_rep_8 = request.form("txt_rep_8") 
	  vtxt_rep_9 = request.form("txt_rep_9")
	  
	   vtxt_rep_10 = request.form("txt_rep_10") 
	  vtxt_rep_11 = request.form("txt_rep_11") 
	  vtxt_rep_12 = request.form("txt_rep_12")
	  
	  
	  
	  
	  
	  
	  '------------------------------------------------------
	  
	  
	  					  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	var1 = "insert into admin(admin_id, admin_pass, nome, ano , d1_1, d1_2, d1_3, d1_4, d1_5, d1_6, d1_7,  d2_1, d2_2, d2_3, d2_4, d2_5, d2_6, d2_7,  d3_1, d3_2, d3_3, d3_4, d3_5, d3_6, d3_7,  d4_1, d4_2, d4_3, d4_4, d4_5, d4_6, d4_7,  d5_1, d5_2, d5_3, d5_4, d5_5, d5_6, d5_7,  d6_1, d6_2, d6_3, d6_4, d6_5, d6_6, d6_7,  d7_1, d7_2, d7_3, d7_4, d7_5, d7_6, d7_7,  d8_1, d8_2, d8_3, d8_4, d8_5, d8_6, d8_7,  d9_1, d9_2, d9_3, d9_4, d9_5, d9_6, d9_7, d10_1, d10_2, d10_3, d10_4, d10_5, d10_6, d10_7, d11_1, d11_2, d11_3, d11_4, d11_5, d11_6, d11_7, d12_1, d12_2, d12_3, d12_4, d12_5, d12_6, d12_7, d13_2, d13_3, d13_4, d13_5, d13_6 , telefone, endereco, venc_1, venc_2, venc_3, venc_4, venc_5, venc_6, venc_7, venc_8, venc_9, venc_10, venc_11, venc_12  ) values"
	var2 ="('"&vtxt_admin_id&"','"&vtxt_admin_pass&"','"&vtxt_nome&"','"&vtxt_ano&"','"&vtxt_1_1&"','"&vtxt_1_2&"','"&vtxt_1_3&"','"&vtxt_1_4&"','"&vtxt_1_5&"','"&vtxt_1_6&"','"&vtxt_1_7&"','"&vtxt_2_1&"','"&vtxt_2_2&"','"&vtxt_2_3&"','"&vtxt_2_4&"','"&vtxt_2_5 &"','"&vtxt_2_6&"','"&vtxt_2_7&"','"&vtxt_3_1&"','"&vtxt_3_2&"','"&vtxt_3_3&"','"&vtxt_3_4&"','"&vtxt_3_5&"','"&vtxt_3_6&"','"&vtxt_3_7&"','"&vtxt_4_1&"','"&vtxt_4_2&"','"& vtxt_4_3&"','"&vtxt_4_4&"','"&vtxt_4_5&"','"&vtxt_4_6&"','"&vtxt_4_7&"','"&vtxt_5_1&"','"&vtxt_5_2&"','"&vtxt_5_3&"','"&vtxt_5_4&"','"&vtxt_5_5&"','"&vtxt_5_6&"','"&vtxt_5_7&"','"&vtxt_6_1&"','"&vtxt_6_2&"','"&vtxt_6_3&"','"&vtxt_6_4&"','"&vtxt_6_5&"','"&vtxt_6_6&"','"&vtxt_6_7&"','"&vtxt_7_1&"','"&vtxt_7_2&"','"&vtxt_7_3&"','"&vtxt_7_4&"','"&vtxt_7_5&"','"&vtxt_7_6&"','"&vtxt_7_7&"','"&vtxt_8_1&"','"&vtxt_8_2&"','"&vtxt_8_3&"','"&vtxt_8_4&"','"&vtxt_8_5&"','"&vtxt_8_6&"','"&vtxt_8_7&"','"&vtxt_9_1&"','"&vtxt_9_2&"','"&vtxt_9_3&"','"&vtxt_9_4&"','"&vtxt_9_5&"','"&vtxt_9_6&"','"&vtxt_9_7&"','"&vtxt_10_1&"','"&vtxt_10_2&"','"&vtxt_10_3&"','"&vtxt_10_4&"','"&vtxt_10_5&"','"&vtxt_10_6&"','"&vtxt_10_7&"','"&vtxt_11_1&"','"&vtxt_11_2&"','"&vtxt_11_3&"','"&vtxt_11_4&"','"&vtxt_11_5&"','"&vtxt_11_6&"','"&vtxt_11_7&"','"&vtxt_12_1&"','"&vtxt_12_2&"','"&vtxt_12_3&"','"&vtxt_12_4&"','"&vtxt_12_5&"','"&vtxt_12_6&"','"&vtxt_12_7&"','"&vtxt_13_2&"','"&vtxt_13_3&"','"&vtxt_13_4&"','"&vtxt_13_5&"','"&vtxt_13_6&"','"& vtxt_telefone &"','"& vtxt_endereco &"','"& vtxt_venc_1 &"' ,'"& vtxt_venc_2 &"' ,'"& vtxt_venc_3 &"' ,'"& vtxt_venc_4 &"' ,'"& vtxt_venc_5 &"' ,'"& vtxt_venc_6 &"' ,'"& vtxt_venc_7 &"' ,'"& vtxt_venc_8 &"' ,'"& vtxt_venc_9 &"' ,'"& vtxt_venc_10 &"' ,'"& vtxt_venc_11 &"' ,'"& vtxt_venc_12 &"')"
	
	Conexao.execute""&var1&""&var2&"" 
	 
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48" ><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
<td width="590" height="105" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    
          <td width="202" height="156" ></img><img src="../proposta/sorriso_proposta2.jpg" width="202" height="156"></td>
          <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36" ></img></td>

</tr>


</table>







 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>

