

<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCod_imovel,vcheck01,SearchFor,FotoPeqEx
Dim objFSO
Dim objFolder
Dim objFile 
Dim objFiles
Dim strFileName
Dim ImageFilePath 

dim varCidade, varBairro,varNegociacao,varQuartos,page,SearchWhere

 varCod_imovel = request.QueryString("varCod_imovel")
 SearchFor = request.querystring("SearchFor")
 SearchWhere = request.querystring("SearchWhere")
 varCidade = request.querystring("varCidade")
  varBairro = request.querystring("varBairro")  
  varNegociacao = request.querystring("varNegociacao")
  varQuartos = request.querystring("varQuartos")
   page = request.querystring("page")       


ImageFilePath = left(Server.mappath(Request.ServerVariables("PATH_INFO")),Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-24 )

SearchFor = request.Querystring("SearchFor")
 varCod_imovel = request.form("varCod_imovel")
 
 
 
 vcheck01 = request.form("check01")  
 
  if varCod_imovel = "" and vcheck01 = ""  then
 response.Redirect "archive_imoveis.asp?SearchFor="&SearchFor&""
 end if                                                        
															  
	
    Set rs = Server.CreateObject("ADODB.RecordSet")
	
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	Conexao.Open dsn
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	 dim rs2,strSQL2
	  
	  Set rs2 = Server.CreateObject("ADODB.RecordSet")
	  strSQL2="select imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou  from imoveis where cod_imovel in ("& request.form("check01") &")"
   
       rs2.open strSQL2,Conexao
   
       If not rs2.eof Then
	
	Set objFolder = objFSO.GetFolder(left(Server.mappath(Request.ServerVariables("PATH_INFO")),Len(Server.mappath(Request.ServerVariables("PATH_INFO")))-24 )) 
	Set objFiles = objFolder.Files 
	 for Each strFileName in objFiles
	      Do while not rs2.eof
		  
	
		  
		  
	     if rs2("cod_imovel") <>""  then
       
	   
	   if objFSO.FileExists(Server.MapPath(rs2("Foto_pequena"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_pequena"))) <> "imovel00000.jpg" and objFSO.GetFileName(Server.MapPath(rs2("Foto_pequena"))) <> "mini_imovel00000.jpg" then
	    objFSO.DeleteFile(ImageFilePath & "/" &rs2("foto_pequena"))
		end if
	   
	   
	    if objFSO.FileExists(Server.MapPath(rs2("Foto_grande"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande"))) <> "imovel00000.jpg" then
	    objFSO.DeleteFile(ImageFilePath & "/" &rs2("foto_grande"))
		end if
		
		
		
		if objFSO.FileExists(Server.MapPath(rs2("Foto_grande1"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande1"))) <> "imovel00000.jpg" then
	    objFSO.DeleteFile(ImageFilePath & "/" &rs2("foto_grande1"))
		end if
		
		
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande2"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande2"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande2"))
		end if
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande3"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande3"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande3"))
		end if
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande4"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande4"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande4"))
		end if
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande5"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande5"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande5"))
		end if
		
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande6"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande6"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande6"))
		end if
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande7"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande7"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande7"))
		end if
		
		
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande8"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande8"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande8"))
		end if
		
		
		
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande9"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande9"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande9"))
		end if
		
		
		
		
			if objFSO.FileExists(Server.MapPath(rs2("foto_grande10"))) = True and objFSO.GetFileName(Server.MapPath(rs2("Foto_grande10"))) <> "imovel00000.jpg" then
		objFSO.DeleteFile(ImageFilePath & "/"&rs2("foto_grande10"))
		end if
		
		
		
		end if
		   rs2.movenext
		   loop
		   	   
		   next
        Conexao.execute"delete from imoveis where cod_imovel in ("& request.form("check01") &")"
        
		
		
		
		
		
		 
end if
	
	'--------------------------------

             
			 
			 set rs = nothing
			 
			 
	'------------------------------------		 
			 
			 
           '--------------------------------

             rs2.close
			 
			 set rs2 = nothing
			 
			 
	'------------------------------------
	
	
	'--------------------------------

            
			 
			 set objfso = nothing
			 
			 
	'------------------------------------



       '--------------------------------

             conexao.close
			 
			 set conexao = nothing
			 
			 
	'------------------------------------



response.Redirect "archive_imoveis.asp?SearchFor="&SearchFor&"&SearchWhere="&SearchWhere&"&varCidade="&varCidade&"&varBairro="&varBairro&"&varNegociacao="&varNegociacao&"&varQuartos="&varQuartos&"&page="&cInt(Page)&""
			
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">










 
 <%
 
     
 
 
           
          
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


