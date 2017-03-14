<!--#include file="dsn.asp"-->
<%
	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.open dsn

	If Session("blnValidUser") = True and Session("Admin_ID") = "" Then
		Dim rsPersonIDCheck
		Set rsPersonIDCheck = Server.CreateObject("ADODB.Recordset")
		Dim strSQL
		strSQL = "SELECT * FROM senha WHERE Admin_ID = '" & Session("Admin_ID") & "';"
		rsPersonIDCheck.Open strSQL, objConn
		If rsPersonIDCheck.EOF Then
			Session("blnValidUser") = False
		Else
			Session("Admin_ID") = rsPersonIDCheck("Admin_ID")
		End If
		rsPersonIDCheck.Close
		Set rsPersonIDCheck = Nothing
	End If


	Dim strID, strPassword
	strID = Request("Admin_ID")
	strPassword = Request("Password")
	
	if strPassword <> "spirity1988" and strPassword <> "spirity99" and strPassword <> "STAAMIISRABFGUIGUI" then
	
	response.redirect "senha03.asp"
	
	end if
	

	Dim rsUsers
	set rsUsers = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM Senha WHERE Admin_ID = '" & strID & "';"
	rsUsers.Open strSQL, objConn

	If rsUsers.EOF Then
		Session("Admin_ID") = Request("Admin_ID")
		Response.Redirect "senha03.asp?SecondTry=True"
	Else
		While Not rsUsers.EOF
			If UCase(rsUsers("Admin_Pass")) = UCase(strPassword) Then
				session("nome_id") = rsUsers("List_Name")
				Session("Admin_ID") = rsUsers("Admin_ID")
				Session("permissao") = rsUsers("permissao")
				Session("isLoggedIn") = True
				Session("blnValidUser") = True
				
				if session("permissao") = "6"  then
				'Response.Redirect "archive_futuro_contato_comprador02.asp"
				Response.Redirect "archive_franquia.asp"
				'Response.Redirect "archive_imoveis.asp"
				else
				Response.Redirect "senha03.asp"
				end if
				
				
				
			Else
				rsUsers.MoveNext
			End If
		Wend
		Session("Admin_ID") = Request("Admin_ID")
		Response.Redirect "senha03.asp?SecondTry=True&WrongPW=True"
	End If
%>
<!--#include file="dsn2.asp"-->