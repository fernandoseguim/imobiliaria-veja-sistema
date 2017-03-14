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


	Dim strID, strPassword,vOrigem_Franquia
	strID = Request("Admin_ID")
	strPassword = Request("Password")
	vOrigem_Franquia = request("vOrigem_Franquia")
	session("vOrigem_Franquia") = vOrigem_Franquia
    session("Password") = strPassword
	Dim rsUsers
	set rsUsers = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM Senha WHERE Admin_ID = '" & strID & "';"
	rsUsers.Open strSQL, objConn

	If rsUsers.EOF Then
		Session("Admin_ID") = Request("Admin_ID")
		Response.Redirect "senha02.asp?SecondTry=True"
	Else
		While Not rsUsers.EOF
			If UCase(rsUsers("Admin_Pass")) = UCase(strPassword) Then
				Session("Admin_ID") = rsUsers("Admin_ID")
				Session("permissao") = rsUsers("permissao")
				Session("isLoggedIn") = True
				Session("blnValidUser") = True
				Response.Redirect "archive_senha.asp"
			Else
				rsUsers.MoveNext
			End If
		Wend
		Session("Admin_ID") = Request("Admin_ID")
		Response.Redirect "senha02.asp?SecondTry=True&WrongPW=True"
	End If
%>
<!--#include file="dsn2.asp"-->