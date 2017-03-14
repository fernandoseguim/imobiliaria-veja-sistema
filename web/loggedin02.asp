<%
Session.TimeOut = 480
	If Session("isLoggedIn") <> True or session("permissao") <> "6" Then
		Response.Redirect "senha02.asp"
	End If
%>