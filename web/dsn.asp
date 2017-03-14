<%
dim dsn
dim Conn


'dsn="Provider=SQLOLEDB;  Data Source=.; User ID=sa; Password=nico; Initial Catalog=definitivo;"

dsn="Provider=SQLOLEDB;  Data Source=sqlserver01.imobiliariaveja.com.br; User ID=imobiliariaveja; Password=vejasql25a35; Initial Catalog=imobiliariaveja;"


Set Conn = Server.CreateObject("ADODB.Connection")

Conn.Open dsn
%>

