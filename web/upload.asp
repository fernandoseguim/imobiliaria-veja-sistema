
<%
' Author Philippe Collignon
' Email PhCollignon@email.com

Sub BuildUploadRequest(RequestBin)
    'RequestBin � a vari�vel que cont�m todas as informa��es
	'recebidas dos formul�rios de maneira bin�ria.
	'Get the boundary
	
	PosBeg = 1
	'PosBeg � a posi��o inicial da string bin�ria.
	
	
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	'PosBeg � a posi��o inicial da string a ser pesquisada
	'requestBin � a string a ser pesquisada
	'chr(13) � a substring a ser procurada na string RequestBin
	'PosEnd � a posi��o da primeira ocorr�ncia da substring chr(13) na
	'string RequestBin
	
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	'boundary � uma substring de RequestBin 
	'RequestBin � a string principal
	'PosBeg � a posi��o inicial da substring a ser capturada de RequestBin
	' PosEnd � a posi��o final da substring a ser capturada de RequestBin
	
	boundaryPos = InstrB(1,RequestBin,boundary)
	'boundaryPos � a posi��o da primeira ocorr�ncia da substring boundary
	'dentro da string RequestBin
	
	
	'Get all data inside the boundaries
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		'boundaryPos � a posi��o da primeira ocorr�ncia da substring boundary
		'concatenada com "--" na string RequestBin
		
		
		
		'Members variable of objects are put in a dictionary object
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		'um objeto dicion�rio � criado para se colocar as informa��es
		'sobre a string RequestBin que foram dividas acima.
		'Get an object name
		
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		'Pos � a primeira ocorr�ncia da substring "Content-Disposition" dentro da
		'string RequestBin
		'BoundaryPos � a posi��o inicial a partir da qual a substring"Content-Disposition
		'ir� ser pesquisada
		
		
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		'Pos � a posi��o da primeira ocorr�ncia da substring "name=" dentro 
		'de RequestBin
		'Pos que est� entre par�nteses � a posi��o inicial da pesquisa a ser
		'feita
		
		PosBeg = Pos+6
		'PosBeg � a posi��o inicial mais 6.
		
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		'PosEnd � a posi��o da primeira ocorr�ncia da substring "chr(34)"
		'na string RequestBin
		'PosBeg � a posi��o inicial a partir da qual a pesquisa ir� ser feita.
		
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		'Name � uma vari�vel que receber� uma substring de RequestBin
		'que come�a na posi��o PosBeg e termina na posi��o PosEnd-PosBeg
		
		
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		'PosFile � a posi��o da primeira ocorr�ncia da substring "filename=" 
		'dentro da string RequestBin
		'BoundaryPos � a posi��o inicial a partir da qual a pesquisa ir� ser feita.
		
		
		
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		'PosBound � a posi��o da primeira ocorr�ncia da substring boundary dentro
		'da string RequestBin
		'PosEnd � a posi��o inicial a partir da qual a pesquisa ir� ser feita.
		
		'Test if object is of file type
		If  PosFile<>0 AND (PosFile<PosBound) Then
		    'aqui ser� verificado se os dados dos formul�rios s�o de 
			'arquivos ou dados comuns.
			'Get Filename, content-type and content of file
			
			PosBeg = PosFile + 10
			'PosBeg � a posi��o inicial a partir da qual o nome do arquivo ir�
			'ser capturado da string RequestBim
			'PosFile � a posi��o da primeira ocorr�ncia de "filename=" dentro da
			'string RequestBin 
			
			
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			'PosEnd � a posi��o da primeira ocorr�ncia da substring "chr(34)" dentro
			'da string RequestBin
			'PosBeg � a posi��o inicial a partir da qual a pesquisa ir� ser feita.
			
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'FileName � o nome do arquivo que ser� capturado da string RequestBin
			'PosBeg � a posi��o inicial da pesquisa
			'PosEnd-PosBeg � a posi��o final da pesquisa
			'MidB � a fun��o que captura a substring que cont�m o nome do arquivo
			
			UploadControl.Add "FileName", FileName
			'o nome do arquivo � jogado num dicion�rio de dados.
			'
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			'Pos � a posi��o da primeira ocorr�ncia da substring "Content-Type:" 
			'dentro da substring RequestBin
			'PosEnd � a posi��o incial da pesquisa que ir� ser feita.
			
			PosBeg = Pos+14
			'PosBeg � a posi��o inicial da substring que ir� ser capturada de
			'RequestBin que corresponder� ao tipo de conte�do do arquivo.
			
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			'PosEnd � a posi��o da primeira ocorr�ncia da substring "chr(13)" dentro
			' da string RequestBin
			
			
			'Add content-type to dictionary object
			ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'ContentType cont�m o tipo de conte�do de arquivo capturado da
			'string RequestBin
			'PosBeg � a posi��o inicial da pesquisa
			'PosEnd-PosBeg � a posi��o final da substring que conter� o tipo de
			'arquivo.
			
			UploadControl.Add "ContentType",ContentType
			'o nome do tipo de arquivo � colocado num dicion�rio de dados.
			'Get content of object
			PosBeg = PosEnd+4
			'PosBeg � a posi��o inicial da pesquisa que ir� verificar a extens�o
			'do arquivo.
			
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			'PosEnd � a posi��o final da pesquisa que ir� verificar a extens�o do
			'arquivo.
			
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			'A vari�vel Value recebe o tipo de extens�o do arquivo
			'que est� contido na string RequestBin.
			
			Else
			'Get content of object
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'Aqui os valores de extens�o de arquivo s�o capturados
			'quando se trata de um formul�rio comum.
		End If
		'Add content to dictionary object
	UploadControl.Add "Value" , Value	
		'a extens�o do arquivo � colocada num dicion�rio de dados.
	UploadRequest.Add name, UploadControl	
		'Loop para o pr�ximo formul�rio
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	'BoundaryPos � a posi��o da primeira ocorr�ncia da substring boundary
	'dentro da string RequestBin
	'BoundaryPos+LenB(boundary) � a posi��o inicial da pesquisa que ir� ser 
	'feita.
	'LenB() � uma fun��o que devolve o n�mero de letras de uma string.
	
	Loop
        
End Sub

'String to byte string conversion
Function getByteString(StringStr)
 For i = 1 to Len(StringStr)
 	char = Mid(StringStr,i,1)
	getByteString = getByteString & chrB(AscB(char))
 Next
End Function
'aqui temos uma fun��o de convers�o de bytes para string.
'Byte string to string conversion
Function getString(StringBin)
 getString =""
 For intCount = 1 to LenB(StringBin)
	getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
 Next
End Function
'aqui temos uma fun��o de convers�o de dados bin�rios para string.
%>