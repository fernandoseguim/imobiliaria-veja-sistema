
<%
' Author Philippe Collignon
' Email PhCollignon@email.com

Sub BuildUploadRequest(RequestBin)
    'RequestBin é a variável que contém todas as informaçòes
	'recebidas dos formulários de maneira binária.
	'Get the boundary
	
	PosBeg = 1
	'PosBeg é a posição inicial da string binária.
	
	
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	'PosBeg é a posição inicial da string a ser pesquisada
	'requestBin é a string a ser pesquisada
	'chr(13) é a substring a ser procurada na string RequestBin
	'PosEnd é a posição da primeira ocorrência da substring chr(13) na
	'string RequestBin
	
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	'boundary é uma substring de RequestBin 
	'RequestBin é a string principal
	'PosBeg é a posição inicial da substring a ser capturada de RequestBin
	' PosEnd é a posição final da substring a ser capturada de RequestBin
	
	boundaryPos = InstrB(1,RequestBin,boundary)
	'boundaryPos é a posição da primeira ocorrência da substring boundary
	'dentro da string RequestBin
	
	
	'Get all data inside the boundaries
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		'boundaryPos é a posição da primeira ocorrência da substring boundary
		'concatenada com "--" na string RequestBin
		
		
		
		'Members variable of objects are put in a dictionary object
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		'um objeto dicionário é criado para se colocar as informações
		'sobre a string RequestBin que foram dividas acima.
		'Get an object name
		
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		'Pos é a primeira ocorrência da substring "Content-Disposition" dentro da
		'string RequestBin
		'BoundaryPos é a posição inicial a partir da qual a substring"Content-Disposition
		'irá ser pesquisada
		
		
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		'Pos é a posição da primeira ocorrência da substring "name=" dentro 
		'de RequestBin
		'Pos que está entre parênteses é a posição inicial da pesquisa a ser
		'feita
		
		PosBeg = Pos+6
		'PosBeg é a posição inicial mais 6.
		
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		'PosEnd é a posição da primeira ocorrência da substring "chr(34)"
		'na string RequestBin
		'PosBeg é a posição inicial a partir da qual a pesquisa irá ser feita.
		
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		'Name é uma variável que receberá uma substring de RequestBin
		'que começa na posição PosBeg e termina na posição PosEnd-PosBeg
		
		
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		'PosFile é a posição da primeira ocorrência da substring "filename=" 
		'dentro da string RequestBin
		'BoundaryPos é a posição inicial a partir da qual a pesquisa irá ser feita.
		
		
		
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		'PosBound é a posição da primeira ocorrência da substring boundary dentro
		'da string RequestBin
		'PosEnd é a posição inicial a partir da qual a pesquisa irá ser feita.
		
		'Test if object is of file type
		If  PosFile<>0 AND (PosFile<PosBound) Then
		    'aqui será verificado se os dados dos formulários são de 
			'arquivos ou dados comuns.
			'Get Filename, content-type and content of file
			
			PosBeg = PosFile + 10
			'PosBeg é a posição inicial a partir da qual o nome do arquivo irá
			'ser capturado da string RequestBim
			'PosFile é a posição da primeira ocorrência de "filename=" dentro da
			'string RequestBin 
			
			
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			'PosEnd é a posição da primeira ocorrência da substring "chr(34)" dentro
			'da string RequestBin
			'PosBeg é a posição inicial a partir da qual a pesquisa irá ser feita.
			
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'FileName é o nome do arquivo que será capturado da string RequestBin
			'PosBeg é a posição inicial da pesquisa
			'PosEnd-PosBeg é a posição final da pesquisa
			'MidB é a função que captura a substring que contêm o nome do arquivo
			
			UploadControl.Add "FileName", FileName
			'o nome do arquivo é jogado num dicionário de dados.
			'
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			'Pos é a posição da primeira ocorrência da substring "Content-Type:" 
			'dentro da substring RequestBin
			'PosEnd é a posição incial da pesquisa que irá ser feita.
			
			PosBeg = Pos+14
			'PosBeg é a posição inicial da substring que irá ser capturada de
			'RequestBin que corresponderá ao tipo de conteúdo do arquivo.
			
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			'PosEnd é a posição da primeira ocorrência da substring "chr(13)" dentro
			' da string RequestBin
			
			
			'Add content-type to dictionary object
			ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'ContentType contêm o tipo de conteúdo de arquivo capturado da
			'string RequestBin
			'PosBeg é a posição inicial da pesquisa
			'PosEnd-PosBeg é a posição final da substring que conterá o tipo de
			'arquivo.
			
			UploadControl.Add "ContentType",ContentType
			'o nome do tipo de arquivo é colocado num dicionário de dados.
			'Get content of object
			PosBeg = PosEnd+4
			'PosBeg é a posição inicial da pesquisa que irá verificar a extensão
			'do arquivo.
			
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			'PosEnd é a posição final da pesquisa que irá verificar a extensão do
			'arquivo.
			
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			'A variável Value recebe o tipo de extensão do arquivo
			'que está contido na string RequestBin.
			
			Else
			'Get content of object
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			'Aqui os valores de extensão de arquivo são capturados
			'quando se trata de um formulário comum.
		End If
		'Add content to dictionary object
	UploadControl.Add "Value" , Value	
		'a extensão do arquivo é colocada num dicionário de dados.
	UploadRequest.Add name, UploadControl	
		'Loop para o próximo formulário
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	'BoundaryPos é a posição da primeira ocorrência da substring boundary
	'dentro da string RequestBin
	'BoundaryPos+LenB(boundary) é a posição inicial da pesquisa que irá ser 
	'feita.
	'LenB() é uma função que devolve o número de letras de uma string.
	
	Loop
        
End Sub

'String to byte string conversion
Function getByteString(StringStr)
 For i = 1 to Len(StringStr)
 	char = Mid(StringStr,i,1)
	getByteString = getByteString & chrB(AscB(char))
 Next
End Function
'aqui temos uma função de conversão de bytes para string.
'Byte string to string conversion
Function getString(StringBin)
 getString =""
 For intCount = 1 to LenB(StringBin)
	getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
 Next
End Function
'aqui temos uma função de conversão de dados binários para string.
%>