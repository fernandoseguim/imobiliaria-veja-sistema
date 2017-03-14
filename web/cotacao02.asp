<% 
'criamos o nome do arquivo 
arquivo= request.serverVariables("APPL_PHYSICAL_PATH") & "provas.txt" 

'conectamos com o FSO 
set confile = createObject("scripting.filesystemobject") 

'criamos o objeto TextStream 
set fich = confile.CreateTextFile(arquivo) 

'escrevemos os números do 0 ao 9 
for i=0 to 9 
   fich.write(i) 
next 

'fechamos o arquivo 
fich.close() 

'voltamos a abrir o arquivo para leitura 
set fich = confile.OpenTextFile(arquivo) 

'lemos o conteúdo do arquivo 
texto_arquivo = fich.readAll() 

'imprimimos na página o conteúdo do arquivo 
response.write(texto_arquivo) 

'fechamos o arquivo 
fich.close() 
%> 

