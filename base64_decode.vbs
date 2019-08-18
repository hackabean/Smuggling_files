Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="PathToWriteBase64To"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write "base64encodedhere" & vbCrLf
objFile.Write "objFile.Write"
objFile.Close



Call Base64Decode((srcFileName), (dstFileName))

	srcFileName = "PathToWriteBase64To"
	dstFileName = "Base64DecodedFile"

'BASE64
Function Base64Decode(srcFileName, dstFileName)

	set fs = wscript.CreateObject("Scripting.FileSystemObject")

	Dim fileSystemObj,file
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
	Set file = fileSystemObj.OpenTextFile(srcFileName, 1)
	txtVar = file.ReadAll() 
	
    Set xmlDom = CreateObject("Microsoft.XMLDOM")
    Set xmlNode = xmlDom.createElement("MyNode")
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = Replace(txtVar, vbCrLf, "")
    
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 1  'adTypeBinary
    adoStream.Open()
    adoStream.Write(xmlNode.nodeTypedValue)
    adoStream.Position = 0
	Call adoStream.SaveToFile (dstFileName, 2)
    adoStream.Close()
	

End Function

