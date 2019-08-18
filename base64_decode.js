//This script takes base64 encoded form, writes it to a file and decodes it using below function.

var fso = new ActiveXObject("Scripting.FileSystemObject");
var b = fso.CreateTextFile("PathToWriteBase64", true);
b.WriteLine("base64encodedhere")'
b.Close();

foForReading          = 1 // Open base 64 code file for reading
foAsASCII             = 0 // Open base 64 code file as ASCII file
adSaveCreateOverWrite = 2 // Mode for ADODB.Stream
adTypeBinary          = 1 // Binary file is encoded

function decode(from, to) {
	
  from = "PathToWriteBase64"
	to = "PathToDecodedBinary"
	
    var fileSystemObj = WScript.CreateObject("Scripting.FileSystemObject");
    var file = fileSystemObj.GetFile(from)
    var inputStream = file.OpenAsTextStream(foForReading, foAsASCII);
    
    var xmlObj = WScript.CreateObject("MSXml2.DOMDocument");
    var docElement = xmlObj.createElement("Base64Data");
    docElement.dataType = "bin.base64";
    
    docElement.text = inputStream.ReadAll();
    
    var outputStream = WScript.CreateObject("ADODB.Stream");
    outputStream.Type = adTypeBinary;
    outputStream.Open();
    
    outputStream.Write(docElement.nodeTypedValue);
    outputStream.SaveToFile(to, adSaveCreateOverWrite);
    
    inputStream.Close();
    outputStream.Close();
}

function getDataName(txtName) {
    return txtName.substring(0, txtName.length - ".txt".length);
}

function getFileNames() {
    var fileSystem = WScript.CreateObject("Scripting.FileSystemObject");
    var files = new Enumerator(fileSystem.GetFolder(".").files);
    
    var result = [];
    for (; !files.atEnd(); files.moveNext()) {
        var file = files.item();
        if(file.Name.match(/\.txt/)) {
				result.push(file.Name);				
				
        }
    }
    
    return result;

}

function main() {
    var files = getFileNames();
    for (i in files) {
        decode(files[i], getDataName(files[i]));
    }
}

main();

