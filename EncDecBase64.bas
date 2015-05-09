Attribute VB_Name = "EncDecBase64"
Option Explicit

Public Sub EncodeDecodeBase64()
    Debug.Print vbLf
    Debug.Print DateTime.Now
    
    Dim textToEnc As String, b64string As String, TextDecoded As String
    
    textToEnc = "username:P@s$w0Rd"

    Debug.Print TypeName(textToEnc); Tab(10); textToEnc
     
    b64string = EncodeBase64(textToEnc)
    
    Debug.Print TypeName(b64string); Tab(10); b64string
    
    TextDecoded = DecodeBase64(b64string)
    
    Debug.Print TypeName(TextDecoded); Tab(10); TextDecoded

End Sub

Function EncodeBase64(textStr As String) As String

  Dim arrData() As Byte
  arrData = StrConv(textStr, vbFromUnicode)

  'Notes re: StrConv:
    'In Office 2013, the StrConv function returns a Variant (String) converted as specified.
    'The vbFromUnicode conversion argument converts the string from Unicode to the default code page of the system. (Not available on the Macintosh.)
    'Constant: vbFromUnicode
    'Value: 128

  'Add "Microsoft XML, v6.0" reference (Tools menu) to avoid "user-defined type not defined" error
  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  'help from MSXML
  Set objXML = New MSXML2.DOMDocument

  'byte array to base64
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  'This function inserts a line feed every 72 characters
  'Removing the line feeds to avoid problems with the encoded string
  EncodeBase64 = Replace(objNode.text, vbLf, "")
  'Original return statement
  'EncodeBase64 = objNode.text

  'thanks, bye
  Set objNode = Nothing
  Set objXML = Nothing

End Function

Function DecodeBase64(encStr As String) As String

    Dim arrData() As Byte
    Dim strData As String
        
    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    'help from MSXML
    Set objXML = New MSXML2.DOMDocument
    
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = encStr
    arrData = objNode.nodeTypedValue
    strData = StrConv(arrData, vbUnicode)
    
        'The vbUnicode conversion argument converts the string to Unicode using the default code page of the system. (Not available on the Macintosh.)
        'Constant: vbUnicode
        'Value: 64
    
    DecodeBase64 = strData
    
    'thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

End Function

