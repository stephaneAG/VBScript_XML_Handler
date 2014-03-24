'* -- neatFramework__VBS__XML_module.vbs @StephaneAG - 2014 -- *'
'*
'*  The current file defines functions to be called externally
'*  thanks to the "Include" Sub defined in the "<calling_file>.vbs" file

'* -- helpers start -- *'
'* shortcuts for easier access *'
Set CliArgs = WScript.Arguments
'* -- helpers end -- *'


'* -- Functions & Subs definitions start -- *'

'* a "Hello World"-like fcn implm *'
Sub HelloBox_sub
  MsgBox "Hello World"
End Sub

'* display the 'no arguments received' error *'
Sub DisplayError_NoArgs_sub
  WScript.Echo "Error: No command line arguments were received !"
End Sub

'* display the help *'
Sub DisplayHelp_sub
  WScript.Echo "The program accept 2 methods: 'get' and 'set' " & vbcrlf & _
               "Examples:" & vbcrlf & vbcrlf & _
               "getting the value of the item present at the <Path> XPath in the <filename> XML file " & vbcrlf & _
               ".\handle_XML_file.vbs .\app_config__barebone.xml /AppConfig/CheckProcOrWinRunning get" & vbcrlf & vbcrlf & _
               "setting the value of the item present at the <Path> XPath in the <filename> XML file to 'true' " & vbcrlf & _
               ".\handle_XML_file.vbs .\app_config__barebone.xml /AppConfig/CheckProcOrWinRunning set true" & vbcrlf & vbcrlf
End Sub

'* FileExist_fcn(XMLfilePath) - check that the file actually exist *'
'* returns the path of the file if found, else display an error & quit *'
Function FileExist_fcn(XMLfilePath)
  Dim Fso
  Set Fso = CreateObject("Scripting.FileSystemObject")
  If ( Fso.FileExists( XMLfilePath ) ) Then
    FileExist_fcn = XMLfilePath
  Else
    WScript.Echo "Error: the file passed as cli parameter doesn't exists !"
    WScript.Quit()
  End If
End Function

'* IsXML_fcn(XMLfilePath) - check that the file actually has the '.xml' extension *'
'* returns the path of the file if it has the correct extension, else display an error & quit *'
Function IsXML_fcn(XMLfilePath)
  Dim Fso
  Set Fso = CreateObject("Scripting.FileSystemObject")
  If ( Fso.GetExtensionName( XMLfilePath ) = "xml" ) Then
    IsXML_fcn = XMLfilePath
  Else
    WScript.Echo "Error: the file passed as cli parameter doesn't have the '.xml' extension !"
    WScript.Quit()
  End If
End Function

'* IsNodePath_fcn(NodeXPath) - check that we have an actual path ( aka a string starting with a forward slash ( "/" )  ) *'
'* returns the XPath if it starts with a forward slash, else display an error & quit *'
Function IsNodePath_fcn(NodeXPath)
  If InStr(1, WScript.Arguments(1), "/") Then '* make sure that the XPath path received starts with a forward slash *'
    IsNodePath_fcn = NodeXPath
  Else
    WScript.Echo "The XPath path passed as cli parameter seems not valid" & vbcrlf  & _
                 " ( it should start with a forward slash ( '/' ) )"
    WScript.Quit()
  End If
End Function

'* NodePathExist_fcn(XMLfilePath, NodeXPath) - check that the XML node corresponding to the XPath received from the cli parameters actually exist *'
'* returns the XPath if a node is present at this location in the XML file, else display an error & quit *'
Function NodePathExist_fcn(XMLfilePath, NodeXPath)
  Dim xmlDoc
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.Async = "False"
  xmlDoc.Load( XMLfilePath )
  Dim requ_node
  Set requ_node = xmlDoc.SelectSingleNode( NodeXPath )
  If Not requ_node Is Nothing Then
    NodePathExist_fcn = NodeXPath
  Else
    WScript.Echo "The XPath path passed is not present in the XML file passed as cli parameter"
    WScript.Quit()
  End If
End Function

'* IsNumber_fcn(theValue) - check whether the value passed is a number or a string *'
'* returns either "Number" or "String" depending to the data contained in "theValue" *'
Function IsNumber_fcn(theValue)
  If IsNumeric(theValue) Then
    IsNumber_fcn = "Number"
  Else
    IsNumber_fcn = "String"
  End If
End Function

'*  - compare the two values passed and determine if their type differs *'
'* returns either "differs" or "match" *'
Function CompareTypes_fcn(firstValue, secondValue)
  If IsNumber_fcn(firstValue) = IsNumber_fcn(secondValue) Then
    CompareTypes_fcn = "matches"
  Else
    CompareTypes_fcn = "differs"
  End If
End Function

'* RequestType_fcn(requestArg) - return the request type depending on the argument passed *'
'* will either return "get", "set", or "unsupported" *'
Function RequestType_fcn(requestArg)
If requestArg = "get" Then
  RequestType_fcn = "get"
ElseIf requestArg = "set" Then
  RequestType_fcn = "set"
ElseIf requestArg = "tree" Then
  RequestType_fcn = "tree"
Else
  RequestType_fcn = "unsupported"
End If
End Function

'* HandleRequestGET_fcn(XMLfilePath, NodeXPath) - handle "get" requests *'
'* returns the value requested, extracted from the specified XML file at the desired XPath *'
Function HandleRequestGET_fcn(XMLfilePath, NodeXPath)
  Dim xmlDoc
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.Async = "False"
  xmlDoc.Load(XMLfilePath)
  Dim requ_node
  Set requ_node = xmlDoc.SelectSingleNode(NodeXPath)
  If Not requ_node Is Nothing Then
    HandleRequestGET_fcn = requ_node.Text
  Else
    WScript.Echo "An error happened while retrieving the value of the specified XPath XML node"
    WScript.Quit()
  End If
End Function
'* handle "set" requests *'
Function HandleRequestSET_fcn(XMLfilePath, NodeXPath, NewValue)
  Dim xmlDoc
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.Async = "False"
  xmlDoc.Load( XMLfilePath )
  Dim requ_node
  Set requ_node = xmlDoc.SelectSingleNode( NodeXPath )
  '* check if the types of the current value of the XPath XML node and the new value are the same *'
  CurrentNodeValue = requ_node.Text
  If CompareTypes_fcn(CurrentNodeValue, NewValue) = "differs" Then
    '* TO THNK ABOUT: CHECK FOR A FIFTH CLI ARG TO "FORCE" A NEW VALUE WITH A DIFFERENT TYPE (..) *'
    WScript.Echo "The types of the current XPath XML node value and the new one are different: operation aborted"
    WScript.Quit()
  End If
  '* udpate the value *'
  requ_node.Text = newValue
  '* save the file *'
  xmlDoc.Save XMLfilePath
  '* return the new value of XML XPath node after successful update *'
  HandleRequestSET_fcn = requ_node.Text
End Function

'* display a structured tree of the XML file *'
Function DisplayTree_fcn(XMLfilePath, NodeXPath)
  Dim xmlDoc
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.Async = "False"
  xmlDoc.Load( XMLfilePath )
  Dim root_node
  Set root_node = xmlDoc.SelectNodes(NodeXPath)
End Function
'* -- Functions & Subs definitions end -- *'