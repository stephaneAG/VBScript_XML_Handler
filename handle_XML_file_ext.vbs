'* -- handle_XML_file_ext.vbs @StephaneAG - 2014 -- *'
'*
'*  2nd implementation of the "handle_XML_file.vbs",  
'*  making use of functions é subs defined externally
'*

'* -- helpers start -- *'

'* - the "Include" Sub that makes this possible - *'
Sub Include_sub(module_file)
  Dim Fso, F
  Set Fso = CreateObject("Scripting.FileSystemObject")
  Set F = Fso.OpenTextFile(module_file & ".vbs", 1)
  Str = F.ReadAll
  F.Close
  ExecuteGlobal Str
End Sub

'* -- helpers end -- *'

'* -- includes start -- *'
'* make use of one of our functions defined in an external .vbs file *'
Include_sub "neatFramework__VBS__XML_module"
'* -- includes end -- *'

'* -- script start -- *'

'* call our test Sub defined in the included module file *'
HelloBox_sub

'* handle the cli arguments passed *'
If Not CliArgs.Count = 0 Then 
  If CliArgs.Count = 1 And CliArgs.Item(0) = "help" Then
    DisplayHelp_sub '* display the help *'
  ElseIf CliArgs.Count = 3 Or CliArgs.Count = 4 Then
    
    '* do the checks (..) *'
    XMLfilePath = CliArgs.Item(0)
    XMLfilePath = FileExist_fcn(XMLfilePath)
    XMLfilePath = IsXML_fcn(XMLfilePath)
    WScript.Echo "processed existing .xml: " & XMLfilePath

    NodeXPath = CliArgs.Item(1)
    NodeXPath = IsNodePath_fcn(NodeXPath)
    NodeXPath = NodePathExist_fcn(XMLfilePath, NodeXPath)
    WScript.Echo "processed existing XML XPath: " & NodeXPath

    '* check the request type & act accordingly *'
    RequestType = CliArgs.Item(2)
    RequestType = RequestType_fcn(RequestType)
    WScript.Echo "processed request type: " & RequestType
    
    If RequestType = "get" Then
      RequestResult = HandleRequestGET_fcn(XMLfilePath, NodeXPath)
      WScript.Echo "Request result: " & RequestResult
      'WScript.Echo "Request result: " & RequestResult
    ElseIf RequestType = "set" Then
      NewValue = CliArgs.Item(3)
      RequestResult = HandleRequestSET_fcn(XMLfilePath, NodeXPath, NewValue)
      WScript.Echo "Request result: " & RequestResult
    ElseIf RequestType ="tree" Then
      TreRetVal = DisplayTree_fcn(XMLfilePath, NodeXPath)
    ElseIf RequestType ="unsupported" Then
      
    End If

  Else
    DisplayHelp_sub '* display the help *'
  End If
Else
  DisplayError_NoArgs_sub
  DisplayHelp_sub '* display the help *'
End If
'* -- script end -- *'