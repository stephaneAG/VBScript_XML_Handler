'* -- read_xml_file.vbs @StephaneAG - 2014 -- *'

'* -- script start -- *'

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
xmlDoc.Load(".\app_config__barebone.xml")

'* build our first "query" to the XML file *'
Set stuffNodes = xmlDoc.SelectNodes("/AppConfig/*") '* we make use of the wildcard to get all the properties ( aka "child nodes" ) of the "AppConfig" node *'
For Each stuffNode in stuffNodes
  WScript.Echo stuffNode.Text
Next

'* build a "query" that returns a specific property *'
Set cpowrNode = xmlDoc.SelectSingleNode("/AppConfig/CheckProcOrWinRunning")
WScript.Echo "cpowr value: " & cpowrNode.Text

'* build a "query" that returns more than one property *'
Set someProps = xmlDoc.SelectNodes("/AppConfig/(CheckProcOrWinRunning|CheckPositionAndSize)")
WScript.Echo "cpowr & cpas values: " & someProps(0).Text & someProps(1).Text 
For Each someProp in someProps
  WScript.Echo someProp.NodeName & someProp.Text
Next

'* build a complete overview of the current state of the XML file's content *'
MsgBox "AppConfig from xmlDoc:"

Set cpowr_node = xmlDoc.SelectSingleNode("/AppConfig/CheckProcOrWinRunning")
cpowr_name = cpowr_node.NodeName
cpowr_value = cpowr_node.Text

Set cpas_node = xmlDoc.SelectSingleNode("/AppConfig/CheckPositionAndSize")
cpas_name = cpas_node.NodeName
cpas_value = cpas_node.Text

Set ssfa_node = xmlDoc.SelectSingleNode("/AppConfig/SetupSubWindowsForAutomation")
ssfa_name = ssfa_node.NodeName
ssfa_value = ssfa_node.Text

Set riba_node = xmlDoc.SelectSingleNode("/AppConfig/ReInitBeforeActions")
riba_name = riba_node.NodeName
riba_value = riba_node.Text

Set jdn_dn_node = xmlDoc.SelectSingleNode("/AppConfig/JobDoneNotification/DisplayNotification")
jdn_dn_name = jdn_dn_node.NodeName
jdn_dn_value = jdn_dn_node.Text

Set jdn_acn_node = xmlDoc.SelectSingleNode("/AppConfig/JobDoneNotification/AutoCloseNotification")
jdn_acn_name = jdn_acn_node.NodeName
jdn_acn_value = jdn_acn_node.Text

Set jdn_dbac_node = xmlDoc.SelectSingleNode("/AppConfig/JobDoneNotification/DelayBeforeAutoClose")
jdn_dbac_name = jdn_dbac_node.NodeName
jdn_dbac_value = jdn_dbac_node.Text

Set fbta_node = xmlDoc.SelectSingleNode("/AppConfig/FocusBackToApp")
fbta_name = fbta_node.NodeName
fbta_value = fbta_node.Text

'* display the overview using a popup box ( easier for debug implm ) *'
MsgBox "    ------ App config start ------    " & vbcrlf & vbcrlf &  _
       cpowr_name & ": " & cpowr_value & vbcrlf & _
       cpas_name & ": " & cpas_value & vbcrlf & _
       ssfa_name & ": " & ssfa_value & vbcrlf & _
       riba_name & ": " & riba_value & vbcrlf & _
       "JobDoneNotification: " & vbcrlf & _
       "  " & jdn_dn_name & ": " & jdn_dn_value & vbcrlf & _
       "  " & jdn_acn_name & ": " & jdn_acn_value & vbcrlf & _
       "  " & jdn_dbac_name & ": " & jdn_dbac_value & vbcrlf & _
       fbta_name & ": " & fbta_value & vbcrlf & vbcrlf & _ 
       "    ------- App config end -------    ", , _ 
       "AppConfig from xmlDoc:"

'* -- script end -- *'