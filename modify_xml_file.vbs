'* -- modify_xml_file.vbs @StephaneAG - 2014 -- *'

'* -- script start -- *'

'* load the xml file to modify *'
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = "False"
xmlDoc.Load(".\app_config__barebone.xml")

'* modify the content of the xml file *'
xmlDoc.SelectSingleNode("/AppConfig/CheckProcOrWinRunning").Text = "false"
xmlDoc.SelectSingleNode("/AppConfig/CheckPositionAndSize").Text = "false"
xmlDoc.SelectSingleNode("/AppConfig/SetupSubWindowsForAutomation").Text = "true"
xmlDoc.SelectSingleNode("/AppConfig/ReInitBeforeActions").Text = "true"
xmlDoc.SelectSingleNode("/AppConfig/JobDoneNotification/DisplayNotification").Text = "false"
xmlDoc.SelectSingleNode("/AppConfig/JobDoneNotification/AutoCloseNotification").Text = "false"
xmlDoc.SelectSingleNode("/AppConfig/JobDoneNotification/DelayBeforeAutoClose").Text = 10
xmlDoc.SelectSingleNode("/AppConfig/FocusBackToApp").Text = "false"

'* save the modified xml file*'
xmlDoc.Save ".\app_config__modified.xml"

'* -- script end -- *'