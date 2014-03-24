'* -- create_xml_file.vbs @StephaneAG - 2014 -- *'

'* -- script start -- *'

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set objRoot = xmlDoc.CreateElement("AppConfig")
xmlDoc.AppendChild objRoot

Set obj_cpowr = xmlDoc.CreateElement("CheckProcOrWinRunning")
obj_cpowr.Text = "true"
objRoot.AppendChild obj_cpowr

Set obj_cpas = xmlDoc.CreateElement("CheckPositionAndSize")
obj_cpas.Text = "true"
objRoot.AppendChild obj_cpas

Set obj_ssfa = xmlDoc.CreateElement("SetupSubWindowsForAutomation")
obj_ssfa.Text = "false"
objRoot.AppendChild obj_ssfa

Set obj_riba = xmlDoc.CreateElement("ReInitBeforeActions")
obj_riba.Text = "false"
objRoot.AppendChild obj_riba

Set obj_jdn = xmlDoc.CreateElement("JobDoneNotification")
objRoot.AppendChild obj_jdn

Set obj_jdn_dn = xmlDoc.CreateElement("DisplayNotification")
obj_jdn_dn.Text = "true"
obj_jdn.AppendChild obj_jdn_dn

Set obj_jdn_acn = xmlDoc.CreateElement("AutoCloseNotification")
obj_jdn_acn.Text = "true"
obj_jdn.AppendChild obj_jdn_acn

Set obj_jdn_dbac = xmlDoc.CreateElement("DelayBeforeAutoClose")
obj_jdn_dbac.Text = 20
obj_jdn.AppendChild obj_jdn_dbac

Set obj_fbta = xmlDoc.CreateElement("FocusBackToApp")
obj_fbta.Text = "true"
objRoot.AppendChild obj_fbta

'* -- simplest way of writing an XML file ,but indentation is not available whatever the program opening the XML file -- *'
'Set objIntro = xmlDoc.CreateProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
'xmlDoc.InsertBefore objIntro, xmlDoc.ChildNodes(0)
'xmlDoc.Save ".\app_config__generated.xml"

'* -- writing an XML file while indenting it for easier use whatever the program used to view or edit the file -- *'
Set rdr = CreateObject("MSXML2.SAXXMLReader")
Set wrt = CreateObject("MSXML2.MXXMLWriter")
Set oStream = CreateObject("ADODB.STREAM")
oStream.Open
oStream.Charset = "UTF-8"
 
wrt.Indent = True
wrt.Encoding = "UTF-8"
wrt.Output = oStream
Set rdr.ContentHandler = wrt
Set rdr.ErrorHandler = wrt
rdr.Parse xmlDoc
wrt.Flush
 
oStream.SaveToFile "app_config__generated.xml", 2
 
Set rdr = Nothing
Set wrt = Nothing


'* -- script end -- *'