Dim XMLRequest
Dim AppUser, folder, datafiles, fullpath

'-----The appuser will require security to run the syutlf process, please udpate XXXXXX
'-----The path will need to be fully qualified and accessable by the workflow engine ( same box is good )
'-----File name does not require extension

AppUser = "XXXXXX"
folder = "\\fullydefinedpath to file\"  
datafiles = "filename"
fullpath = """" & folder & datafiles & """"

Dim BrokerPath, AppEnvironment, btDBData

Set btDBData = CreateObject("BTDBData.BTDBData")
BrokerPath = btDBData.WWWServer & "isapi/btwebrqb.dll"
AppEnvironment = btDBData.ConnectionName

XMLRequest = XMLRequest & "<sbixml>"
XMLRequest = XMLRequest & "<NetSightMessage>"
XMLRequest = XMLRequest & "<Header>"
XMLRequest = XMLRequest & "<Connection>" & AppEnvironment & "</Connection>"
XMLRequest = XMLRequest & "<UserID>" & AppUser & "</UserID>"
XMLRequest = XMLRequest & "<CurrentLedgers GL=""SO"" JL=""--"" />"
XMLRequest = XMLRequest & "<SubSystem></SubSystem>"
XMLRequest = XMLRequest & "<Timeout>90000</Timeout>"
XMLRequest = XMLRequest & "</Header>"
XMLRequest = XMLRequest & "<Request Type=""WorkflowRun"" >"
XMLRequest = XMLRequest & "<WorkflowRun>"
XMLRequest = XMLRequest & "<Model>JOB</Model>"
XMLRequest = XMLRequest & "<Mask>SYUTLF</Mask>"
XMLRequest = XMLRequest & "<Assembly>BT70SY</Assembly>"
XMLRequest = XMLRequest & "<Class>DelimitedFileLoader</Class>"
XMLRequest = XMLRequest & "<Description>Load Delimited Data File</Description>"
XMLRequest = XMLRequest & "<Prompts>"
XMLRequest = XMLRequest & "<Prompt Id=""Mask""       Response=""SYUTLF""/>"

'-----Please replace XX with your GLLedger value------

XMLRequest = XMLRequest & "<Prompt Id=""GLLedger""   Response=""XX""/>"
XMLRequest = XMLRequest & "<Prompt Id=""JLLedger""   Response=""--""/>"                
XMLRequest = XMLRequest & "<Prompt Id=""WFIF""       Response="  & fullpath &  "  ShowInTailsheet=""True""/>"
XMLRequest = XMLRequest & "<Prompt Id=""NUN2""        Response=""NO""   ShowInTailsheet=""True""/>"
XMLRequest = XMLRequest & "<Prompt Id=""GLIZ""        Response=""NO""   ShowInTailsheet=""True""/>"
XMLRequest = XMLRequest & "<Prompt Id=""GLKM""        Response=""NO""   ShowInTailsheet=""True""/>"
XMLRequest = XMLRequest & "<Prompt Id=""GL4F""        Response=""NO""   ShowInTailsheet=""True""/>"
XMLRequest = XMLRequest & "</Prompts>"
XMLRequest = XMLRequest & "</WorkflowRun>"
XMLRequest = XMLRequest & "</Request>"
XMLRequest = XMLRequest & "</NetSightMessage>"
XMLRequest = XMLRequest & "</sbixml>"


Call XML_Request(XMLRequest)

Function XML_Request(XML)
	Dim DIAG
	Dim result, subject, message
	Dim oSendDoc, oServerHttp
	DIAG = "N"
	Set oSendDoc = CreateObject("MSXML2.DOMDocument")
	oSendDoc.async = False
	result = oSendDoc.loadXML(XML) 
	DIAG = "Failed to create object MSXML2.DOMDocument"
	If IsObject(oSendDoc) Then
		Set oServerHttp = CreateObject("MSXML2.SERVERXMLHTTP")
		DIAG = "Failed to create object MSXML2.SERVERXMLHTTP"
		If IsObject(oServerHttp) Then
			oServerHttp.open"POST",BrokerPath,0
			oServerHttp.setOption(2) = 4096
			oServerHttp.send(oSendDoc)
			DIAG = "HTTP Status = " & oServerHttp.status  & " - " & oServerHttp.statusText
			 If Int(oServerHttp.status) = 200 Then
				Dim XMLResponse
				Set XMLResponse = oServerHttp.responseXML
				DIAG = "No XML Response"
				If Not XMLResponse Is Nothing Then
					Dim JobInfo
					Set JobInfo = XMLResponse.selectSingleNode("//Response/JobNumber")
					DIAG = "No Job Info"
					If Not JobInfo Is Nothing Then
						Dim JobNumber
                        '------The variable needs to be defined in the workflow designer, then uncomment
						'------The msgbox (as text) is not allowed in script, other than debugging, so m is removed		
						'Variables.JobNo = JobInfo.Text
						'sgbox  JobInfo.getAttribute("IfasJobno")
						DIAG = "Y"
					End If
				End If                  
			End If
		End If
	End If
	If DIAG = "Y" Then
	Else
		'Variables.JobNo = "N"
		'sgbox DIAG
	End If
End Function
