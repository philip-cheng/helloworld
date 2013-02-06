'cscript //nologo FindFailureTicket.vbs (do not runas admin)
'[inbox -> A_PEACY -> 1_Msg -> 0_SolutionGroup_CENTRAL iLAB AP]
Const olFolderInbox = 6
Const olMailItem = 0
Const G_CNS_ForReading = 1
Const G_CNS_ForWriting = 2
Const G_CNS_ForAppending = 8

'0: no debug, 1: main info, 9: details
Const CNS_DEBUG_MODE = 1
Dim g_objOutlook, g_objNamespace, g_objInbox
Dim g_dicFailureItem
Dim strStartTime, strEndTime
strStartTime = Time()
Call PrintBeginEnd("FindFailureTicket.vbs", 0)
Call FindFailureTicket_Main()
Call PrintBeginEnd("FindFailureTicket.vbs", 1)
strEndTime = Time()
WScript.Echo "[FindFailureTicket.vbs] Duration: " & DateDiff("s", strStartTime, strEndTime) & " seconds."
WScript.Quit

'Function Name: FindFailureTicket_Main
'Description: Main procedure for this script
'Param: none
'Return: none
Sub FindFailureTicket_Main()

  Dim strStartTime, strEndTime
	strStartTime = Time()

	Call Init()

	If CNS_DEBUG_MODE = 9 Then
		WScript.Echo String(50,"-")
		WScript.Echo "Inbox name[" & g_objInbox.name & "]"
		'WScript.Echo "Inbox Sub folders count[" & g_objInbox.Folders.count & "]"
		'WScript.Echo "Inbox items count[" & g_objInbox.items.count & "]"
		WScript.Echo String(50,"-")
	End If

	Call GetFailureTicket()

	Call Export2Csv(g_dicFailureItem)

	strEndTime = Time()
	WScript.Echo "[FindFailureTicket_Main] Duration: " & DateDiff("s", strStartTime, strEndTime) & " seconds."

End Sub

'Function Name: GetFailureTicket
'Description: Get failure mail from "0_SolutionGroup_CENTRAL iLAB AP"
'Param: none
'Return: none
Private Sub GetFailureTicket()

	Dim objTarFolder
	Dim strSubject, strTicketNr, strSummary, strBody, strFrom, strDept, strSite, strDate, strFailureInfo
	Dim strStartTime, strEndTime

	strStartTime = Time()

	If CNS_DEBUG_MODE = 1 Then
		Call PrintBeginEnd("GetFailureTicket", 0)
	End If

	'inbox - A_PEACY - 1_Msg - 0_SolutionGroup_CENTRAL iLAB AP
	Set objMailItem = g_objOutlook.CreateItem(olMailItem)
	Set objTarFolder = g_objInbox.Folders.Item("A_PEACY").Folders.Item("1_Msg").Folders.Item("0_SolutionGroup_CENTRAL iLAB AP")
	If CNS_DEBUG_MODE = 9 Then
		WScript.Echo "objTarFolder name[" & objTarFolder.Name & "]"
		WScript.Echo "objTarFolder item count[" & objTarFolder.Items.Count & "]"
		WScript.Echo "objTarFolder.FolderPath[" & objTarFolder.FolderPath & "]"
	End If

	WScript.Echo String(50,"-")

	For Each itm In objTarFolder.Items
		Set objMailItem = itm
		strBody = objMailItem.Body
		If InStr(strBody, "Failure") > 0 Then
			strSubject = Trim(objMailItem.Subject)
			strTicketNr = Mid(strSubject, InStr(strSubject, "Incident") + 9, 15)
			strSummary = Right(strSubject, Len(strSubject) - InStr(strSubject, "Summary") - 8)
			strFrom = Trim(Split(Split(strBody, ":")(1), vbCrLf)(0))
			strDept = Trim(Split(Split(strBody, ":")(2), vbCrLf)(0))
			strSite = Trim(Split(Split(strBody, ":")(3), vbCrLf)(0))
			strDate = objMailItem.ReceivedTime
			strFailureInfo = strDate & "," & strFrom & "," & strDept & "," & strSite & "," & strSummary
			If Not g_dicFailureItem.Exists(strTicketNr) Then
				g_dicFailureItem.Add strTicketNr, strFailureInfo
			Else
				WScript.Echo "Duplicated[" & strTicketNr & "]"
			End If
			If CNS_DEBUG_MODE = 9 Then
				'WScript.Echo "Subject[" & strSubject & "]"
				'WScript.Echo "Body[" & strBody & "]"
				WScript.Echo "Nr[" & strTicketNr & "]"
				WScript.Echo "strDate[" & strDate & "]"
				WScript.Echo "strFrom[" & strFrom & "]"
				WScript.Echo "strDept[" & strDept & "]"
				WScript.Echo "strSite[" & strSite & "]"
				WScript.Echo "Summary[" & strSummary & "]"
				'WScript.Echo "LastModificationTime[" & objMailItem.LastModificationTime & "]"
				'WScript.Echo itm.Class
			End If
			'WScript.Quit
		End If
	Next
	WScript.Echo String(15,"-") & "Total failure count[" & g_dicFailureItem.Count & "]" & String(15,"-")
	WScript.Echo "[Ticket Number]=[Date,From,Dept,Site,Summary]"
	For Each key In g_dicFailureItem.Keys
		If Not IsEmpty(key) Then
			WScript.Echo "[" & key & "]=[" & g_dicFailureItem.Item(key) & "]"
		End If
	Next
	WScript.Echo String(50,"-")

	If CNS_DEBUG_MODE = 1 Then
		Call PrintBeginEnd("GetFailureTicket", 1)
	End If

	strEndTime = Time()
	WScript.Echo "[GetFailureTicket] Duration: " & DateDiff("s", strStartTime, strEndTime) & " seconds."
End Sub

'Function Name: Init
'Description: Initilization
'Param: none
'Return: none
Private Sub Init()

	Dim strStartTime, strEndTime

	strStartTime = Time()

	If CNS_DEBUG_MODE = 9 Then
		Call PrintBeginEnd("Init", 0)
	End If

	'Initialize for global variables
	Set g_objOutlook = CreateObject("Outlook.Application")
	Set g_objNamespace = g_objOutlook.GetNamespace("MAPI")
	Set g_objInbox = g_objNamespace.GetDefaultFolder(olFolderInbox)
	Set g_dicFailureItem = CreateObject("Scripting.Dictionary")

	If CNS_DEBUG_MODE = 9 Then
		Call PrintBeginEnd("Init", 1)
	End If

	strEndTime = Time()
	WScript.Echo "[Init] Duration: " & DateDiff("s", strStartTime, strEndTime) & " seconds."
End Sub

'Function Name: Export2Csv
'Description: Export data to csv file for excel use
'Param: g_dicFailureItem -> Name of Sub/func
'       p_strBeginEnd -> Indicate start(0) or end(1)
'Return: none
Private Sub Export2Csv(ByVal p_dicFailureItem)
	Dim objFso, objCsvFile
	Dim strFilename, strLine
	Dim key
	Dim strStartTime, strEndTime

	strStartTime = Time()

	strFilename = "FailureTicketInfo_" & Date() & ".log"
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objCsvFile = objFso.CreateTextFile(strFilename, G_CNS_ForWriting)
	
	For Each key In p_dicFailureItem.Keys
		strLine = key & "," & p_dicFailureItem(key)
		objCsvFile.WriteLine strLine
	Next
	objCsvFile.Close

	strEndTime = Time()
	WScript.Echo "[Export2Csv] Duration: " & DateDiff("s", strStartTime, strEndTime) & " seconds."
End Sub

'Function Name: PrintBeginEnd
'Description: Print info to show a Sub/func starts
'Param: p_strProcName -> Name of Sub/func
'       p_strBeginEnd -> Indicate start(0) or end(1)
'Return: none
Private Sub PrintBeginEnd(ByVal p_strProcName, ByVal p_intBeginEnd)
	Dim strSuffix, strMessage, strSpace
	Dim strDateTime

	strDateTime = " (" & Date() & " " & Time() & ")"

	If p_intBeginEnd = 0 Then
		strSuffix = " started." & strDateTime
	Else
		strSuffix = " ended." & strDateTime
	End If

	strMessage = "' " & p_strProcName & strSuffix

	If Len(strMessage) < 50 Then
		strSpace = Space(50 - Len(strMessage) - 1)
		strMessage = strMessage & strSpace & "'"
	End If

	WScript.Echo String(50,"'")
	WScript.Echo strMessage
	WScript.Echo String(50,"'")
End Sub
