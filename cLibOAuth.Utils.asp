<%
'******************************************************************************
'	CLASS:		cLibOAuthUtils
'	PURPOSE:	
'
'	AUTHOR:	sdesapio		DATE: 04.04.10			LAST MODIFIED: 04.04.10
'******************************************************************************
	Class cLibOAuthUtils
	'**************************************************************************
'***'PRIVATE CLASS MEMBERS
	'**************************************************************************
		Private m_intOffsetMinutes
		Private m_intTimeStamp
		Private	m_strNonce

	'**************************************************************************
'***'CLASS_INITIALIZE / CLASS_TERMINATE
	'**************************************************************************
		Private Sub Class_Initialize()
			m_intOffsetMinutes = Null
			m_intTimeStamp = Null
			m_strNonce = Null
		End Sub
		Private Sub Class_Terminate()
		End Sub

	'**************************************************************************
'***'PUBLIC PROPERTIES
	'**************************************************************************
		Public Property Get OffsetMinutes
			If IsNull(m_intOffsetMinutes) Then
				Set_OffsetMinutes()
			End If
			
			OffsetMinutes = m_intOffsetMinutes
		End Property

		Public Property Get Nonce
			If IsNull(m_strNonce) Then
				Set_Nonce()
			End If
			
			Nonce = m_strNonce
		End Property

		Public Property Get TimeStamp
			If IsNull(m_intTimeStamp) Then
				Set_Timestamp()
			End If
			
			TimeStamp = m_intTimeStamp
		End Property

	'**************************************************************************
'***'PUBLIC FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	FUNCTION:	Get_ResponseValue
	'	PARAMETERS:	strResponseText, strKey
	'	PURPOSE:	Rips out a specific key/value pair from the service 
	'				provider response and returns the value
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Function Get_ResponseValue(strResponseText, strKey)
			Dim arrPair, arrPairs : arrPairs = Split(strResponseText, "&")
			Dim strRetVal : strRetVal = Null

			Dim i : i = 0 : Do While i < UBound(arrPairs) + 1
				arrPair = arrPairs(i)
				arrPair = Split(arrPair, "=")

				If arrPair(0) = strKey Then
					strRetVal = arrPair(1)
					Exit Do
				End If
				
				i = i + 1
			Loop

			Get_ResponseValue = strRetVal
		End Function

	'**************************************************************************
	'	SUB:			SortDictionary
	'	PARAMETERS:		objDict (collection), intSort (type)
	'	PURPOSE:		Sorts a dictionary on key or item. 
	'	REF:			http://support.microsoft.com/kb/246067
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Sub SortDictionary(objDict, intSort)
			Const dictKey  = 1
			Const dictItem = 2

			' declare our variables
			Dim strDict()
			Dim objKey
			Dim strKey,strItem
			Dim X,Y,Z

			' get the dictionary count
			Z = objDict.Count

			' we need more than one item to warrant sorting
			If Z > 1 Then
				' create an array to store dictionary information
				ReDim strDict(Z,2)
				X = 0
				' populate the string array
				For Each objKey In objDict
					strDict(X,dictKey)  = CStr(objKey)
					strDict(X,dictItem) = CStr(objDict(objKey))
					X = X + 1
				Next

				' perform a a shell sort of the string array
				For X = 0 to (Z - 2)
				For Y = X to (Z - 1)
					If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
						strKey  = strDict(X,dictKey)
						strItem = strDict(X,dictItem)
						strDict(X,dictKey)  = strDict(Y,dictKey)
						strDict(X,dictItem) = strDict(Y,dictItem)
						strDict(Y,dictKey)  = strKey
						strDict(Y,dictItem) = strItem
					End If
				Next
				Next

				' erase the contents of the dictionary object
				objDict.RemoveAll

				' repopulate the dictionary with the sorted information
				For X = 0 to (Z - 1)
					objDict.Add strDict(X,dictKey), strDict(X,dictItem)
				Next

			End If
		End Sub

	'**************************************************************************
	'	FUNCTION:		URLEncode
	'	PARAMETERS:		s (string)
	'	PURPOSE:		URL Encodes only those characters required by the oAuth 
	'					standard because native Server.URLEncode encodes too 
	'					much causing call to fail.
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Function URLEncode(s)
			Dim strTmpVal : strTmpVal = s
			Dim strRetVal : strRetVal = ""
			Dim intAsc : intAsc = 0
			Dim strHex : strHex = ""

			Dim i, strChr : For i = 1 To Len(strTmpVal)
				strChr = Mid(strTmpVal, i, 1)
				
				If InStr(1, OAUTH_UNRESERVED, strChr) = 0 Then
					intAsc = Asc(strChr)
					
					If intAsc < 32 Or intAsc > 126 Then
						strHex = encodeURIComponent(strChr)
					Else
						strHex = "%" & Hex(intAsc)
					End If

					strRetVal = strRetVal & strHex
				Else
					strRetVal = strRetVal & strChr
				End If
			Next

			URLEncode = strRetVal
		End Function

	'**************************************************************************
'***'PRIVATE FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	SUB:		Set_OffsetMinutes()
	'	PARAMETERS:	
	'	PURPOSE:	Pull out GMT Offset Minutes from registry
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Private Sub Set_OffsetMinutes()
			Dim objWshShell : Set objWshShell = Server.CreateObject("WScript.Shell")
			m_intOffsetMinutes = objWshShell.RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
			Set objWshShell = Nothing
		End Sub

	'**************************************************************************
	'	SUB:		Set_Nonce()
	'	PARAMETERS:	
	'	PURPOSE:	Returns string based on timestamp to be used as "random"
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 
	'**************************************************************************
		Private Sub Set_Nonce()
			m_strNonce = Me.TimeStamp + (Timer() * 1000)
		End Sub

	'**************************************************************************
	'	SUB:		Set_Timestamp()
	'	PARAMETERS:	
	'	PURPOSE:	Returns numnber of seconds from UNIX Epoch Time
	'				January 1, 1970 00:00:00 GMT
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 
	'**************************************************************************
		Private Sub Set_Timestamp()
			Dim dteFrom : dteFrom = "01/01/1970 00:00:00 AM"

			Dim dteNow : dteNow = Now()
				dteNow = DateAdd("n", Me.OffsetMinutes, dteNow)
			
			m_intTimeStamp = DateDiff("s", dteFrom, dteNow)
		End Sub

	End Class
%>