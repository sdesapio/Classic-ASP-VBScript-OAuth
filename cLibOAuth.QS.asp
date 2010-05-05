<%
'******************************************************************************
'	CLASS:		cLibOAuthQS
'	PURPOSE:	
'
'	AUTHOR:	sdesapio		DATE: 04.04.10			LAST MODIFIED: 04.04.10
'******************************************************************************
	Class cLibOAuthQS
	'**************************************************************************
'***'PRIVATE CLASS MEMBERS
	'**************************************************************************
		Private m_objDictionary
		Private m_objUtils 
		Private m_strSorted

	'**************************************************************************
'***'CLASS_INITIALIZE / CLASS_TERMINATE
	'**************************************************************************
		Private Sub Class_Initialize()
			Set m_objDictionary = Nothing
			Set m_objUtils = Nothing
			m_strSorted = Null
		End Sub
		Private Sub Class_Terminate()
			Set m_objUtils = Nothing
			Set m_objDictionary = Nothing
		End Sub

	'**************************************************************************
'***'PUBLIC PROPERTIES
	'**************************************************************************

	'**************************************************************************
'***'PRIVATE PROPERTIES
	'**************************************************************************
		Private Property Get Dictionary
			If m_objDictionary Is Nothing Then
				Set m_objDictionary = Server.CreateObject("Scripting.Dictionary")
			End If

			Set Dictionary = m_objDictionary
		End Property		

		Private Property Get Sorted
			If IsNull(m_strSorted) Then
				Call Get_Sorted()
			End If

			Sorted = m_strSorted
		End Property		

		Private Property Get Utils
			If m_objUtils Is Nothing Then
				Set m_objUtils = New cLibOAuthUtils
			End If

			Set Utils = m_objUtils
		End Property		

	'**************************************************************************
'***'PUBLIC FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	SUB:			Add
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Sub Add(strKey, strValue)
			Dictionary.Add strKey, Utils.URLEncode(strValue)
		End Sub

	'**************************************************************************
	'	SUB:			Get_Parameters
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Function Get_Parameters()
			Get_Parameters = Sorted
		End Function

	'**************************************************************************
'***'PRIVATE FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	SUB:			Get_Sorted
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Private Sub Get_Sorted()
			Dim intCount : intCount = Dictionary.Count
			Dim i : i = 1

			m_strSorted = ""

			Call Utils.SortDictionary(Dictionary, 1)

			Dim Item : For Each Item In Dictionary
				m_strSorted = m_strSorted & Item & "=" & Dictionary.Item(Item)

				If i < intCount Then
					m_strSorted = m_strSorted & "&"
				End If

				i = i + 1
			Next
		End Sub

	End Class
%>