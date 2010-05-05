<%
'******************************************************************************
'	CLASS:		cLibOAuthRequestURL
'	PURPOSE:	
'
'	AUTHOR:	sdesapio		DATE: 04.04.10			LAST MODIFIED: 04.04.10
'******************************************************************************
	Class cLibOAuthRequestURL
	'**************************************************************************
'***'PRIVATE CLASS MEMBERS
	'**************************************************************************
		Private m_objUtils 
		Private m_strConsumerSecret
		Private m_strEndPoint
		Private m_strMethod
		Private m_strParameters
		Private m_strTokenSecret

	'**************************************************************************
'***'CLASS_INITIALIZE / CLASS_TERMINATE
	'**************************************************************************
		Private Sub Class_Initialize()
			Set m_objUtils = Nothing
			m_strTokenSecret = ""
		End Sub
		Private Sub Class_Terminate()
			Set m_objUtils = Nothing
		End Sub

	'**************************************************************************
'***'PUBLIC PROPERTIES
	'**************************************************************************
		Public Property Let ConsumerSecret(pData)
			m_strConsumerSecret = pData
		End Property

		Public Property Let EndPoint(pData)
			m_strEndPoint = pData
		End Property

		Public Property Let Method(pData)
			m_strMethod = pData
		End Property

		Public Property Let Parameters(pData)
			m_strParameters = pData
		End Property

		Public Property Let TokenSecret(pData)
			m_strTokenSecret = pData
		End Property

	'**************************************************************************
'***'PRIVATE PROPERTIES
	'**************************************************************************
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
	'	FUNCTION:		Get_RequestURL
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Function Get_RequestURL()
			Dim strSignature : strSignature = Get_Signature()
				strSignature = Utils.URLEncode(strSignature)

			Get_RequestURL = _
				m_strEndPoint & "?" & _
				m_strParameters & "&" & _
				"oauth_signature=" & strSignature
		End Function

	'**************************************************************************
'***'PRIVATE FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	FUNCTION:		Get_Signature
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Private Function Get_Signature()
			Dim strBaseSignature : strBaseSignature = _
				m_strMethod & "&" & _
				Utils.URLEncode(m_strEndPoint) & "&" & _
				Utils.URLEncode(m_strParameters)

			Dim strSecret : strSecret = _
				m_strConsumerSecret & "&" & _
				m_strTokenSecret

			Get_Signature = b64_hmac_sha1(strSecret, strBaseSignature)
		End Function

	End Class
%>