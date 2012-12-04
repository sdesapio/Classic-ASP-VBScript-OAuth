<!--#include file="_inc/_base.asp"-->
<!--#include file="cLibOAuth.QS.asp"-->
<!--#include file="cLibOAuth.RequestURL.asp"-->
<!--#include file="cLibOAuth.Utils.asp"-->
<%

'******************************************************************************
'	CLASS:		cLibOAuth
'	PURPOSE:	
'
'	AUTHOR:	sdesapio		DATE: 04.04.10			LAST MODIFIED: 04.04.10
'******************************************************************************
	Class cLibOAuth
	'**************************************************************************
'***'PRIVATE CLASS MEMBERS
	'**************************************************************************

		' boolean indicating the users current logged in state. This state
		' variable is EXCLUSIVE to the session state of the local application 
		' and NOT the oAuth provider's logged in state.
		Private m_blnLoggedIn

		' additional parameters exclusive to the current call
		Private m_objParameters

		' reference to the utilities class (Encoding, response extraction, 
		' dictionary sorting, etc.)
		Private m_objUtils 

		' reference to the consumer key acquired after registering with the
		' oAuth service provider
		Private m_strConsumerKey

		' reference to the consumer secret acquired after registering with the
		' oAuth service provider
		Private m_strConsumerSecret

		' the request URL
		Private m_strEndPoint

		' used to globally identify process errors
		Private m_strErrorCode

		' used to set host header
		Private m_strHost

		' the request type - e.g. POST, GET
		Private m_strRequestMethod

		' the response string returned by the service provider 
		Private m_strResponseText

		' where to forward the user if call to oAuth provider times out. 
		' Absolute URL is recommended
		Private m_strTimeoutURL

		' used to set user-agent header
		Private m_strUserAgent

	'**************************************************************************
'***'CLASS_INITIALIZE / CLASS_TERMINATE
	'**************************************************************************
		Private Sub Class_Initialize()
			' set default value to Null so we can check for null before get/set
			m_blnLoggedIn = Null

			' set default value to Nothing so we can check "If ... Is Nothing"
			Set m_objParameters = Nothing

			' instantiate the Utils class
			Set m_objUtils = New cLibOAuthUtils

			' set default to Null to ensure we're returning a verifiable value
			m_strErrorCode = Null

			' set default to POST
			m_strRequestMethod = OAUTH_REQUEST_METHOD_POST 
		End Sub
		Private Sub Class_Terminate()
			' kill obj refs
			Set m_objUtils = Nothing
			Set m_objParameters = Nothing
		End Sub

	'**************************************************************************
'***'PUBLIC PROPERTIES
	'**************************************************************************
		Public Property Let ConsumerKey(pData)
			m_strConsumerKey = pData
		End Property

		Public Property Let ConsumerSecret(pData)
			m_strConsumerSecret = pData
		End Property

		Public Property Let EndPoint(pData)
			m_strEndPoint = pData
		End Property

		Public Property Get ErrorCode
			ErrorCode = m_strErrorCode
		End Property

		Public Property Let Host(pData)
			m_strHost = pData
		End Property

		Public Property Get LoggedIn
			If IsNull(m_blnLoggedIn) Then
				Call Get_LoggedIn()
			End If
			
			LoggedIn = m_blnLoggedIn 
		End Property

		Public Property Get Parameters
			If m_objParameters Is Nothing Then
				Set m_objParameters = Server.CreateObject("Scripting.Dictionary")
			End If

			Set Parameters = m_objParameters
		End Property

		Public Property Let RequestMethod(pData)
			m_strRequestMethod = pData
		End Property

		Public Property Get ResponseText
			ResponseText = m_strResponseText
		End Property

		Public Property Let TimeoutURL(pData)
			m_strTimeoutURL = pData
		End Property

		Public Property Let UserAgent(pData)
			m_strUserAgent = pData
		End Property

	'**************************************************************************
'***'PUBLIC FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	SUB:		Send()
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 12.04.12 
	'**************************************************************************
		Public Sub Send()
			' build Request URL
			Dim strRequestURL : strRequestURL = Get_RequestURL()

			' make the call
			On Error Resume Next

			Dim objXMLHTTP : Set objXMLHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
				objXMLHTTP.setTimeouts OAUTH_TIMEOUT_RESOLVE, OAUTH_TIMEOUT_CONNECT, OAUTH_TIMEOUT_SEND, OAUTH_TIMEOUT_RECEIVE
				objXMLHTTP.Open m_strRequestMethod, strRequestURL, False
				objXMLHTTP.SetRequestHeader "Content-Type","application/x-www-form-urlencoded"
				objXMLHTTP.SetRequestHeader "User-Agent", m_strUserAgent
				objXMLHTTP.SetRequestHeader "Host", m_strHost

				objXMLHTTP.Send()

			' check for errors
			If Err.Number <> 0 Then
				Select Case CStr(Err.Number)
					Case CStr(OAUTH_ERROR_TIMEOUT)
						Response.Redirect m_strTimeoutURL
						Response.End
					Case Else
						m_strErrorCode = Err.Number
				End Select
			Else
				m_strResponseText = objXMLHTTP.ResponseText
			End If

			Set objXMLHTTP = Nothing

			On Error Goto 0
		End Sub

	'**************************************************************************
	'	FUNCTION:		Get_ResponseValue()
	'	PARAMETERS:		strParamName
	'	PURPOSE:		Returns a value ripped from service provider response
	'
	'	AUTHOR:	sdesapio		DATE: 04.04.10		LAST MODIFIED: 04.04.10
	'**************************************************************************
		Public Function Get_ResponseValue(strParamName)
			Get_ResponseValue = m_objUtils.Get_ResponseValue(m_strResponseText, strParamName)
		End Function

	'**************************************************************************
'***'PRIVATE FUNCTIONS
	'**************************************************************************
	'**************************************************************************
	'	SUB:		Get_LoggedIn
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sjd		DATE: 			LAST MODIFIED: 
	'**************************************************************************
		Private Sub Get_LoggedIn()
			On Error Resume Next

			If Session(OAUTH_TOKEN) <> "" And Session(OAUTH_TOKEN_SECRET) <> "" Then
				m_blnLoggedIn = True
			Else
				m_blnLoggedIn = False
			End If

			If Err.Number <> 0 Then
				m_blnLoggedIn = Null
			End If

			On Error Goto 0
		End Sub

	'**************************************************************************
	'	FUNCTION:	Get_RequestURL
	'	PARAMETERS:	
	'	PURPOSE:	
	'
	'	AUTHOR:	sjd		DATE: 			LAST MODIFIED: 
	'**************************************************************************
		Private Function Get_Parameters()
			Dim objQS : Set objQS = New cLibOAuthQS

			' add proprieatary param set
			If Not m_objParameters Is Nothing Then
				Dim Item : For Each Item In m_objParameters
					objQS.Add Item, m_objParameters.Item(Item)
				Next
			End If

			' add required standard param set
			objQS.Add "oauth_consumer_key", m_strConsumerKey
			objQS.Add "oauth_nonce", m_objUtils.Nonce
			objQS.Add "oauth_signature_method", OAUTH_SIGNATURE_METHOD
			objQS.Add "oauth_timestamp", m_objUtils.TimeStamp
			objQS.Add "oauth_version", OAUTH_VERSION

			Get_Parameters = objQS.Get_Parameters()

			Set objQS = Nothing
		End Function

	'**************************************************************************
	'	FUNCTION:	Get_RequestURL
	'	PARAMETERS:	strParameters
	'	PURPOSE:	Returns a fully formatted request URL
	'
	'	AUTHOR:	sjd		DATE: 			LAST MODIFIED: 
	'**************************************************************************
		Private Function Get_RequestURL()
			Dim strParameters : strParameters = Get_Parameters()

			Dim objRequestURL : Set objRequestURL = New cLibOAuthRequestURL
				objRequestURL.ConsumerSecret = m_strConsumerSecret
				objRequestURL.EndPoint = m_strEndPoint
				objRequestURL.Method = m_strRequestMethod
				objRequestURL.Parameters = strParameters
				objRequestURL.TokenSecret = Session(OAUTH_TOKEN_SECRET)

			Get_RequestURL = objRequestURL.Get_RequestURL()

			Set objRequestURL = Nothing
		End Function

	End Class
%>