<%
'******************************************************************************
'	ERROR CODE CONSTANTS (do NOT change)
'******************************************************************************
	Const OAUTH_ERROR_TIMEOUT = "-2147012894"

'******************************************************************************
'	TIMEOUT CONSTANTS (change to suite)
'******************************************************************************
	Const OAUTH_TIMEOUT_RESOLVE = 2500
	Const OAUTH_TIMEOUT_CONNECT = 10000
	Const OAUTH_TIMEOUT_SEND = 10000
	Const OAUTH_TIMEOUT_RECEIVE = 10000

'******************************************************************************
'	PARAM CONSTANTS (do NOT change)
'******************************************************************************
	Const OAUTH_TOKEN_REQUEST = "oauth_token_request"
	Const OAUTH_TOKEN = "oauth_token"
	Const OAUTH_TOKEN_SECRET = "oauth_token_secret"
	Const OAUTH_VERIFIER = "oauth_verifier"

'******************************************************************************
'	MISC CONSTANTS (do NOT change)
'******************************************************************************
	Const OAUTH_UNRESERVED = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~"
	Const OAUTH_REQUEST_METHOD_GET = "GET"
	Const OAUTH_REQUEST_METHOD_POST = "POST"
	Const OAUTH_SIGNATURE_METHOD = "HMAC-SHA1"
	Const OAUTH_VERSION = "1.0" 'editable?
%>