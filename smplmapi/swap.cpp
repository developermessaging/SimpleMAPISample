/*
+---------------------------------------------------------------------
|
|   File:		Swap.cpp
|
|   Purpose:	This is the implementation of the CApp class. It 
|				supports the following features:
|
|	Logging onto and off of the messaging system.
|	Finding a message
|	Getting a message ID
|	Memory managament
|	Getting an e-mail address and resolving an e-mail address
|	Reading a message
|	Sending a message with and without UI and with an attachment
|	Creating a new message
|				
+---------------------------------------------------------------------
*/

#include "swap.h"


CApp::CApp ( ) 
{
	m_lhSession			= 0L;
	m_MAPIAddress		= NULL;
	m_MAPIDetails		= NULL;
	m_MAPIFindNext		= NULL;
	m_MAPIFreeBuffer	= NULL;
	m_MAPILogoff		= NULL;
	m_MAPILogon			= NULL;
	m_MAPIReadMail		= NULL;
	m_MAPIResolveName	= NULL;
	m_MAPISendDocuments	= NULL;
	m_MAPISendMail		= NULL;
	m_MAPISaveMail		= NULL;
}

CApp::~CApp ( ) 
{
	m_lhSession			= 0L;
	m_MAPIAddress		= NULL;
	m_MAPIDetails		= NULL;
	m_MAPIFindNext		= NULL;
	m_MAPIFreeBuffer	= NULL;
	m_MAPILogoff		= NULL;
	m_MAPILogon			= NULL;
	m_MAPIReadMail		= NULL;
	m_MAPIResolveName	= NULL;
	m_MAPISendDocuments = NULL;
	m_MAPISendMail		= NULL;
	m_MAPISaveMail		= NULL;
}

/*
+-------------------------------------------------------------------------------------
|
|	Function:	cAddress()
|
|	Parameters:	[OUT] pcOutRecips == pointer to the count of recipients selected in
|				the address book selection list.
|
|				[OUT] ppOutRecips == pointer to a pointer to an array of 
|				lpMapiRecipDesc structures containing the address specific information
|				of each recipient contained in the array.
|
|	Purpose:	Displays a dialog box that allows the user to select multiple entries 
|				from the address book. The list of selected recipients is stored in 
|				ppOutRecips. The count of selected recipients is stored in pcOutRecips.
|
|	Note:		Any user of this method must release ppOutRecips when done with it 
|				using MAPIFreeBuffer or cFreeBuffer.
+-------------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cAddress ( ULONG *pcOutRecips, lpMapiRecipDesc *ppOutRecips )
{
	HRESULT hRes		= S_OK;
	FLAGS	flFlags		= 0L;
	ULONG	ulReserved	= 0L;
	LPTSTR	lpszCaption	= (LPTSTR) "Console Address List";
	LPTSTR	lpszLabels	= (LPTSTR) "";
	ULONG	nEditFields	= 4L;
	ULONG	cOutRecips	= 0L;
	lpMapiRecipDesc pOutRecips = NULL;
	
	// Always check to make sure there is an active session
	if ( m_lhSession )	
	{
		hRes = m_MAPIAddress(
								m_lhSession,		// Global session handle.
								0L,					// Parent window.  Set to 0 since console app.
								lpszCaption,		// Title of Address window.
								nEditFields,		// Number of edit controls on Address window.
								lpszLabels,			// Label for edit control if nEditControl = 1L.
								0L,					// Number of recipients.
								NULL,				// Address of MapiRecipDesc structure.
								flFlags,			// MAPI_LOGON_UI or MAPI_NEW_SESSION. 0L since explicit session.
								ulReserved,			// Reserved. Must be 0L;
								&cOutRecips,		// Address of number of recipients listed in Address dialog.
								&pOutRecips			// Address of array  of MapiRecipDesc structures.
							 );
		
		if ( hRes == SUCCESS_SUCCESS )
		{ 
			// Set the out parameters accordingly
			*pcOutRecips = cOutRecips;
			*ppOutRecips = pOutRecips;
		}
		else
		{ 
			// Inform the user that an error has occured
			printf ( "Call to MAPIAddress failed due to error code %d.\r\n", hRes );
			
			switch ( hRes )
			{ 
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred while addressing the message. No list of recipient entries was returned.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. No list of recipient entries was returned.\r\n" );
				break;
			case MAPI_E_INVALID_EDITFIELDS:
				printf ( "The value of the nEditFields parameter was outside the range of 0 through 4. No list of recipient entries was returned.\r\n" );
			case MAPI_E_INVALID_RECIPS:
				printf ( "One or more of the recipients in the address list was not valid. No list of recipient entries was returned.\r\n" );
				break;
			case MAPI_E_INVALID_SESSION:
				printf ( "An invalid session handle was used for the lhSession parameter. No list of recipient entries was returned.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No list of recipient entries was returned.\r\n" );
				break;
			case MAPI_E_NOT_SUPPORTED:
				printf ( "The operation was not supported by the underlying messaging system.\r\n" );
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled one of the dialog boxes. No list of recipient entries was returned.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
				break;
			} 
		}
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n" );
	}

	return hRes;
}





/*
+------------------------------------------------------------------------------
|
|	Function:	cCaptureText()
|
|	Parameters:	[IN] lpszPrompt == Text that user will see printed to console
|
|				[OUT] lpszTextOut == Buffer for text captured by console. The
|				Maximum character length permitted by this function is 
|				MAX_TEXT_LENGTH
|
|	Purpose:	Generic text retrieval function. User supplies prompt and 
|				buffer for text storage. Users of this function must free
|				memory pointed to by lpszTextOut using MAPIAllocateBuffer or
|				cFreeBuffer.
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cCaptureText(LPSTR lpszPrompt, LPSTR *lpszTextOut )
{

	// Any user of this method MUST release the lpszTextOut buffer
	// when done with it using MAPIFreeBuffer or cFreeBuffer.

	HRESULT hRes = S_OK;
	char lpszTextIn[MAX_TEXT_LENGTH];
	BOOL HadToForceInit = FALSE;

	// Prompt the user and capture text from console
	printf ( lpszPrompt );
	scanf ( "\n%[^\n]", lpszTextIn );

	// If not logged on, call MAPIInitialize so allocation will succeed.
	if ( MAPI_E_INVALID_SESSION == cValidateSession ( ) )
	{
		HadToForceInit = TRUE;	// Had to force initialization of MAPI
		MAPIInitialize ( NULL );
	}
	
	// Allocate buffer and set out parameter accordingly
	if ( SUCCEEDED ( hRes = MAPIAllocateBuffer ( strlen ( lpszTextIn ) + 1, ( LPVOID * ) lpszTextOut ) ) )		
		strcpy ( *lpszTextOut, lpszTextIn );

	// If we had to force initialization, uninitialize MAPI.
	if ( HadToForceInit )
		MAPIUninitialize ( );

	return hRes;
}




/*
+------------------------------------------------------------------------------
|
|	Function:	cCreateMessage ( )
|
|	Parameters:	[IN] flFlags == Switched for MAPISaveMail.
|
|				[OUT] lppMessage == Pointer to pointer to Message structure.
|
|				[OUT] lppszMessageID == EID of new message. Can be set to NULL.
|
|	Purpose:	Creates a new message and returns it to the caller of this
|				function.
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cCreateMessage( FLAGS flFlags, lpMapiMessage *lppMessage, LPTSTR *lppszMessageID )
{
	HRESULT hRes = S_OK;
	MapiMessage Message;
	char lpszMessageID[MAX_MSGID] = {NULL};
	ULONG ulReserved = 0L;

	// Always make sure there is a valid session
	if ( m_lhSession ) 
	{
		// Clear out memory for Message.
		ZeroMemory ( &Message, sizeof ( MapiMessage ) );

		std::string sPrompt = "A new message from Swap.exe";
		Message.lpszSubject = (LPSTR)(sPrompt.c_str());
		// Save the message
		hRes = m_MAPISaveMail ( m_lhSession,
								0L,
								&Message,
								flFlags,
								ulReserved,
								lpszMessageID );

		// If successful, allocate memory and set out parametres accordingly. Otherwise
		// inform the user of the failure and report the reason for failure.
		if ( SUCCESS_SUCCESS == hRes )
		{
			printf ( "\r\nA new message has been created in your Inbox.\r\n" );
			MAPIAllocateBuffer ( sizeof ( lpMapiMessage ), (LPVOID *) lppMessage );
			*lppMessage = &Message;

			MAPIAllocateBuffer (strlen ( lpszMessageID ) + 1, (LPVOID *) lppszMessageID );
			strcpy ( *lppszMessageID, lpszMessageID );
		}
		else
		{
			printf ( "Call to MAPIFindNext failed due to error code %d.\r\n", hRes );
			switch ( hRes )
			{
			case MAPI_E_ATTACHMENT_NOT_FOUND:
				printf ( "An attachment could not be located at the specified path. Either the drive letter was invalid, the path was not found on that drive, or the file was not found in that path.\r\n"); 
				break;
			case MAPI_E_BAD_RECIPTYPE :
				printf ( "The recipient type in the lpMessage was invalid.\r\n" );
				break;
			case MAPI_E_FAILURE :
				printf ( "One or more unspecified errors occurred while saving the message. No message was saved.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to save the message. No message was saved.\r\n" );
				break;
			case MAPI_E_INVALID_MESSAGE:
				printf ( "An invalid message identifier was passed in the lpszMessageID parameter; no message was saved.\r\n" );
				break;
			case MAPI_E_INVALID_RECIPS:
				printf ( "One or more recipients of the message were invalid or could not be identified.\r\n" );
				break;
			case MAPI_E_INVALID_SESSION:
				printf ( "An invalid session handle was passed in the lhSession parameter. No message was saved.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No message was saved.\r\n" );
				break;
			case MAPI_E_NOT_SUPPORTED:
				printf ( "The operation was not supported by the underlying messaging system.\r\n" );
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled one of the dialog boxes. No message was saved.\r\n" );
				break;
			}
		}
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "\r\nNot logged on to messaging system." );
	}
	return hRes;
}



/*
+------------------------------------------------------------------------------
|
|	Function:	cGetMessageID ( )
|
|	Parameters:	[IN] SeedMsgID == If NULL, gets the first message that meets
|				the conditions set by flFLags parameter. Otherwise, gets the
|				next message that meets the conditions set by flFlags.
|
|				[IN] flFlags == The criterion that describes what messages and
|				what order to retrieve messages.
|
|				[OUT] prgchMsgID == Message EID of a message found by 
|				MAPIFindNext.
|
|	Purpose:	Get the message ID of the next message that meets the criteria
|				defined by flFlags.
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cFindMessageID ( LPTSTR SeedMsgID, FLAGS flFlags, LPTSTR *prgchMsgID )
{
	HRESULT hRes = S_OK;
	ULONG ulReserved = 0L;
	CHAR rgchMsgID[MAX_MSGID];

	hRes = m_MAPIFindNext (
							m_lhSession,	// Global session handle
							0L,				// Parent window. Set to 0 since console app
							NULL,			// NULL specifies interpersonal mail message
							SeedMsgID,		// Seed message ID; NULL == get first message
							flFlags,
							ulReserved,		// Reserved.  Must be 0L
							rgchMsgID
						   );

	if ( hRes == MAPI_E_NO_MESSAGES )
	{
		printf ( "No messages to print.\r\n" );
	}

	if ( hRes != SUCCESS_SUCCESS && hRes != MAPI_E_NO_MESSAGES )
	{
		printf ( "Call to MAPIFindNext failed due to error code %d.\r\n", hRes );
		switch ( hRes )
		{
		case MAPI_E_FAILURE:
			printf ( "One or more unspecified errors occurred while matching the message type. The call failed before message type matching could take place.\r\n" );
			break;
		case MAPI_E_INSUFFICIENT_MEMORY:
			printf ( "There was insufficient memory to proceed. No message was found.\r\n" );
			break;
		case MAPI_E_INVALID_MESSAGE:
			printf ( "An invalid message identifier was passed in the lpszSeedMessageID parameter. No message was found.\r\n" );
			break;
		case MAPI_E_INVALID_SESSION:
			printf ( "An invalid session handle was passed in the lhSession parameter. No message was found.\r\n" );
			break;
		default:
			printf ( "Unknown error code.\r\n" );
			break;
		}
	}
	else
	{
		MAPIAllocateBuffer (strlen ( rgchMsgID ) + 1, (LPVOID *) prgchMsgID );
		strcpy ( *prgchMsgID, rgchMsgID );
	}

	return hRes;
}




/*
+------------------------------------------------------------------------------
|
|	Function:	cFreeBuffer()
|
|	Parameters:	[IN] pv == Address of the buffer to be freed which was 
|				previously allocated by MAPIAllocateBuffer.
|
|
|	Purpose:	Free any buffer allocated by MAPIAllocateBuffer
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cFreeBuffer( LPVOID pv )
{
	HRESULT hRes S_OK;
	
	hRes = m_MAPIFreeBuffer ( pv );
	pv = NULL;

	return hRes;
}



/*
+------------------------------------------------------------------------------
|
|	Function:	cGetDetails()
|
|	Parameters:	[IN] pRecip == The specific recipient to display details of.
|
|	Purpose:	Display the details dialog box provided by the address book
|				of the recipient contained in pRecip.
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cGetDetails ( lpMapiRecipDesc pRecip )
{
	HRESULT hRes = S_OK;
	FLAGS flFlags = 0L;
	ULONG ulReserved = 0L;

	if ( m_lhSession )	  // Always check to make sure there is an active session
	{
		hRes = m_MAPIDetails (
								m_lhSession,	// Global session handle
								0L,				// Parent window. Set to 0 since console app
								pRecip,		
								flFlags,		
								ulReserved		// Reserved. Must be 0L.
							  );

		if ( hRes == SUCCESS_SUCCESS )
			printf ( "Call to MAPIDetails succeeded.\r\n" );	
		else
		{
			printf ( "Call to MAPIDetails failed due to error code %d.\r\n", hRes );
			switch ( hRes )
			{
			case MAPI_E_AMBIGUOUS_RECIPIENT:
				printf ( "The dialog box could not be displayed because the ulEIDSize member of the structure pointed to by the lpRecips parameter was zero.\r\n" );
				break;
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred. No dialog box was displayed.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. No dialog box was displayed.\r\n" );
				break;
			case MAPI_E_INVALID_RECIPS:
				printf ( "The recipient specified in the lpRecip parameter was unknown or the recipient had an invalid ulEIDSize value. No dialog box was displayed.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No dialog box was displayed.\r\n" );
				break;
			case MAPI_E_NOT_SUPPORTED:
				printf ( "The operation was not supported by the underlying messaging system.\r\n" );
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled either the logon dialog box or the details dialog box.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
			}
		}
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n" );
	}
	return hRes;
}



/*
+---------------------------------------------------------------------
|
|	Function:	cInitApp ()
|
|	Purpose:	Makes sure that MAPI is installed on the machine. If 
|				it is installed, we load the function pointers that
|				we will need later.
|
+---------------------------------------------------------------------
*/
STDMETHODIMP CApp::cInitApp ()
{
	HRESULT hRes = S_OK;

	if ( MAPI_INSTALLED == ( hRes = cIsMapiInstalled () ) )
	{
		// Get instance handle of MAPI32.DLL
		HINSTANCE			hlibMAPI		= LoadLibrary ( szMAPIDLL );	
		
		//  Get the addresses of all the API's supported by this object
		m_MAPILogon			= ( LPMAPILOGON			)	GetProcAddress ( hlibMAPI, "MAPILogon"			);
		m_MAPISendMail		= ( LPMAPISENDMAIL		)	GetProcAddress ( hlibMAPI, "MAPISendMail"		);
		m_MAPISendDocuments	= ( LPMAPISENDDOCUMENTS )	GetProcAddress ( hlibMAPI, "MAPISendDocuments"	);
		m_MAPIFindNext		= ( LPMAPIFINDNEXT		)	GetProcAddress ( hlibMAPI, "MAPIFindNext"		);
		m_MAPIReadMail		= ( LPMAPIREADMAIL		)	GetProcAddress ( hlibMAPI, "MAPIReadMail"		);
		m_MAPIResolveName	= ( LPMAPIRESOLVENAME	)	GetProcAddress ( hlibMAPI, "MAPIResolveName"	);
		m_MAPIAddress		= ( LPMAPIADDRESS		)	GetProcAddress ( hlibMAPI, "MAPIAddress"		);
		m_MAPILogoff		= ( LPMAPILOGOFF		)	GetProcAddress ( hlibMAPI, "MAPILogoff"			);
		m_MAPIFreeBuffer	= ( LPMAPIFREEBUFFER	)	GetProcAddress ( hlibMAPI, "MAPIFreeBuffer"		);   
		m_MAPIDetails		= ( LPMAPIDETAILS		)	GetProcAddress ( hlibMAPI, "MAPIDetails"		);
		m_MAPISaveMail		= ( LPMAPISAVEMAIL		)	GetProcAddress ( hlibMAPI, "MAPISaveMail"		);
	}
	return hRes;
}


/*
+---------------------------------------------------------------------
|
|	Function:	cIsMapiInstalled()
|
|	 Purpose:	Determines if MAPI is installed on current 
|				workstation
|
+---------------------------------------------------------------------
*/ 
STDMETHODIMP CApp::cIsMapiInstalled ( void )
{
	HRESULT hRes = S_OK;
	DWORD  SimpleMAPIInstalled;
	char   szAppName[32];
	char   szKeyName[32];
	char   szDefault = {'0'};
	char   szReturn = {'0'};
	DWORD  nSize = 0L;
	char  szFileName[256];

	strcpy ( szAppName, "MAIL" );
	strcpy ( szKeyName, "MAPI" );
	//nSize = strlen ( szReturnString ) + 1;
	nSize = 1;
	strcpy ( szFileName, "WIN.INI" );

	printf ( "\r\nBefore GetPrivateProfileString." );
	printf ( "\r\nszDefault: %c", szDefault );
	printf ( "\r\nszReturn: %c", szReturn );

	SimpleMAPIInstalled = GetPrivateProfileString ( szAppName, 
													szKeyName, 
													&szDefault, 
													&szReturn, 
													nSize, 
													szFileName);

	printf ( "\r\nAfter GetPrivateProfileString." );
	printf ( "\r\nlpDefault: %c", szDefault );
	printf ( "\r\nlpReturn: %c", szReturn );

	if ( MAPI_INSTALLED == strcmp ( &szDefault, &szReturn ) )
	{
		printf ( "\r\nMAPI is not installed.\r\n" );
		hRes = MAPI_NOT_INSTALLED;
	}
	else
	{
		printf ( "\r\nMAPI is installed.\r\n" );
		hRes = MAPI_INSTALLED;
	}

	return hRes;
}





/*
+------------------------------------------------------------------------------
|
|	Function:	cLogoff()
|
|	Purpose:	Logs the user off of the messaging system and invalidates the
|				session handle.					
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cLogoff ( void )
{
	HRESULT hRes = S_OK;
	FLAGS	flFlags = 0L;
	ULONG	ulReserved = 0L;

	// Free any buffers created by MAPI.

	// Always check to make sure there is an active session
	if ( m_lhSession )	 
	{
		//  If there is a valid session handle, attempt to logoff.
		
		hRes = m_MAPILogoff (
								m_lhSession,	// Global session handle.
								0L,				// Parent window.  Set to 0 since console app.
								flFlags,		// Ignored by MAPILogoff.
								ulReserved		// Reserved.  Must be 0L.								
		  					   );				
	
		if ( hRes == SUCCESS_SUCCESS )
		{ 
			// Invalidate session handle and inform user that logoff was successful
			m_lhSession = 0L;
			printf ( "Logoff attempt succeeded.\r\n" );		
		}
		else
		{ 
			// Inform user that logoff attempt failed and report cause.
			printf ( "Logoff attempt failed due to error code %d.\r\n", hRes );  

			switch ( hRes )
			{ 
			case MAPI_E_FAILURE:
				printf ( "The flFlags parameter is invalid or one or more unspecified errors occurred.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. The session was not terminated." );
				break;
			case MAPI_E_INVALID_SESSION:
				printf ( "An invalid session handle was used for the lhSession parameter. The session was not terminated.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
				break;									 
			} 
		}
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n" );
	}

	return hRes;
}




/*
+-------------------------------------------------------------------
|
|	Function:	cLogon()
|
|	Purpose:	Logs user on to messaging system. If logon succeeds
|				inform the user and return SUCCESS_SUCCESS. 
|				Otherwise, notify the user of the error and return
|				the appropriate error value.
|
+-------------------------------------------------------------------
*/
STDMETHODIMP CApp::cLogon ( void )
{
	HRESULT hRes = S_OK;
	FLAGS	flFlags = 0L;
	ULONG	ulReserved = 0L;
	LPSTR	lpszProfileName = NULL;
	LPSTR	lpszPassword = NULL;

	if ( !m_lhSession )	  // Always ask if there is an active session
	{
		flFlags = MAPI_NEW_SESSION |
			      MAPI_LOGON_UI;  // Logon with a new session and force display of UI.

		std::string sPrompt = "\r\nEnter a profile name: ";
		cCaptureText ((LPSTR)sPrompt.c_str(), &lpszProfileName );
		
	    printf ( "Attempting to logon to messaging system.\r\n" );

		hRes = m_MAPILogon (
							 0L, 				// Handle to parent window or 0.
							 lpszProfileName,	// Default profile name to use for MAPI session.
							 lpszPassword,		// User password for MAPI session.
							 flFlags,			// Various session settings
							 ulReserved,		// Reserved.  Must be 0L.
							 &m_lhSession		// Return handle to MAPI Session
						   );

		if (hRes == SUCCESS_SUCCESS)
		{ 
			// Let user know that logon was successful.	

			printf("Logon successful.\r\n");
		} 
		else
		{ 
			// If logon attempt fails, inform user that an error occured and report
			// the error number.
			printf ( "Logon attempt failed due to error code %d.\r\n", hRes );

			switch (hRes)
			{ 
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred during logon. No session handle was returned.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. No session handle was returned.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No session handle was returned.\r\n" );
				break;
			case MAPI_E_TOO_MANY_SESSIONS:
				printf ( "The user had too many sessions open simultaneously. No session handle was returned.\r\n");
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled the logon dialog box. No session handle was returned.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n");
				break;
			} 
		}
	}
	else
	{
		// If we get to this point, the user is already logged on. A new attempt 
		// to logon is redundant so inform the user they have a session handle 
		// already.
		printf ( "Already logged on to messaging system.\r\n" );
	}

	cFreeBuffer ( lpszProfileName );

	return hRes;
}





/*
+------------------------------------------------------------------------------
|
|	Function:	cReaddMail ( )
|
|	Parameters:	[IN] ReadFlags == Options for level of detail. 
|
|					 MESSAGE_HEADERS_ONLY is the only valid option for this 
|					 flag. If present, will only read the subject heading.
|
|				[IN] prgchMsgID == The message EID to read. Can be NULL.
|
|	Purpose:	Displays the contents of a message to the user.
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cReadMail ( ULONG ReadFlags, LPTSTR prgchMsgID )
{
	HRESULT hRes = S_OK;
	FLAGS flFlags = 0L;
	ULONG ulReserved = 0L;
	lpMapiMessage lpMessage = NULL;

	if ( m_lhSession )	   // Always check to make sure there is an active session
	{	
		if ( SUCCESS_SUCCESS == ( hRes = cFindMessageID ( NULL, 
										 				  MAPI_LONG_MSGID |
												          MAPI_UNREAD_ONLY, 
												          &prgchMsgID ) ) )
		{
			hRes = m_MAPIReadMail (
										m_lhSession,
										0L,
										prgchMsgID,
										flFlags,
										ulReserved,
										&lpMessage
									);
		}

		if ( hRes != SUCCESS_SUCCESS )
		{
			printf ( "Error retrieving message %s.\r\n", prgchMsgID );
			switch ( hRes )
			{
			case MAPI_E_ATTACHMENT_WRITE_FAILURE:
				printf ( "An attachment could not be written to a temporary file. Check directory permissions.\r\n" );
				break;
			case MAPI_E_DISK_FULL:
				printf ( "An attachment could not be written to a temporary file because there was not enough space on the disk.\r\n" );
				break;
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred while reading the message.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to read the message.\r\n" );
				break;
			case MAPI_E_INVALID_MESSAGE:
				printf ( "An invalid message identifier was passed in the lpszMessageID parameter.\r\n" );
				break;
			case MAPI_E_INVALID_SESSION:
				printf ( "An invalid session handle was passed in the lhSession parameter. No message was retrieved.\r\n" );
				break;
			case MAPI_E_TOO_MANY_FILES:
				printf ( "There were too many file attachments in the message. The message could not be read.\r\n" );
				break;
			case MAPI_E_TOO_MANY_RECIPIENTS:
				printf ( "There were too many recipients of the message. The message could not be read.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
				break;
			}
		}
		else
		{
			if ( lpMessage -> lpszSubject != NULL &&
				lpMessage -> lpszSubject[0] != '\0' )
			{			
				printf ("Subject: %s\r\n", (LPSTR)lpMessage -> lpszSubject );
			}

			if ( lpMessage -> lpszNoteText != NULL &&
				lpMessage -> lpszNoteText[0] != '\0' && MESSAGE_HEADERS_ONLY != ReadFlags)
			{
				printf ( "Message Text: %s\r\n", (LPSTR)lpMessage -> lpszNoteText );
			}
			else
				printf ( "No message text.\r\n" );
		
			if ( lpMessage -> nFileCount > 0 )
			{
				for ( int i = 0; i < (int)lpMessage -> nFileCount; i++ )
					printf ( "\r\n Attatchments: %s", (LPSTR)lpMessage -> lpFiles[i].lpszFileName );
			}
		}						
	} 
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf (" Not logged on to messaging system.\r\n");
	}

		hRes = MAPIFreeBuffer ( prgchMsgID );
		hRes = MAPIFreeBuffer ( lpMessage );

		prgchMsgID = NULL;
		lpMessage = NULL;

	return hRes;
}




		
/*
+------------------------------------------------------------------------------
|
|	Function:	cResolveName()
|
|	Parameters:	[IN]	lpszName = Name of e-mail recipient to resolve.
|				[OUT]	ppRecips = Pointer to a pointer to an lpMapiRecipDesc
|
|	Purpose:	Resolves an e-mail address and returns a pointer to a 
|				MapiRecipDesc structure filled with the recipient information
|				contained in the address book.
|
|	Note:		ppRecips is allocated off the heap using MAPIAllocateBuffer.
|				Any user of this method must be sure to release ppRecips when 
|				done with it using either MAPIFreeBuffer or cFreeBuffer.
+-------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cResolveName( LPSTR lpszName, lpMapiRecipDesc *ppRecips )
{	
	HRESULT hRes = S_OK;
	FLAGS flFlags = 0L;
	ULONG ulReserved = 0L;
	lpMapiRecipDesc pRecips = NULL;
	
	// Always check to make sure there is an active session
	if ( m_lhSession )		
	{
		// This method is less automated than cAddress. It does not
		// offer the user a dialog box to choose names from. It accepts input
		// in the form of a paramter passed into it in the form of a recipient
		// name and returns a valid MAPIRecipDesc structure.
		
		// cAddress automates this process and allows the user to select from
		// a list of possible recipients.
		
		hRes = m_MAPIResolveName (
								     m_lhSession,	// Global session handle
									 0L,			// Parent window.  Since console, set to 0L.
									 lpszName,		// Name of recipient.  Passed in by argv.
									 flFlags,		// Flags set to 0 for MAPIResolveName.
									 ulReserved,
									 &pRecips
								  );				

		if ( hRes == SUCCESS_SUCCESS )
		{  
			// Copy the recipient descriptor returned from MAPIResolveName to 
			// the out parameter for this function and inform user that 
			// MAPIResolveName was successful
			*ppRecips = pRecips;
			printf("%s resolved to a single address.\r\n", pRecips -> lpszName);	

		}  
		else
		{  
			// Inform user that MAPIResolveName failed and report error number.
			printf ( "%s did not resolve to a single address.\r\n", lpszName );  
			printf ( "The error code was %d \r\n", hRes );						
			
			switch (hRes)
			{ 
			case MAPI_E_AMBIGUOUS_RECIPIENT:
				printf ( "The recipient requested has not or could not be resolved to a unique address list entry.\r\n" );
				break;
			case MAPI_E_UNKNOWN_RECIPIENT:
				printf ( "The recipient could not be resolved to any address. The recipient might not exist or might be unknown.\r\n" );
				break;
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred. The name was not resolved.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. The name was not resolved.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. The name was not resolved.\r\n" );
				break;
			case MAPI_E_NOT_SUPPORTED:
				printf ( "The operation was not supported by the underlying messaging system.\r\n" );
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled one of the dialog boxes. The name was not resolved.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
				break;
			} 
		} 
		
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n" );
	}
	
	return hRes;
}









/*
+------------------------------------------------------------------------------
|
|	Function:	cSendAttachMail ( )
|
|	Purpose:	Sends a message with an attachment. Asks user for file name and
|				path and links or embedds the attachment to the message.
|
+-------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cSendAttachMail ( )
{
	HRESULT hRes = S_OK;
	ULONG ulReserved = 0L;
	ULONG cRecips = 0L;
	LPSTR lpszFileName = NULL,
		  lpszPathName = NULL;
	
	char lpszFullPath[256] = {NULL};

	lpMapiRecipDesc pRecips = NULL;
	MapiMessage Message;
	MapiFileDesc pFileDesc;
	
	ZeroMemory ( &Message, sizeof ( MapiMessage ) );
	ZeroMemory ( &pFileDesc, sizeof ( MapiFileDesc ) );

	if ( m_lhSession )	 // Always check to make sure there is an active session
	{		
		// Populate members of Message structure.
		LPSTR lpszName = NULL;
		std::string sPrompt = "\r\nEnter an e-mail address: ";
		cCaptureText ((LPSTR)sPrompt.c_str(), &lpszName );

		cResolveName ( lpszName, &pRecips );
		Message.nRecipCount		= 1L;		// Must be set to the correct number of recipients.
		Message.lpRecips		= pRecips;	// Address of list of names returned from MAPIAddress.		
	
		// Capture the file name and path name. strcat the file name to the end of the
		// path name
		sPrompt = "\r\nEnter file name (e.g. win.ini): ";
		cCaptureText ((LPSTR)(sPrompt.c_str()), &lpszFileName );
		sPrompt = "\r\nEnter path (e.g. c:\\windows\\): ";
		cCaptureText ((LPSTR)(sPrompt.c_str()), &lpszPathName );

		strcat ( lpszFullPath, lpszPathName );
		strcat ( lpszFullPath, lpszFileName );
		
		// Set the file and path name members of the MapiFileDesc.
		pFileDesc.lpszFileName = lpszFileName;
		pFileDesc.lpszPathName = lpszFullPath;
	
		// Set the other members of the MapiMessage structure.
		// We only support 1 attachment so nFileCount gets set to 1.
		Message.ulReserved		= ulReserved;
		Message.lpszSubject		= ( LPTSTR ) "Any subject";
		Message.lpszNoteText	= ( LPTSTR ) "Any note text";
		Message.lpOriginator	= NULL;			
		Message.nFileCount		= 1L;
		Message.lpFiles			= &pFileDesc;
		
		hRes = m_MAPISendMail (	m_lhSession,	// Global session handle.
								0L,				// Parent window. Set to 0 since console app.
								&Message,		// Address of Message structure
								0L,		
								ulReserved		// Reserved. Must be 0L.
						      );
		
		if ( hRes == SUCCESS_SUCCESS )
		{ 
			// Inform user that MAPISendMail was successful
			ZeroMemory ( &Message, sizeof ( MapiMessage ) );
			printf( "Message successfully sent.\r\n" ); 
		} 
		else
		{ 
			// Inform user that MAPSendMail failed and report the error number.
			printf( "Message did not get sent due to error code %d.\r\n", hRes ); 
			switch (hRes)
			{  
			case MAPI_E_AMBIGUOUS_RECIPIENT:
				printf ( "A recipient matched more than one of the recipient descriptor structures and MAPI_DIALOG was not set. No message was sent.\r\n" );
				break;
			case MAPI_E_ATTACHMENT_NOT_FOUND:
				printf ( "The specified attachment was not found. No message was sent.\r\n" );
				break;
			case MAPI_E_ATTACHMENT_OPEN_FAILURE:
				printf ( "The specified attachment could not be opened. No message was sent.\r\n" );
				break;
			case MAPI_E_BAD_RECIPTYPE:
				printf ( "The type of a recipient was not MAPI_TO, MAPI_CC, or MAPI_BCC. No message was sent.\r\n" );
				break;
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred. No message was sent.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. No message was sent.\r\n" );
				break;
			case MAPI_E_INVALID_RECIPS:
				printf ( "One or more recipients were invalid or did not resolve to any address.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No message was sent.\r\n" );
				break;
			case MAPI_E_TEXT_TOO_LARGE:
				printf ( "The text in the message was too large. No message was sent.\r\n" );
				break;
			case MAPI_E_TOO_MANY_FILES:
				printf ( "There were too many file attachments. No message was sent.\r\n" );
				break;
			case MAPI_E_TOO_MANY_RECIPIENTS:
				printf ( "There were too many recipients. No message was sent.\r\n" );
				break;
			case MAPI_E_UNKNOWN_RECIPIENT:
				printf ( "A recipient did not appear in the address list. No message was sent.\r\n" );
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled one of the dialog boxes. No message was sent.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
				break;
			}
		}		
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n" );
	}

	m_MAPIFreeBuffer ( pRecips );
	cFreeBuffer ( lpszPathName );
	cFreeBuffer ( lpszFileName );
	
	return hRes;
}



/*
+------------------------------------------------------------------------------
|
|	Function:	cSendMessage ( )
|
|	Parameters: [IN] flFlags == Flags to be set when sending the message.
|
|   Purpose:	This method allows the user to enter 1 and only 1 recipient
|				to send a hard coded message to.
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cSendMessage ( FLAGS flFlags )
{
	HRESULT hRes = S_OK;
	ULONG ulReserved = 0L;
	ULONG cRecips = 0L;
	lpMapiRecipDesc pRecips = NULL;
	MapiMessage Message;
	
	ZeroMemory ( &Message, sizeof ( MapiMessage ) );	
	
	// Always make sure there is an active session
	if ( m_lhSession )	 
	{		
		// Populate members of Message structure.
		// If no dialog is to be displayed capture the recipient to 
		// deliver the message to.
		if ( MAPI_DIALOG != flFlags )
		{
			LPSTR lpszName = NULL;
			std::string sPrompt = "\r\nEnter and e-mail address: ";
			cCaptureText ((LPSTR)sPrompt.c_str(), &lpszName );

			cResolveName ( lpszName, &pRecips );
			Message.nRecipCount		= 1L;		// Must be set to the correct number of recipients.
			Message.lpRecips		= pRecips;	// Address of list of names returned from MAPIAddress.		
		}
	
		Message.ulReserved		= ulReserved;
		Message.lpszSubject		= ( LPTSTR ) "Any subject";
		Message.lpszNoteText	= ( LPTSTR ) "Any note text";
		Message.lpOriginator	= NULL;			
		Message.nFileCount		= 0L;
		
		hRes = m_MAPISendMail (	m_lhSession,	// Global session handle.
								0L,				// Parent window.  Set to 0 since console app.
								&Message,		// Address of Message structure
								flFlags,		
								ulReserved		// Reserved.  Must be 0L.
							   );									

		if ( hRes == SUCCESS_SUCCESS )
		{ 
			// Inform user that MAPISendMail was successful
			printf( "Message successfully sent.\r\n" ); 
		} 
		else
		{ 
			// Inform user that MAPSendMail failed and report the error number.
			printf( "Message did not get sent due to error code %d.\r\n", hRes ); 
			switch (hRes)
			{  
			case MAPI_E_AMBIGUOUS_RECIPIENT:
				printf ( "A recipient matched more than one of the recipient descriptor structures and MAPI_DIALOG was not set. No message was sent.\r\n" );
				break;
			case MAPI_E_ATTACHMENT_NOT_FOUND:
				printf ( "The specified attachment was not found. No message was sent.\r\n" );
				break;
			case MAPI_E_ATTACHMENT_OPEN_FAILURE:
				printf ( "The specified attachment could not be opened. No message was sent.\r\n" );
				break;
			case MAPI_E_BAD_RECIPTYPE:
				printf ( "The type of a recipient was not MAPI_TO, MAPI_CC, or MAPI_BCC. No message was sent.\r\n" );
				break;
			case MAPI_E_FAILURE:
				printf ( "One or more unspecified errors occurred. No message was sent.\r\n" );
				break;
			case MAPI_E_INSUFFICIENT_MEMORY:
				printf ( "There was insufficient memory to proceed. No message was sent.\r\n" );
				break;
			case MAPI_E_INVALID_RECIPS:
				printf ( "One or more recipients were invalid or did not resolve to any address.\r\n" );
				break;
			case MAPI_E_LOGIN_FAILURE:
				printf ( "There was no default logon, and the user failed to log on successfully when the logon dialog box was displayed. No message was sent.\r\n" );
				break;
			case MAPI_E_TEXT_TOO_LARGE:
				printf ( "The text in the message was too large. No message was sent.\r\n" );
				break;
			case MAPI_E_TOO_MANY_FILES:
				printf ( "There were too many file attachments. No message was sent.\r\n" );
				break;
			case MAPI_E_TOO_MANY_RECIPIENTS:
				printf ( "There were too many recipients. No message was sent.\r\n" );
				break;
			case MAPI_E_UNKNOWN_RECIPIENT:
				printf ( "A recipient did not appear in the address list. No message was sent.\r\n" );
				break;
			case MAPI_E_USER_ABORT:
				printf ( "The user canceled one of the dialog boxes. No message was sent.\r\n" );
				break;
			default:
				printf ( "Unknown error code.\r\n" );
				break;
			}
		}		
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n" );
	}

	m_MAPIFreeBuffer ( pRecips );

	return hRes;
}





/*
+------------------------------------------------------------------------------
|
|	Function:	cValidateSession()
|
|	Purpose:	Determines if there is a valid session handle. If there is no
|				session handle, return MAPI_E_INVALID_SESSION. Otherwise return
|				SUCCESS_SUCCESS or S_OK;
|
+------------------------------------------------------------------------------
*/
STDMETHODIMP CApp::cValidateSession()
{
	HRESULT hRes = S_OK;
	
	// If m_lhSession is 0, then there is no active session.
	if ( 0L == m_lhSession )
		hRes = MAPI_E_INVALID_SESSION;

	return hRes;
}

STDMETHODIMP CApp::cListInboxMessages()
{
	HRESULT hRes = S_OK;
	char szMsgID[512];
    char szSeedMsgID[512];
    lpMapiMessage lpMessage;

	if ( m_lhSession )
	{
		/* Populate List Box with all messages in InBox. */
		/* This is a painfully slow process for now.     */

		hRes = m_MAPIFindNext ( m_lhSession, 0L, NULL, NULL,
			MAPI_GUARANTEE_FIFO | MAPI_LONG_MSGID, 0, szMsgID);

		while (hRes == SUCCESS_SUCCESS)
		{
			hRes = m_MAPIReadMail ( m_lhSession, 
									0L, 
									szMsgID,
									MAPI_PEEK | 
									MAPI_ENVELOPE_ONLY,
									0, 
									&lpMessage);

			if (SUCCESS_SUCCESS == hRes)
			{
				strcat ( lpMessage->lpszSubject, "\r\n" );
				printf ( lpMessage->lpszSubject );
			}
			MAPIFreeBuffer (lpMessage);

			lstrcpy (szSeedMsgID, szMsgID);
			hRes = m_MAPIFindNext (m_lhSession, 0L, NULL, szSeedMsgID,
				MAPI_GUARANTEE_FIFO | MAPI_LONG_MSGID, 0, szMsgID);
		}
	}
	else
	{
		hRes = MAPI_E_INVALID_SESSION;
		printf ( "Not logged on to messaging system.\r\n");
	}

		return hRes;
}



