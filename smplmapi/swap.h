/*
+---------------------------------------------------------------------
|
|   File:		Swap.h
|
|   Purpose:	This is the header file for the application.  It 
|				contains all of the constants and function declarations
|				and is responsible for bringing all of the other
|				include files needed by the app.
|				
+---------------------------------------------------------------------
*/


#ifndef _SWAP_H
#define _SWAP_H

#pragma comment (lib, "mapi32.lib")
//  Be sure to specify what bitness you are compiling for
#define _WIN32
//#define _WIN16

#define STRICT

#include <objbase.h>
#include <windows.h>			// Standard windows headers
#include <windowsx.h>
#include <winuser.h>
#include <winerror.h>

#include <stdlib.h>				// Standard library header file.
#include <stdio.h>				// Standard I/O stream header file.
#include <string.h>				// Standard string manipulation header file.
#include <malloc.h>				// Memory management header file.
#include <mapi.h>				// MAPI Header file.
#include <mapix.h>
#include <string>

// If compiling for 32 bit platforms we will use MAPI32.DLL
// If compiling for 16 bit platforms we will use MAPI.DLL
#ifdef _WIN32
#define szMAPIDLL			"MAPI32.DLL"
#else 
#define szMAPIDLL			"MAPI.DLL"
#endif

#define MAPI_NOT_INSTALLED	1
#define MAPI_INSTALLED		SUCCESS_SUCCESS
#define MAX_MSGID			512
#define MAX_TEXT_LENGTH		256
#define MESSAGE_HEADERS_ONLY		1

/* Structure Definitions */


class CApp
{

private:

	LHANDLE		m_lhSession;			// Used to capture MAPI session handle.

	// Handles for Simple MAPI functions. Simple MAPI requires us to 
	// call its functions by obtaining the address of the function.
	LPMAPILOGON			m_MAPILogon;		
	LPMAPILOGOFF		m_MAPILogoff;			
	LPMAPISENDMAIL		m_MAPISendMail;		
	LPMAPISENDDOCUMENTS	m_MAPISendDocuments;	
	LPMAPIFINDNEXT		m_MAPIFindNext;		
	LPMAPIREADMAIL		m_MAPIReadMail;		
	LPMAPIRESOLVENAME	m_MAPIResolveName;	
	LPMAPIADDRESS		m_MAPIAddress;		
	LPMAPIFREEBUFFER	m_MAPIFreeBuffer;		
	LPMAPIDETAILS		m_MAPIDetails;
	LPMAPISAVEMAIL		m_MAPISaveMail;

public:
	STDMETHOD(cListInboxMessages )( );
		
	CApp ( );
	~CApp ( );	
	STDMETHODIMP cAddress			( ULONG *, lpMapiRecipDesc * );
	STDMETHODIMP cCaptureText		( LPSTR, LPSTR * );
	STDMETHODIMP cCreateMessage		( FLAGS, lpMapiMessage *, LPTSTR * );
	STDMETHODIMP cFindMessageID		( LPTSTR, FLAGS, LPTSTR *);
	STDMETHODIMP cFreeBuffer		( LPVOID );
	STDMETHODIMP cGetDetails		( lpMapiRecipDesc );
	STDMETHODIMP cInitApp			( void );
	STDMETHODIMP cIsMapiInstalled	( void );
	STDMETHODIMP cLogoff			( void );
	STDMETHODIMP cLogon				( void );
	STDMETHODIMP cResolveName		( LPSTR, lpMapiRecipDesc * );
	STDMETHODIMP cSendMessage		( FLAGS );
	STDMETHODIMP cSendAttachMail	( );
	STDMETHODIMP cReadMail			( ULONG, LPTSTR );
	STDMETHODIMP cValidateSession	( );	
};

typedef CApp *lpCApp;


#endif


