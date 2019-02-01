// smplmapi.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
#include "smplmapi.h"



lpCApp pCApp = NULL;	// Global application object

/*
+---------------------------------------------------------------------
|
|	Function:	main()
|
|	 Purpose:	This is the entry point for application. It is also
|				the main control loop for tha app. Any command the
|				user enters is processed by this loop.
|
+---------------------------------------------------------------------
*/

void main(int argc, char *argv[], char *envp[])
{
	HRESULT hRes = S_OK;
	int		lpMenuChoice;   // Selection made by the user
	int		nNotDone = 0;	// Determined when to quit control loop.

	pCApp = new (CApp);

	//	Initialize application object

	if (SUCCESS_SUCCESS != (hRes = pCApp->cInitApp()))
		goto Quit;

	//  Display main menu to user.
	PrintMenuToConsole();

	// 	Command loop
	do
	{
		ULONG	cRecips = 0L;
		lpMapiRecipDesc Recips = NULL;

		printf("\r\nEnter your choice: ");
		scanf("%d", &lpMenuChoice);

		switch (lpMenuChoice)
		{
		case LOGON:
			hRes = pCApp->cLogon();
			break;
		case SELECT_RECIPIENT:
		{
			hRes = pCApp->cAddress(&cRecips, &Recips);

			if (SUCCESS_SUCCESS == hRes)
				printf(Recips->lpszAddress);
		}
		break;
		case GET_DETAILS:
		{
			LPSTR lpszName = NULL;
			lpMapiRecipDesc Recips = NULL;

			pCApp->cCaptureText(LPSTR(L"\r\nEnter an e-mail address to resolve: "), &lpszName);

			if (SUCCESS_SUCCESS == (hRes = pCApp->cResolveName(lpszName, &Recips)))
				hRes = pCApp->cGetDetails(Recips);

			pCApp->cFreeBuffer(lpszName);
			pCApp->cFreeBuffer(Recips);
		}break;
		case ENTER_RECIPIENT:
		{
			if (SUCCESS_SUCCESS == (hRes = pCApp->cValidateSession()))
			{
				char lpszName[256];
				lpMapiRecipDesc lpRecips = NULL;
				printf("Enter an e-mail address: ");
				scanf("%s", &lpszName);
				hRes = pCApp->cResolveName(lpszName, &lpRecips);
			}
			else
			{
				printf("Not logged on to messaging system.\r\n");
			}
		}
		break;
		case SEND_NO_UI:
			hRes = pCApp->cSendMessage(0L);
			break;
		case SEND_UI:
			hRes = pCApp->cSendMessage(MAPI_DIALOG);
			break;
		case SEND_ATTACH:
			hRes = pCApp->cSendAttachMail();
			break;
		case CREATE_MSG:
		{
			lpMapiMessage lpMessage = NULL;
			LPTSTR lpszEID = NULL;
			hRes = pCApp->cCreateMessage(0L, &lpMessage, &lpszEID);

			if (SUCCESS_SUCCESS == hRes)
			{
				pCApp->cFreeBuffer(lpszEID);
				pCApp->cFreeBuffer(lpMessage);
			}
		}break;
		case LIST_INBOX:
			pCApp->cListInboxMessages();
			break;
		case READ_MAIL:
			hRes = pCApp->cReadMail(0L, NULL);
			break;
		case LOGOFF:
			hRes = pCApp->cLogoff();
			PrintMenuToConsole();
			break;
		case EXIT:
			hRes = pCApp->cLogoff();
			nNotDone = 1;
			break;
		case REFRESH:
			PrintMenuToConsole();
			break;
		default:
			printf("Not a valid choice. Please try again.\r\n");
			break;
		}
		if (hRes == E_NOTIMPL)
			printf("Not yet implemented.\r\n");

		pCApp->cFreeBuffer(Recips);
	} while (!nNotDone);

Quit:

	delete pCApp;
}


/*
+---------------------------------------------------------------------
|
|	Function:	PrintMenuToConsole()
|
|	 Purpose:	Prints menu to the console so the user knows
|				what choices are available.
|
+---------------------------------------------------------------------
*/
void PrintMenuToConsole(void)
{
	printf("\r\nConsole Messaging Client Main Menu.\r\n\r\n");
	printf("Please choose one of the following:\r\n\r\n");
	printf("[ 1] Logon to messaging system.\r\n");
	printf("[ 2] Select recipient/Get e-mail address.\r\n");
	printf("[ 3] Enter email address/Get recipient.\r\n");
	printf("[ 4] Get details of recipient.\r\n");
	printf("[ 5] Send Mail message to recipient with no UI.\r\n");
	printf("[ 6] Send Mail message to recipient with UI.\r\n");
	printf("[ 7] Send Mail with attachment.\r\n");
	printf("[ 8] Create new message. \r\n");
	printf("[ 9] List messages in Inbox.\r\n");
	printf("[10] Read next unread message.\r\n");
	printf("[11] Logoff the message system.\r\n");
	printf("[12] Exit Client.\r\n");
	printf("[13] Refresh Menu.\r\n");
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
