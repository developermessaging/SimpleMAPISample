#pragma once

#ifndef _SWPMAIN_
#define _SWPMAIN_

#include "swap.h"
#include <iostream>
#include <string>

// Declare Menu constants
#define LOGON					1
#define SELECT_RECIPIENT		2
#define ENTER_RECIPIENT			3
#define GET_DETAILS				4
#define SEND_NO_UI				5
#define SEND_UI					6
#define SEND_ATTACH				7
#define CREATE_MSG				8
#define LIST_INBOX				9
#define READ_MAIL				10
#define LOGOFF					11
#define EXIT				    12
#define REFRESH					13

void main(int argc, char *argv[], char *envp[]);
void PrintMenuToConsole(void);

#endif


