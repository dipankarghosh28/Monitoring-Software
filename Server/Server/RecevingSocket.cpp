// RecevingSocket.cpp : implementation file
//

#include "stdafx.h"
#include "Server.h"
#include "RecevingSocket.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CRecevingSocket

CRecevingSocket::CRecevingSocket()
{
}

CRecevingSocket::~CRecevingSocket()
{
}


// Do not edit the following lines, which are needed by ClassWizard.
#if 0
BEGIN_MESSAGE_MAP(CRecevingSocket, CSocket)
	//{{AFX_MSG_MAP(CRecevingSocket)
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()
#endif	// 0

/////////////////////////////////////////////////////////////////////////////
// CRecevingSocket member functions

void CRecevingSocket::OnReceive(int nErrorCode) 
{
	char strRec[256]="";
	Receive(strRec,256);

	((CServerApp*)AfxGetApp())->m_pServerView->AddMsg(strRec);
	CSocket::OnReceive(nErrorCode);
}
