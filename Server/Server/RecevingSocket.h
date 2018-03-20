#if !defined(AFX_RECEVINGSOCKET_H__E6B5E8B5_C06F_4A31_A677_173F37F8FB32__INCLUDED_)
#define AFX_RECEVINGSOCKET_H__E6B5E8B5_C06F_4A31_A677_173F37F8FB32__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
// RecevingSocket.h : header file
//



/////////////////////////////////////////////////////////////////////////////
// CRecevingSocket command target

class CRecevingSocket : public CSocket
{
// Attributes
public:

// Operations
public:
	CRecevingSocket();
	virtual ~CRecevingSocket();

// Overrides
public:
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CRecevingSocket)
	public:
	virtual void OnReceive(int nErrorCode);
	//}}AFX_VIRTUAL

	// Generated message map functions
	//{{AFX_MSG(CRecevingSocket)
		// NOTE - the ClassWizard will add and remove member functions here.
	//}}AFX_MSG

// Implementation
protected:
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_RECEVINGSOCKET_H__E6B5E8B5_C06F_4A31_A677_173F37F8FB32__INCLUDED_)
