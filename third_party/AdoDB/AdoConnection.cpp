
#include "stdafx.h"
#include "AdoLib.h"
#include "AdoException.h"
#include "AdoConnection.h"

using namespace AdoLib;

// 积己磊
CAdoConnection::CAdoConnection()
: m_ConnectionPtr(NULL),
m_lTimeOut(0)
{
}

CAdoConnection::CAdoConnection(char* szConnection, long lTimeOut, bool bNowConnect)
: m_ConnectionPtr(NULL)
{
	SetConnectionInfo(szConnection, lTimeOut, bNowConnect);
}

CAdoConnection::CAdoConnection(char* strIP, 
							   UINT nPortNum, 
							   char* szDbName, 
							   char* szID, 
							   char* szPwd, 
							   long lTimeOut, 
							   bool bNowConnect)
: m_ConnectionPtr(NULL)
{
	SetConnectionInfo(strIP, nPortNum, szDbName, szID, szPwd, lTimeOut, bNowConnect);
}

CAdoConnection::~CAdoConnection()
{
	if(IsConnected()) { DisconnectDBMS(); }
}

// 目臣记 沥焊 汲沥
bool CAdoConnection::SetConnectionInfo(char* szConnection, long lTimeOut, bool bNowConnect)
{
	m_strConnection = szConnection;
	m_lTimeOut = lTimeOut;

	if(bNowConnect) { return ConnectDBMS(); }
	return true;
}

bool CAdoConnection::SetConnectionInfo(char* strIP,
									   UINT nPortNum, 
									   char* szDbName, 
									   char* szID, 
									   char* szPwd, 
									   long lTimeOut, 
									   bool bNowConnect)
{
	m_lTimeOut = lTimeOut;

	m_strConnection = "Provider=sqloledb; ";

	m_strConnection += "Data Source="; 
	m_strConnection += strIP;
	m_strConnection += ",";

	char strPort[16] = {0};
	_itoa_s(nPortNum, strPort, _countof(strPort), 10);
	m_strConnection += strPort;
	m_strConnection += "; ";

	m_strConnection += "Initial Catalog=";
	m_strConnection += szDbName;
	m_strConnection += "; ";
	
	m_strConnection += "User ID=";
	m_strConnection += szID;
	m_strConnection += "; ";
	
	m_strConnection += "Password=";
	m_strConnection += szPwd;
	m_strConnection += ";";

	if(bNowConnect) { return ConnectDBMS(); }
	return true;
}

// 叼厚 立加
bool CAdoConnection::ConnectDBMS()
{
	try 
	{
		if(IsConnected()) { return true; }

		if(NULL == m_ConnectionPtr)
		{
			m_ConnectionPtr.CreateInstance(__uuidof(ADODB::Connection));
			if(NULL == m_ConnectionPtr)
			{
				return false;
			}
		}

		if(0 != m_lTimeOut) m_ConnectionPtr->ConnectionTimeout = m_lTimeOut;
		if(0 != m_lTimeOut) m_ConnectionPtr->CommandTimeout = m_lTimeOut;
		m_ConnectionPtr->CursorLocation = ADODB::adUseClient;

		m_ConnectionPtr->Open(m_strConnection.c_str(), "", "", ADODB::adConnectUnspecified);
		return IsConnected();
	}
	catch(_com_error& e)
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoConnection::ConnectDBMS"));
		return false;
	}
}

// 叼厚 立加 秦力
bool CAdoConnection::DisconnectDBMS()
{
	try
	{
		if(NULL != m_ConnectionPtr)
		{
			m_ConnectionPtr->Close();
			m_ConnectionPtr.Release();
			m_ConnectionPtr = NULL;
		}
		return true;
	}
	catch(_com_error& e)
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoConnection::DisconnectDBMS"));
		return false;
	}
}

// 立加 咯何
bool CAdoConnection::IsConnected()
{
	if(NULL == m_ConnectionPtr || ADODB::adStateClosed == m_ConnectionPtr->State)
	{
		return false;
	}

	int a = 255;
	switch(m_ConnectionPtr->State)
	{
	case ADODB::adStateClosed:
		a = 0;
		break;
	case ADODB::adStateOpen:
		a = 1;
		break;
	case ADODB::adStateConnecting:
		a = 2;
		break;
	case ADODB::adStateExecuting:
		a = 4;
		break;
	case ADODB::adStateFetching:
		a = 18;
		break;
	}

	return true;
}

bool CAdoConnection::IsDisconnected()
{
	return !IsConnected();
}
