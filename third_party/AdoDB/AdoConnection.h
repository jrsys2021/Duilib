
#pragma once

namespace AdoLib
{
	class CAdoConnection
	{
	public:			
		CAdoConnection();
		CAdoConnection(char* szConnection, 
					   long lTimeOut = 0,
					   bool bNowConnect = true);
		CAdoConnection(char* strIP, 
					   UINT nPortNum, 
					   char* szDbName, 
					   char* szID, 
					   char* szPwd, 
					   long lTimeOut = 0,
					   bool bNowConnect = true);
		~CAdoConnection();
		
		bool	SetConnectionInfo(char* szConnection, 
								  long lTimeOut = 0,
								  bool bNowConnect = true);
		bool	SetConnectionInfo(char* strIP, 
								  UINT nPortNum, 
								  char* szDbName, 
								  char* szID, 
								  char* szPwd, 
								  long lTimeOut = 0,
								  bool bNowConnect = true);

		bool	ConnectDBMS();
		bool	DisconnectDBMS();

		bool	IsConnected();
		bool	IsDisconnected();

		ADODB::_ConnectionPtr GetConnectionPtr(){ return m_ConnectionPtr; }
		
	private:		
		ADODB::_ConnectionPtr	m_ConnectionPtr;
		std::string				m_strConnection;
		long					m_lTimeOut;
	};
};
