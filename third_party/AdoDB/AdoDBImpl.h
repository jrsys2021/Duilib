
#pragma once

#include "AdoConnection.h"
#include "AdoCommand.h"

namespace AdoLib
{
	typedef void (*FN_CALLBACL_ADO_EXCEPTION_LOG)(const TCHAR*);

	class CAdoDBImpl
	{	
		friend class CAdoConnection;
		friend class CAdoCommand;

	public:			
		CAdoDBImpl();
		virtual ~CAdoDBImpl();

		void SetCallbackExceptionLog(FN_CALLBACL_ADO_EXCEPTION_LOG fnExceptionLog);

		bool SetConnectionInfo(char* szConnection, 
							   long lTimeOut = 0,
							   bool bNowConnect = true);
		bool SetConnectionInfo(char* strIP, 
							   UINT nPortNum, 
							   char* szDbName, 
							   char* szID, 
							   char* szPwd, 
							   long lTimeOut = 0,
							   bool bNowConnect = true);

		bool SetCommond(std::string strQuery);
		bool Execute();
		bool ExecuteEx();

		bool ReleaseRecordset();
	
		bool SetInputParam(const char* szName, unsigned __int64& nValue);
		bool SetInputParam(const char* szName, unsigned int& nValue);
		bool SetInputParam(const char* szName, unsigned long& nValue);
		bool SetInputParam(const char* szName, unsigned short& nValue);
		bool SetInputParam(const char* szName, __int64& nValue);
		bool SetInputParam(const char* szName, int& nValue);
		bool SetInputParam(const char* szName, long& nValue);
		bool SetInputParam(const char* szName, short& nValue);
		bool SetInputParam(const char* szName, BYTE& nValue);
		bool SetInputParam(const char* szName, bool& nValue);
		bool SetInputParam(const char* szName, double& nValue);
		bool SetInputParam(const char* szName, tm& nValue);
		bool SetInputParam(const char* szName, SYSTEMTIME& nValue);
		bool SetInputParam(const char* szName, char& nValue);
		bool SetInputParam(const char* szName, WCHAR& nValue);
		bool SetInputParam(const char* szName, std::string nValue);
		bool SetInputParam(const char* szName, std::wstring nValue);
		bool SetInputParam(const char* szName, BYTE* nValue, DWORD dwSize);
		bool SetInputParam(const char* szName, _variant_t& nValue, ADODB::DataTypeEnum eType, DWORD dwSize);
		
		bool SetOutputParam(const char* szName, unsigned __int64& nValue);
		bool SetOutputParam(const char* szName, unsigned int& nValue);
		bool SetOutputParam(const char* szName, unsigned long& nValue);
		bool SetOutputParam(const char* szName, unsigned short& nValue);
		bool SetOutputParam(const char* szName, __int64& nValue);
		bool SetOutputParam(const char* szName, int& nValue);
		bool SetOutputParam(const char* szName, long& nValue);
		bool SetOutputParam(const char* szName, short& nValue);
		bool SetOutputParam(const char* szName, BYTE& nValue);
		bool SetOutputParam(const char* szName, bool& nValue);
		bool SetOutputParam(const char* szName, double& nValue);
		bool SetOutputParam(const char* szName, tm& nValue);
		bool SetOutputParam(const char* szName, SYSTEMTIME& nValue);
		bool SetOutputParam(const char* szName, char& nValue);
		bool SetOutputParam(const char* szName, WCHAR& nValue);
		bool SetOutputParam(const char* szName, std::string nValue, DWORD dwLength);
		bool SetOutputParam(const char* szName, std::wstring nValue, DWORD dwLength);
		bool SetOutputParam(const char* szName, BYTE* nValue, DWORD dwSize);
		bool SetOutputParam(const char* szName, ADODB::DataTypeEnum eType, DWORD dwSize);

		bool SetParam(const char* szName, _variant_t& nValue, ADODB::DataTypeEnum eType, ADODB::ParameterDirectionEnum eDirection, DWORD dwSize);

		bool GetParamValue(const char* szName, unsigned __int64& nValue);
		bool GetParamValue(const char* szName, unsigned int& nValue);
		bool GetParamValue(const char* szName, unsigned long& nValue);
		bool GetParamValue(const char* szName, unsigned short& nValue);
		bool GetParamValue(const char* szName, __int64& nValue);
		bool GetParamValue(const char* szName, int& nValue);
		bool GetParamValue(const char* szName, long& nValue);
		bool GetParamValue(const char* szName, short& nValue);
		bool GetParamValue(const char* szName, BYTE& nValue);
		bool GetParamValue(const char* szName, bool& nValue);
		bool GetParamValue(const char* szName, double& nValue);
		bool GetParamValue(const char* szName, tm& nValue);
		bool GetParamValue(const char* szName, SYSTEMTIME& nValue);
		bool GetParamValue(const char* szName, char& nValue);
		bool GetParamValue(const char* szName, WCHAR& nValue);
		bool GetParamValue(const char* szName, char* nValue, DWORD dwLength = 0);
		bool GetParamValue(const char* szName, WCHAR* nValue, DWORD dwLength = 0);
		bool GetParamValue(const char* szName, BYTE* nValue, DWORD dwSize);

		bool GetFirstFieldValue(int& nValue);
		bool GetFieldValue(const char* szName, unsigned __int64& nValue);
		bool GetFieldValue(const char* szName, unsigned int& nValue);
		bool GetFieldValue(const char* szName, unsigned long& nValue);
		bool GetFieldValue(const char* szName, unsigned short& nValue);
		bool GetFieldValue(const char* szName, __int64& nValue);
		bool GetFieldValue(const char* szName, int& nValue);
		bool GetFieldValue(const char* szName, long& nValue);
		bool GetFieldValue(const char* szName, short& nValue);
		bool GetFieldValue(const char* szName, BYTE& nValue);
		bool GetFieldValue(const char* szName, bool& nValue);
		bool GetFieldValue(const char* szName, double& nValue);
		bool GetFieldValue(const char* szName, tm& nValue);
		bool GetFieldValue(const char* szName, SYSTEMTIME& nValue);
		bool GetFieldValue(const char* szName, char& nValue);
		bool GetFieldValue(const char* szName, WCHAR& nValue);
		bool GetFieldValue(const char* szName, char* nValue, DWORD dwLength = 0);
		bool GetFieldValue(const char* szName, WCHAR* nValue, DWORD dwLength = 0);
		bool GetFieldValue(const char* szName, BYTE* nValue, DWORD dwSize);

		int	 GetRecordCount();
		bool FirstRecord();
		bool NextRecord();
		bool NextRecordSet();

		tstring GetAdoLastError();

		bool IsConnect() { return m_CAdoConnection.IsConnected(); }
	private:
		bool BindParam(const char* szName, 
					   _variant_t& nValue, 
					   ADODB::DataTypeEnum eType, 
					   ADODB::ParameterDirectionEnum eDirection,
					   DWORD dwSize);

		// string
		std::string ConvertSafeSQLString(std::string str);
		std::wstring ConvertSafeSQLString(std::wstring str);

		// time utility
		FILETIME wce_getFILETIMEFromYear(WORD year);
		__int64 wce_FILETIME2int64(FILETIME f);
		time_t wce_getYdayFromSYSTEMTIME(const SYSTEMTIME* s);

		// VariantArray <-> Bytes
		bool VariantArrayToBytes(VARIANT Variant,
			 					 BYTE *ppBytes,
			 					 DWORD *pdwBytes);
		bool BytesToVariantArray(BYTE *pValue,
			 					 DWORD cValueElements,
			 					 VARIANT *pVariant);


		ADODB::_RecordsetPtr			m_RecordsetPtr;
		CAdoConnection					m_CAdoConnection;
		CAdoCommand						m_CAdoCommand;
		long							m_lTimeOut;
		FN_CALLBACL_ADO_EXCEPTION_LOG	m_fnExceptionLog;
		tstring							m_strLastError;
	};
};
