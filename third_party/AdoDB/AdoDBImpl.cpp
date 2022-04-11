
#include "stdafx.h"
#include "AdoLib.h"
#include "AdoException.h"
#include "AdodbImpl.h"

using namespace AdoLib;

// 생성자
CAdoDBImpl::CAdoDBImpl()
: m_fnExceptionLog(NULL),
m_RecordsetPtr(NULL),
m_lTimeOut(0)
{
}

CAdoDBImpl::~CAdoDBImpl()
{
	ReleaseRecordset();
}

// Release Recordset
bool CAdoDBImpl::ReleaseRecordset()
{
	try
	{
		if(NULL != m_RecordsetPtr) 
		{
			m_RecordsetPtr.Release();
			m_RecordsetPtr = NULL;
		}

		m_strLastError.clear();
		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// Set Callback function ExceptionLog
void CAdoDBImpl::SetCallbackExceptionLog(FN_CALLBACL_ADO_EXCEPTION_LOG fnExceptionLog)
{
	m_fnExceptionLog = fnExceptionLog;
}

// Set Connection Info - connection string
bool CAdoDBImpl::SetConnectionInfo(char* szConnection, long lTimeOut, bool bNowConnect)
{
	try
	{
		m_lTimeOut = lTimeOut;
		m_CAdoConnection.SetConnectionInfo(szConnection, lTimeOut, bNowConnect);

		m_strLastError.clear();
		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// Set Connection Info - data
bool CAdoDBImpl::SetConnectionInfo(char* strIP,
								   UINT nPortNum, 
								   char* szDbName, 
								   char* szID, 
								   char* szPwd, 
								   long lTimeOut,
								   bool bNowConnect)
{
	try
	{
		m_lTimeOut = lTimeOut;
		m_CAdoConnection.SetConnectionInfo(strIP, nPortNum, szDbName, szID, szPwd, lTimeOut, bNowConnect);

		m_strLastError.clear();
		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// Set Commond
bool CAdoDBImpl::SetCommond(std::string strQuery)
{
	try
	{
		if(!ReleaseRecordset())				{ return false; }
		if(!m_CAdoConnection.ConnectDBMS())	{ return false; }

		m_strLastError.clear();
		return m_CAdoCommand.SetCommand(m_CAdoConnection.GetConnectionPtr(), strQuery, m_lTimeOut);
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// Execute - not Record set
bool CAdoDBImpl::Execute()
{
	try 
	{
		m_strLastError.clear();
		return m_CAdoCommand.Execute();
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();

		m_CAdoConnection.DisconnectDBMS();

		return false;
	}
}

// Execute - return Record set
bool CAdoDBImpl::ExecuteEx()
{
	try 
	{
		m_strLastError.clear();
		return m_CAdoCommand.Execute(m_RecordsetPtr);
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		
		m_CAdoConnection.DisconnectDBMS();
		
		return false;
	}
}

//-----------------------------------------------------------------------------
// Set Input Output Bind Param - total : Exception 처리
//-----------------------------------------------------------------------------
bool CAdoDBImpl::BindParam(const char* szName, 
						   _variant_t& nValue, 
						   ADODB::DataTypeEnum eType, 
						   ADODB::ParameterDirectionEnum eDirection,
						   DWORD dwSize)
{
	try 
	{
		m_strLastError.clear();
		return m_CAdoCommand.BindParam(szName, nValue, eType, eDirection, dwSize);
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

///////////////////////////////////////////////////////////////////////////////
// Set Input Param
///////////////////////////////////////////////////////////////////////////////

// unsigned __int64
bool CAdoDBImpl::SetInputParam(const char* szName, unsigned __int64& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adBigInt, ADODB::adParamInput, sizeof(unsigned __int64));
}

// unsigned int
bool CAdoDBImpl::SetInputParam(const char* szName, unsigned int& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamInput, sizeof(unsigned int));
}

// unsigned long
bool CAdoDBImpl::SetInputParam(const char* szName, unsigned long& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamInput, sizeof(unsigned long));
}

// unsigned short
bool CAdoDBImpl::SetInputParam(const char* szName, unsigned short& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adSmallInt, ADODB::adParamInput, sizeof(unsigned short));
}

// __int64
bool CAdoDBImpl::SetInputParam(const char* szName, __int64& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adBigInt, ADODB::adParamInput, sizeof(__int64));
}

// int
bool CAdoDBImpl::SetInputParam(const char* szName, int& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamInput, sizeof(int));
}

// long
bool CAdoDBImpl::SetInputParam(const char* szName, long& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamInput, sizeof(long));
}

// short
bool CAdoDBImpl::SetInputParam(const char* szName, short& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adSmallInt, ADODB::adParamInput, sizeof(short));
}

// BYTE
bool CAdoDBImpl::SetInputParam(const char* szName, BYTE& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adTinyInt, ADODB::adParamInput, sizeof(BYTE));
}

// bool
bool CAdoDBImpl::SetInputParam(const char* szName, bool& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adBoolean, ADODB::adParamInput, sizeof(bool));
}

// double
bool CAdoDBImpl::SetInputParam(const char* szName, double& nValue)
{
	_variant_t varValue = nValue;
	return BindParam(szName, varValue, ADODB::adDouble, ADODB::adParamInput, sizeof(double));
}

// tm
bool CAdoDBImpl::SetInputParam(const char* szName, tm& nValue)
{
	SYSTEMTIME tmpValue = {0};

	tmpValue.wYear      	= nValue.tm_year + 1900;
	tmpValue.wMonth     	= nValue.tm_mon  + 1;
	tmpValue.wDayOfWeek 	= nValue.tm_wday;
	tmpValue.wDay       	= nValue.tm_mday;
	tmpValue.wHour      	= nValue.tm_hour;
	tmpValue.wMinute    	= nValue.tm_min;
	tmpValue.wSecond    	= nValue.tm_sec;
	tmpValue.wMilliseconds	= 0;

	return SetInputParam(szName, tmpValue);
}

// SYSTEMTIME
bool CAdoDBImpl::SetInputParam(const char* szName, SYSTEMTIME& nValue)
{
	_variant_t varValue;
	varValue.vt = VT_DATE;

	if(!SystemTimeToVariantTime(&nValue, &varValue.date)) { return false; }

	return BindParam(szName, varValue, ADODB::adDate, ADODB::adParamInput, sizeof(double));
}

// char
bool CAdoDBImpl::SetInputParam(const char* szName, char& nValue)
{
	std::string strValue;
	strValue = nValue;

	return SetInputParam(szName, strValue);
}

// WCHAR
bool CAdoDBImpl::SetInputParam(const char* szName, WCHAR& nValue)
{
	std::wstring strValue;
	strValue = nValue;

	return SetInputParam(szName, strValue);
}

// char*
bool CAdoDBImpl::SetInputParam(const char* szName, std::string nValue)
{
	_variant_t varValue = ConvertSafeSQLString(nValue).c_str();

	DWORD dwLength = (DWORD)nValue.length();
	if(nValue.empty())
	{
		varValue.vt = VT_NULL;
		dwLength		= 1;
	}

	return BindParam(szName, varValue, ADODB::adVarChar, ADODB::adParamInput, sizeof(char) * dwLength);
}

// WCHAR*
bool CAdoDBImpl::SetInputParam(const char* szName, std::wstring nValue)
{
	_variant_t varValue = ConvertSafeSQLString(nValue).c_str();

	DWORD dwLength = (DWORD)nValue.length();
	if(nValue.empty())
	{
		varValue.vt = VT_NULL;
		dwLength		= 1;
	}

	return BindParam(szName, varValue, ADODB::adVarWChar, ADODB::adParamInput, sizeof(WCHAR) * dwLength);
}

// Binary
bool CAdoDBImpl::SetInputParam(const char* szName, BYTE* nValue, DWORD dwSize)
{
	_variant_t varValue;

	if(!BytesToVariantArray(nValue, dwSize, &varValue)) { return false; }

	bool bResult = BindParam(szName, varValue, ADODB::adBinary, ADODB::adParamInput, dwSize);

	VariantClear(&varValue);

	return bResult;
}

// variant
bool CAdoDBImpl::SetInputParam(const char* szName, _variant_t& nValue, ADODB::DataTypeEnum eType, DWORD dwSize)
{
	return BindParam(szName, nValue, eType, ADODB::adParamInput, dwSize);
}

///////////////////////////////////////////////////////////////////////////////
// Set Output Param
///////////////////////////////////////////////////////////////////////////////

// unsigned __int64
bool CAdoDBImpl::SetOutputParam(const char* szName, unsigned __int64& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adBigInt, ADODB::adParamOutput, sizeof(unsigned __int64));
}

// unsigned int
bool CAdoDBImpl::SetOutputParam(const char* szName, unsigned int& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamOutput, sizeof(unsigned int));
}

// unsigned long
bool CAdoDBImpl::SetOutputParam(const char* szName, unsigned long& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamOutput, sizeof(unsigned long));
}

// unsigned short
bool CAdoDBImpl::SetOutputParam(const char* szName, unsigned short& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adSmallInt, ADODB::adParamOutput, sizeof(unsigned short));
}

// __int64
bool CAdoDBImpl::SetOutputParam(const char* szName, __int64& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adBigInt, ADODB::adParamOutput, sizeof(__int64));
}

// int 
bool CAdoDBImpl::SetOutputParam(const char* szName, int& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamOutput, sizeof(int));
}

// long 
bool CAdoDBImpl::SetOutputParam(const char* szName, long& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adInteger, ADODB::adParamOutput, sizeof(long));
}

// short
bool CAdoDBImpl::SetOutputParam(const char* szName, short& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adSmallInt, ADODB::adParamOutput, sizeof(short));
}

// BYTE
bool CAdoDBImpl::SetOutputParam(const char* szName, BYTE& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adTinyInt, ADODB::adParamOutput, sizeof(BYTE));
}

// bool
bool CAdoDBImpl::SetOutputParam(const char* szName, bool& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adBoolean, ADODB::adParamOutput, sizeof(bool));
}

// double
bool CAdoDBImpl::SetOutputParam(const char* szName, double& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adDouble, ADODB::adParamOutput, sizeof(double));
}

// tm
bool CAdoDBImpl::SetOutputParam(const char* szName, tm& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adDate, ADODB::adParamOutput, sizeof(double));
}

// SYSTEMTIME
bool CAdoDBImpl::SetOutputParam(const char* szName, SYSTEMTIME& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adDate, ADODB::adParamOutput, sizeof(double));
}

// char
bool CAdoDBImpl::SetOutputParam(const char* szName, char& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adChar, ADODB::adParamOutput, sizeof(char));
}

// WCHAR
bool CAdoDBImpl::SetOutputParam(const char* szName, WCHAR& nValue)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adWChar, ADODB::adParamOutput, sizeof(WCHAR));
}

// char*
bool CAdoDBImpl::SetOutputParam(const char* szName, std::string nValue, DWORD dwLength)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adVarChar, ADODB::adParamOutput, sizeof(char) * dwLength);
}

// WCHAR*
bool CAdoDBImpl::SetOutputParam(const char* szName, std::wstring nValue, DWORD dwLength)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adVarWChar, ADODB::adParamOutput, sizeof(WCHAR) * dwLength);
}

// Binary
bool CAdoDBImpl::SetOutputParam(const char* szName, BYTE* nValue, DWORD dwSize)
{
	_variant_t varValue;
	return BindParam(szName, varValue, ADODB::adBinary, ADODB::adParamOutput, dwSize);
}

// variant
bool CAdoDBImpl::SetOutputParam(const char* szName, ADODB::DataTypeEnum eType, DWORD dwSize)
{
	_variant_t varValue;
	return BindParam(szName, varValue, eType, ADODB::adParamOutput, dwSize);
}

// Set Param : in out param 설정
bool CAdoDBImpl::SetParam(const char* szName, 
						  _variant_t& nValue, 
						  ADODB::DataTypeEnum eType, 
						  ADODB::ParameterDirectionEnum eDirection, 
						  DWORD dwSize)
{
	return BindParam(szName, nValue, eType, eDirection, dwSize);
}

///////////////////////////////////////////////////////////////////////////////
// Get Param Value
///////////////////////////////////////////////////////////////////////////////

// unsigned __int64
bool CAdoDBImpl::GetParamValue(const char* szName, unsigned __int64& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned __int64)varValue;	break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// unsigned int
bool CAdoDBImpl::GetParamValue(const char* szName, unsigned int& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned int)varValue;		break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// unsigned long
bool CAdoDBImpl::GetParamValue(const char* szName, unsigned long& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned long)varValue;		break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// unsigned short
bool CAdoDBImpl::GetParamValue(const char* szName, unsigned short& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned short)varValue;		break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// __int64
bool CAdoDBImpl::GetParamValue(const char* szName, __int64& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (__int64)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// int
bool CAdoDBImpl::GetParamValue(const char* szName, int& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (int)varValue;					break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// long
bool CAdoDBImpl::GetParamValue(const char* szName, long& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (long)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// short
bool CAdoDBImpl::GetParamValue(const char* szName, short& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (short)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// BYTE
bool CAdoDBImpl::GetParamValue(const char* szName, BYTE& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (BYTE)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// bool
bool CAdoDBImpl::GetParamValue(const char* szName, bool& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// double
bool CAdoDBImpl::GetParamValue(const char* szName, double& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (double)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// tm
bool CAdoDBImpl::GetParamValue(const char* szName, tm& nValue)
{
	try 
	{
		m_strLastError.clear();

		ZeroMemory(&nValue, sizeof(tm));

		SYSTEMTIME tmpValue;
		if(!GetParamValue(szName, tmpValue)) return false;

		nValue.tm_year  = tmpValue.wYear - 1900;
		nValue.tm_mon   = tmpValue.wMonth - 1;
		nValue.tm_wday  = tmpValue.wDayOfWeek;
		nValue.tm_mday  = tmpValue.wDay;
		nValue.tm_yday  = (int)wce_getYdayFromSYSTEMTIME(&tmpValue);
		nValue.tm_hour  = tmpValue.wHour;
		nValue.tm_min   = tmpValue.wMinute;
		nValue.tm_sec   = tmpValue.wSecond;    
		nValue.tm_isdst = 0;
		
		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// SYSTEMTIME
bool CAdoDBImpl::GetParamValue(const char* szName, SYSTEMTIME& nValue)
{
	try 
	{
		m_strLastError.clear();

		ZeroMemory(&nValue, sizeof(SYSTEMTIME));

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		if(VT_EMPTY != varValue.vt && VT_NULL != varValue.vt)
		{
			if(!VariantTimeToSystemTime(varValue.date, &nValue))
			{
				ZeroMemory(&nValue, sizeof(SYSTEMTIME));
			}
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// char
bool CAdoDBImpl::GetParamValue(const char* szName, char& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);
		
		if(VT_BSTR == varValue.vt)
			nValue = (char)varValue.bstrVal[0];
		else
			nValue = 0;
		
		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// WCHAR
bool CAdoDBImpl::GetParamValue(const char* szName, WCHAR& nValue)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		if(VT_BSTR == varValue.vt)
			nValue = varValue.bstrVal[0];
		else
			nValue = 0;

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// char*
bool CAdoDBImpl::GetParamValue(const char* szName, char* nValue, DWORD dwLength)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		if(VT_BSTR == varValue.vt)
		{
			_bstr_t bstrVal(varValue.bstrVal);

			if(dwLength)
				strncpy_s(nValue, dwLength, bstrVal, dwLength);
			else
				strcpy(nValue, bstrVal);
		}
		else
		{
			nValue[0] = 0;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// WCHAR*
bool CAdoDBImpl::GetParamValue(const char* szName, WCHAR* nValue, DWORD dwLength)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		if(VT_BSTR == varValue.vt)
		{
			_bstr_t bstrVal(varValue.bstrVal);

			if(dwLength)
				wcsncpy_s(nValue, dwLength,bstrVal, dwLength);
			else
				wcscpy(nValue, bstrVal);
		}
		else
		{
			nValue[0] = 0;
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

// Binary
bool CAdoDBImpl::GetParamValue(const char* szName, BYTE* nValue, DWORD dwSize)
{
	try 
	{
		m_strLastError.clear();

		_variant_t varValue;
		m_CAdoCommand.GetParamValue(szName, varValue);

		DWORD Size = dwSize;
		if(!VariantArrayToBytes(varValue, nValue, &Size))
		{
			ZeroMemory(nValue, dwSize);
		}

		return true;
	}
	catch(CAdoException& e)
	{
		if(m_fnExceptionLog) m_fnExceptionLog(e.toString().c_str());
		m_strLastError = e.toString().c_str();
		return false;
	}
}

///////////////////////////////////////////////////////////////////////////////
// Get Field Value
///////////////////////////////////////////////////////////////////////////////

bool CAdoDBImpl::GetFirstFieldValue(int& nValue)
{
	try
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr)				{ return false; }
		if(1 > GetRecordCount())				{ return false; }
		if(1 > m_RecordsetPtr->Fields->Count)	{ return false; }
	
		_variant_t index = (short)0;

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(&index)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (int)varValue;	break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetReturnValue (int)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// unsigned __int64
bool CAdoDBImpl::GetFieldValue(const char* szName, unsigned __int64& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned __int64)varValue;	break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (unsigned __int64)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// unsigned int
bool CAdoDBImpl::GetFieldValue(const char* szName, unsigned int& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned int)varValue;		break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (unsigned int)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// unsigned long
bool CAdoDBImpl::GetFieldValue(const char* szName, unsigned long& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned long)varValue;		break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (unsigned long)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// unsigned short
bool CAdoDBImpl::GetFieldValue(const char* szName, unsigned short& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (unsigned short)varValue;		break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (unsigned short)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// __int64
bool CAdoDBImpl::GetFieldValue(const char* szName, __int64& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (__int64)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (__int64)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// int
bool CAdoDBImpl::GetFieldValue(const char* szName, int& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (int)varValue;					break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (int)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// long
bool CAdoDBImpl::GetFieldValue(const char* szName, long& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;		
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (long)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (long)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// short
bool CAdoDBImpl::GetFieldValue(const char* szName, short& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (short)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (short)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// BYTE
bool CAdoDBImpl::GetFieldValue(const char* szName, BYTE& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (BYTE)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (BYTE)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// bool
bool CAdoDBImpl::GetFieldValue(const char* szName, bool& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = 0;								break;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (bool)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// double
bool CAdoDBImpl::GetFieldValue(const char* szName, double& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();

		switch(varValue.vt)
		{
		case VT_I2:
		case VT_I4:
		case VT_R4:
		case VT_R8:
		case VT_CY:
		case VT_DECIMAL:
		case VT_I1:
		case VT_UI1:
		case VT_UI2:
		case VT_UI4:
		case VT_I8:
		case VT_UI8:
		case VT_INT:
		case VT_UINT:	nValue = (double)varValue;				break;
		case VT_BOOL:	nValue = ((bool)varValue) ? 1 : 0;		break;
		default :		nValue = -1;							break;
		}

		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (double)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// tm
bool CAdoDBImpl::GetFieldValue(const char* szName, tm& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		ZeroMemory(&nValue, sizeof(tm));

		SYSTEMTIME tmpValue;
		if(!GetFieldValue(szName, tmpValue)) return false;

		nValue.tm_year  = tmpValue.wYear - 1900;
		nValue.tm_mon   = tmpValue.wMonth - 1;
		nValue.tm_wday  = tmpValue.wDayOfWeek;
		nValue.tm_mday  = tmpValue.wDay;
		nValue.tm_yday  = (int)wce_getYdayFromSYSTEMTIME(&tmpValue);
		nValue.tm_hour  = tmpValue.wHour;
		nValue.tm_min   = tmpValue.wMinute;
		nValue.tm_sec   = tmpValue.wSecond;    
		nValue.tm_isdst = 0;

		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (tm)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// SYSTEMTIME
bool CAdoDBImpl::GetFieldValue(const char* szName, SYSTEMTIME& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		ZeroMemory(&nValue, sizeof(SYSTEMTIME));

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();

		if(VT_EMPTY != varValue.vt && VT_NULL != varValue.vt)
		{
			if(!VariantTimeToSystemTime(varValue.date, &nValue))
			{
				ZeroMemory(&nValue, sizeof(SYSTEMTIME));
			}
		}

		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (SYSTEMTIME)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// char
bool CAdoDBImpl::GetFieldValue(const char* szName, char& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		if(VT_BSTR == varValue.vt)
			nValue = (char)varValue.bstrVal[0];
		else
			nValue = 0;
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (char)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// WCHAR
bool CAdoDBImpl::GetFieldValue(const char* szName, WCHAR& nValue)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		if(VT_BSTR == varValue.vt)
			nValue = varValue.bstrVal[0];
		else
			nValue = 0;
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (WCHAR)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// char*
bool CAdoDBImpl::GetFieldValue(const char* szName, char* nValue, DWORD dwLength)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		if(VT_BSTR == varValue.vt)
		{
			_bstr_t bstrVal(varValue.bstrVal);

			if(dwLength)
				strncpy_s(nValue, dwLength, bstrVal, dwLength);
			else
				strcpy(nValue, bstrVal);
		}
		else
		{
			nValue[0] = 0;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (char*)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// WCHAR*
bool CAdoDBImpl::GetFieldValue(const char* szName, WCHAR* nValue, DWORD dwLength)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }
	
		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();
		
		if(VT_BSTR == varValue.vt)
		{
			_bstr_t bstrVal(varValue.bstrVal);

			if(dwLength)
				wcsncpy_s(nValue, dwLength, bstrVal, dwLength);
			else
				wcscpy(nValue, bstrVal);
		}
		else
		{
			nValue[0] = 0;
		}
		
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (WCHAR*)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

// Binary
bool CAdoDBImpl::GetFieldValue(const char* szName, BYTE* nValue, DWORD dwSize)
{
	try 
	{
		m_strLastError.clear();
		if(NULL == m_RecordsetPtr) { return false; }

		_variant_t varValue;
		varValue = m_RecordsetPtr->Fields->GetItem(szName)->GetValue();

		DWORD Size = dwSize;
		if(!VariantArrayToBytes(varValue, nValue, &Size))
		{
			ZeroMemory(nValue, dwSize);
		}

		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetFieldValue (Binary)"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		m_strLastError = ex.toString().c_str();
		return false;
	}
}

//---------------------------------------------------------------------------//
// Move Recode
//---------------------------------------------------------------------------//
int	CAdoDBImpl::GetRecordCount()
{
	try
	{
		if(NULL == m_RecordsetPtr) { return 0; }
		return m_RecordsetPtr->RecordCount;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::GetRecordCount"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		return false;
	}
}

bool CAdoDBImpl::FirstRecord()
{
	try
	{
		if(NULL == m_RecordsetPtr) { return false; }
		m_RecordsetPtr->MoveFirst();
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::FirstRecord"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		return false;
	}
}

bool CAdoDBImpl::NextRecord()
{
	try
	{
		if(NULL == m_RecordsetPtr) { return false; }
		m_RecordsetPtr->MoveNext();
		if(m_RecordsetPtr->EndOfFile) { return false; }
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::NextRecord"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		return false;
	}
}

bool CAdoDBImpl::NextRecordSet()
{
	try
	{
		if(NULL == m_RecordsetPtr) { return false; }
		m_RecordsetPtr = m_RecordsetPtr->NextRecordset(NULL);
		return true;
	}
	catch(_com_error& e)
	{
		CAdoException ex(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoDBImpl::NextRecordSet"));
		if(m_fnExceptionLog) m_fnExceptionLog(ex.toString().c_str());
		return false;
	}
}

//---------------------------------------------------------------------------//
// time utility
//---------------------------------------------------------------------------//
FILETIME CAdoDBImpl::wce_getFILETIMEFromYear(WORD year)
{
	SYSTEMTIME s = {0};
	FILETIME f;

	s.wYear      = year;
	s.wMonth     = 1;
	s.wDayOfWeek = 1;
	s.wDay       = 1;

	SystemTimeToFileTime( &s, &f );
	return f;
}

// __int64 <--> FILETIME
__int64 CAdoDBImpl::wce_FILETIME2int64(FILETIME f)
{
	__int64 t;

	t = f.dwHighDateTime;
	t <<= 32;
	t |= f.dwLowDateTime;
	return t;
}

time_t CAdoDBImpl::wce_getYdayFromSYSTEMTIME(const SYSTEMTIME* s)
{
	__int64 t;
	FILETIME f1, f2;

	f1 = wce_getFILETIMEFromYear(s->wYear);
	SystemTimeToFileTime(s, &f2);

	t = wce_FILETIME2int64(f2)-wce_FILETIME2int64(f1);

	return (time_t)((t/10000000)/(60*60*24));
}

//---------------------------------------------------------------------------//
// VariantArray <-> Bytes
//---------------------------------------------------------------------------//

// VariantArray -> Bytes
bool CAdoDBImpl::VariantArrayToBytes(VARIANT Variant,	// in
									 BYTE *pBytes,		// out
									 DWORD *pdwBytes)	// size
{
	if(!(Variant.vt & VT_ARRAY) ||
	!Variant.parray ||
	!pBytes ||
	!pdwBytes)
	{
		return false;
	}

	SAFEARRAY *pArrayVal = NULL;
	CHAR HUGEP *pArray = NULL;

	pArrayVal = Variant.parray;
	DWORD dwBytes = pArrayVal->rgsabound[0].cElements;

	if(*pdwBytes < dwBytes) { return false; }
	*pdwBytes = 0;

	if(!SUCCEEDED(SafeArrayAccessData(pArrayVal, (void HUGEP * FAR *) &pArray)))
	{
		return false;
	}

	memcpy(pBytes, pArray, dwBytes);
	SafeArrayUnaccessData(pArrayVal);
	*pdwBytes = dwBytes;

	return true;
}

// Bytes -> VariantArray
bool CAdoDBImpl::BytesToVariantArray(BYTE *pValue,			// in
									 DWORD cValueElements,	// size
									 VARIANT *pVariant)		// out - 사용 후 VariantClear(&var); 해야한다!!
{
	SAFEARRAY *pArrayVal = NULL;
	SAFEARRAYBOUND arrayBound;
	CHAR HUGEP *pArray = NULL;

	arrayBound.lLbound = 0;
	arrayBound.cElements = cValueElements;

	pArrayVal = SafeArrayCreate(VT_UI1, 1, &arrayBound);
	if (pArrayVal == NULL) { return false; }

	if (!SUCCEEDED(SafeArrayAccessData(pArrayVal, (void HUGEP* FAR*)&pArray)))
	{
		if (pArrayVal) { SafeArrayDestroy(pArrayVal); }
		return false;
	}

	memcpy( pArray, pValue, arrayBound.cElements );
	SafeArrayUnaccessData( pArrayVal );

	V_VT(pVariant) = VT_ARRAY | VT_UI1;
	V_ARRAY(pVariant) = pArrayVal;

	return true;
}

//---------------------------------------------------------------------------//
// Convert Safe SQL String
//---------------------------------------------------------------------------//

std::wstring CAdoDBImpl::ConvertSafeSQLString(std::wstring str)
{
	std::wstring strSafe = str.c_str();

	std::wstring::size_type Index = 0;

	while(true)
	{
		Index = strSafe.find(L"'", Index);
		if(std::wstring::npos == Index) break;
		strSafe.replace(Index, 1, L"''");
		Index += 2;
	}	

	return strSafe;
}

std::string CAdoDBImpl::ConvertSafeSQLString(std::string str)
{
	std::string strSafe = str.c_str();

	std::string::size_type Index = 0;

	while(true)
	{
		Index = strSafe.find("'", Index);
		if(std::wstring::npos == Index) break;
		strSafe.replace(Index, 1, "''");
		Index += 2;
	}	

	return strSafe;
}

tstring CAdoDBImpl::GetAdoLastError()
{
	return m_strLastError;
}
