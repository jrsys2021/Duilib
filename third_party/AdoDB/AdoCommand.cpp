
#include "stdafx.h"
#include "AdoLib.h"
#include "AdoException.h"
#include "AdoCommand.h"

using namespace AdoLib;

// »ý¼ºÀÚ
CAdoCommand::CAdoCommand()
: m_CommandPtr(NULL),
m_CommandType(ADODB::adCmdUnspecified)
{
}

CAdoCommand::~CAdoCommand()
{
	ReleaseCommand();
}

// Set Command
bool CAdoCommand::SetCommand(ADODB::_ConnectionPtr ConnectionPtr, std::string strQuery, long lTimeOut)
{
	try
	{
		if(!ReleaseCommand()) { return false; }
		m_CommandType = ADODB::adCmdUnspecified;

		if(NULL == m_CommandPtr)
		{
			m_CommandPtr.CreateInstance(__uuidof(ADODB::Command));
		}

		if(0 != lTimeOut) m_CommandPtr->CommandTimeout = lTimeOut;
		m_CommandPtr->ActiveConnection = ConnectionPtr;
		m_CommandPtr->CommandText = strQuery.c_str();

		return true;
	}
	catch(_com_error& e)
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoCommand::SetCommand"));
		return false;
	}
}

// Release Command
bool CAdoCommand::ReleaseCommand()
{
	try
	{
		if(NULL != m_CommandPtr) 
		{
			m_CommandPtr.Release();
			m_CommandPtr = NULL;
		}
		return true;
	}
	catch(_com_error& e)
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoCommand::ReleaseCommand"));
		return false;
	}
}

// Execute
bool CAdoCommand::Execute()
{
	try 
	{
		m_CommandPtr->Execute(NULL, NULL, m_CommandType);
		return true;
	}
	catch( _com_error& e ) 
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoCommand::Execute (1)"));
		return false;
	}
}

bool CAdoCommand::Execute(ADODB::_RecordsetPtr& RecordsetPtr)
{
	try 
	{
		RecordsetPtr = m_CommandPtr->Execute(NULL, NULL, m_CommandType);
		return true;
	}
	catch( _com_error& e ) 
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoCommand::Execute (2)"));
		return false;
	}
}

// Bind Param
bool CAdoCommand::BindParam(const char* szName, 
							_variant_t& nValue, 
							ADODB::DataTypeEnum eType, 
							ADODB::ParameterDirectionEnum eDirection,
							long nSize)
{
	if(ADODB::adCmdUnspecified == m_CommandType)
		m_CommandType = ADODB::adCmdStoredProc;

	try
	{
		m_CommandPtr->Parameters->Append(
			m_CommandPtr->CreateParameter(SetParamName(szName).c_str(), eType, eDirection, nSize, nValue));
		return true;
	}
	catch(_com_error& e)
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoCommand::BindParam"));
		return false;
	}
}

// Get Param Value
bool CAdoCommand::GetParamValue(const char* szName, _variant_t&  nValue)
{
	try
	{
		nValue = m_CommandPtr->Parameters->GetItem(SetParamName(szName).c_str())->Value;
		//if(VT_EMPTY == nValue.vt || VT_NULL == nValue.vt) return false;
		return true;
	}
	catch(_com_error& e)
	{
		throw CAdoException(e.ErrorMessage(), e.Error(), (TCHAR*)e.Description(), (TCHAR*)e.Source(), _TEXT("CAdoCommand::GetParamValue"));
		return false;
	}
}

// Set Param Name
std::string CAdoCommand::SetParamName(const char* szName)
{
	std::string strName;
	if('@' != szName[0]) { strName = "@"; }
	strName += szName;

	return strName;
}
