
#pragma once

namespace AdoLib
{
	class CAdoCommand
	{
	public:			
		CAdoCommand();
		~CAdoCommand();

		bool SetCommand(ADODB::_ConnectionPtr ConnectionPtr, std::string strQuery, long lTimeOut);
		bool ReleaseCommand();
		
		bool Execute();
		bool Execute(ADODB::_RecordsetPtr& RecordsetPtr);

		bool GetParamValue(const char* szName, _variant_t&  nValue);
		bool BindParam(const char* szName, 
					   _variant_t& nValue, 
					   ADODB::DataTypeEnum eType, 
					   ADODB::ParameterDirectionEnum eDirection,
					   long nSize);
		
	private:
		std::string SetParamName(const char* szName);

		ADODB::_CommandPtr		m_CommandPtr;
		ADODB::CommandTypeEnum	m_CommandType;
	};
};
