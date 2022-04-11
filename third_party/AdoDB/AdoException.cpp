
#include "stdafx.h"
#include "AdoException.h"

using namespace AdoLib;

CAdoException::CAdoException()
: m_warnCode(0)
{
}

CAdoException::CAdoException(tstring message, int errorCode, tstring description)
: m_warnCode(errorCode),
m_message(message),
m_description(description)
{
}

CAdoException::CAdoException(tstring message, int errorCode, tstring description, tstring source, tstring location)
: m_warnCode(errorCode),
m_message(message),
m_description(description),
m_source(source),
m_Location(location)
{
}

CAdoException::~CAdoException()
{
}

int CAdoException::getErrorCode()
{
	return m_warnCode;
}

tstring	CAdoException::getMessage()
{
	return m_message;
}

tstring CAdoException::getDescription()
{
	return m_description;
}

tstring CAdoException::getSource()
{
	return m_source;
}

tstring CAdoException::getLocation()
{
	return m_Location;
}

tstring CAdoException::toString()
{
	TCHAR szErrorCode[16];
	swprintf_s(szErrorCode, _TEXT("%d"), getErrorCode());

	tstring buffer;
	buffer.append(_TEXT("\t[CAdoException]\r\n"));
	buffer.append(_TEXT("\t[code]")		); buffer.append(szErrorCode 		);buffer.append(_TEXT("\r\n"));
	buffer.append(_TEXT("\t[message]")		); buffer.append(getMessage()		);buffer.append(_TEXT("\r\n"));
	buffer.append(_TEXT("\t[description]")	); buffer.append(getDescription()	);buffer.append(_TEXT("\r\n"));
	buffer.append(_TEXT("\t[souce]")		); buffer.append(getSource()		);buffer.append(_TEXT("\r\n"));
	buffer.append(_TEXT("\t[location]")	); buffer.append(getLocation()		);buffer.append(_TEXT("\r\n"));
	
	return buffer;
}
