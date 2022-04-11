
#pragma once

namespace AdoLib
{
	class CAdoException
	{
	public:
		CAdoException(); 
		CAdoException(tstring message, int errorCode, tstring description); 
		CAdoException(tstring message, int errorCode, tstring description, tstring source, tstring location); 
		~CAdoException();
			
		tstring		getMessage();
		int				getErrorCode();
		tstring		getDescription();
		tstring		getSource();
		tstring		getLocation();
		tstring		toString();

	protected:
		tstring		m_message;
		int				m_warnCode;
		tstring		m_description;
		tstring		m_source;

		// Exception 발생한 위치 표시하기 위해 추가
		tstring		m_Location;
	};
};
