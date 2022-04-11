
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

		// Exception �߻��� ��ġ ǥ���ϱ� ���� �߰�
		tstring		m_Location;
	};
};
