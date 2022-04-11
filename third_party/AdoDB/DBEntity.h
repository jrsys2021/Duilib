#pragma once
#include <string>
#include "AdoLib.h"
#include "AdoDBImpl.h"
class DBEntity :public AdoLib::CAdoDBImpl
{
public:
	DBEntity();
	~DBEntity();

	bool Connect(std::string ip, 
		UINT port, std::string name,
		std::string user, std::string pwd);
	bool Connect();

private:
	bool mIsConnect;
	std::string mIp;
	std::string mName;
	std::string mUser;
	std::string mPwd;
	UINT        mPort;
};