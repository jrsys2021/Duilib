#include "stdafx.h"
#include "DBEntity.h"


DBEntity::DBEntity()
	:mIsConnect(false)
{
}

DBEntity::~DBEntity()
{
}

bool DBEntity::Connect(std::string ip,
	UINT port, std::string name,
	std::string user, std::string pwd)
{
	CoInitialize(0);
	SetConnectionInfo((char*)ip.c_str(), port, (char*)name.c_str(), (char*)user.c_str(), (char*)pwd.c_str());
	if (IsConnect())
	{
		mIp= ip;
		mName= name;
		mUser= user;
		mPwd= pwd;
		mPort= port;
	}
	return IsConnect();
}
bool DBEntity::Connect()
{
	return SetConnectionInfo((char*)mIp.c_str(), mPort, (char*)mName.c_str(), (char*)mUser.c_str(), (char*)mPwd.c_str());
}
