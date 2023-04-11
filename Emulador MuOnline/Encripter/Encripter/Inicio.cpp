#include "stdafx.h"
#include "EncDec.h"

unsigned char OutBuff[1024];


extern "C" __declspec( dllexport ) char *Dec(int Tipo, unsigned char *buff)
{
	char TrimString[1024] = {0};
	unsigned char *DecC3=new unsigned char[(buff[1]-2)*8/11];
	int DecLen=DecryptC3asServer(DecC3,buff+2,buff[1]-2);

	if(Tipo == 1){
	sprintf(TrimString, "%d", DecC3[1]);
	}

	if(Tipo == 2){
	DecXor32(DecC3+1,2,DecLen-1);
	sprintf(TrimString, "%d", DecC3[2]);
	}

	if(Tipo == 3){
		DecXor32(DecC3+1,2,DecLen-1);
		EncDecLogin(DecC3+3,10);								//Decrypt Login
		EncDecLogin(DecC3+13,10);
		char Login[11]={0},Password[11]={0},ClientSerial[17]={0},ClientVersion[6]={0};
		strncpy(Login,(char*)DecC3+3,10);
		strncpy(Password,(char*)DecC3+13,10);
		strncpy(ClientVersion,(char*)DecC3+27,5);
		strncpy(ClientSerial,(char*)DecC3+32,16);
		sprintf(TrimString, "%s=%s=%s=%s", Login, Password, ClientVersion, ClientSerial);
	}

	if(Tipo == 4){
		DecXor32(&buff[3],3,buff[1]-3);
		memcpy(TrimString, (char*)buff + 4, 10);
	}

	if(Tipo == 5){
		//DecXor32(&buff[3],3,buff[1]-3);
		unsigned char * tmpbuf = new unsigned char[15];

		EncryptC3asServer((unsigned char *)tmpbuf, (unsigned char *)&buff, 15);
		sprintf(TrimString, "%s", tmpbuf);
	}


	int i;
	char *OutTrimBuffer;
	char *garbagestring = new char[strlen(TrimString)+1];
	strcpy(garbagestring,TrimString);
	i=strlen(garbagestring);
	while(i--)
	{
		if(TrimString[i]!=32) break;
	}
	garbagestring[++i]=0;
	if (garbagestring)
	{
		for (OutTrimBuffer=garbagestring;*OutTrimBuffer && (OutTrimBuffer[0]==32); ++OutTrimBuffer);
		if (garbagestring!=OutTrimBuffer)
			memcpy(garbagestring,OutTrimBuffer,strlen(OutTrimBuffer)+1);
	}

	return garbagestring;
}
