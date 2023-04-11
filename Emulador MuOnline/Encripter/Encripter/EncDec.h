#include <windows.h>
//////////////////////////////////////////////////////////////////////////////////////
//----------------------------User Functions------------------------------------------
int DecryptC3asClient(unsigned char*Dest,unsigned char*Src,int Len);
int EncryptC3asClient(unsigned char*Dest,unsigned char*Src,int Len);
int DecryptC3asServer(unsigned char*Dest,unsigned char*Src,int Len);
int EncryptC3asServer(unsigned char*Dest,unsigned char*Src,int Len);
int LoadKeys(char*File,unsigned long*Where);
void DecXor32(unsigned char*Buff,int SizeOfHeader,int Len);
void EncXor32(unsigned char*Buff,int SizeOfHeader,int Len);
void EncDecLogin(unsigned char*Buff,int Len);
//-----------------------------internal functions-------------------------------------
int DecryptC3(unsigned char*Dest,unsigned char*Src,int Len,unsigned long*Keys);
int EncryptC3(unsigned char*Dest,unsigned char*Src,int Len,unsigned long*Keys);
int DecC3Bytes(unsigned char*Dest,unsigned char*Src,unsigned long*Keys);
void EncC3Bytes(unsigned char*Dest,unsigned char*Src,int Len,unsigned long*Keys); 
int HashBuffer(unsigned char*Dest,int Param10,unsigned char*Src,int Param18,int Param1c);
void ShiftBuffer(unsigned char*Buff,int Len,int ShiftLen);
 