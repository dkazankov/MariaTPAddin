
#ifndef __MARIATPADDIN_H__
#define __MARIATPADDIN_H__

extern const wchar_t *MARIATP_CLASS_NAME;

#define MARIATP_VERSION L"1.0.2.4"
#define MARIATP_DESCRIPTION L"Maria transport protocol driver"
//#define ADDIN_EQUIPMENT_TYPE L"‘искальный–егистратор"

#ifdef __linux__
	typedef int HANDLE;
	#define INVALID_HANDLE_VALUE -1
	#define INFINITE 0xFFFFFFFF
#else
//	#include <wtypes.h>
#endif //__linux__

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"

///////////////////////////////////////////////////////////////////////////////
// class CAddInNative
class CMariaTPAddin : public IComponentBase
{
public:
	enum Props
	{
		ePropUseCRC = 0,
		ePropLast      // Always last
	};

	enum Methods
	{
		eMethGetVersion = 0,
		eMethGetDescription,
		eMethGetLastErrorCode,
		eMethGetLastErrorDescription,
		eMethOpen,
		eMethClose,
		eMethSendConnect,
		eMethSendReconnect,
		eMethSend,
		eMethReceive,
		eMethSetDebugLog,
		eMethLast      // Always last
	};

	CMariaTPAddin(void);
	virtual ~CMariaTPAddin();
	// IInitDoneBase
	virtual bool ADDIN_API Init(void*);
	virtual bool ADDIN_API setMemManager(void* mem);
	virtual long ADDIN_API GetInfo();
	virtual void ADDIN_API Done();
	// ILanguageExtenderBase
	virtual bool ADDIN_API RegisterExtensionAs(WCHAR_T** wsExtensionName);
	virtual long ADDIN_API GetNProps();
	virtual long ADDIN_API FindProp(const WCHAR_T* wsPropName);
	virtual const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias);
	virtual bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal);
	virtual bool ADDIN_API SetPropVal(const long lPropNum, tVariant* varPropVal);
	virtual bool ADDIN_API IsPropReadable(const long lPropNum);
	virtual bool ADDIN_API IsPropWritable(const long lPropNum);
	virtual long ADDIN_API GetNMethods();
	virtual long ADDIN_API FindMethod(const WCHAR_T* wsMethodName);
	virtual const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum, const long lMethodAlias);
	virtual long ADDIN_API GetNParams(const long lMethodNum);
	virtual bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum, tVariant *pvarParamDefValue);
	virtual bool ADDIN_API HasRetVal(const long lMethodNum);
	virtual bool ADDIN_API CallAsProc(const long lMethodNum, tVariant* paParams, const long lSizeArray);
	virtual bool ADDIN_API CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
	operator IComponentBase*() { return (IComponentBase*)this; }
	// LocaleBase
	virtual void ADDIN_API SetLocale(const WCHAR_T* loc);
private:
	// Attributes
	IAddInDefBase		*m_Connection;
	IMemoryManager		*m_MemoryManager;
	WCHAR_T *m_LocaleName;

	HANDLE	m_hFile;
	bool	m_UseCRC;
	long	m_Baudrate;

	long Open(WCHAR_T *device);
	long SendConnect(long speed);
	long SendReconnect();
	long Close();
	long Send(const char *block, long length, long timeout = INFINITE);
	long Receive(char *block, long size, long *length, long timeout = INFINITE);

	long WaitEvent(long *Event, long timeout = INFINITE);
	long Write(const void* data, long length, long* Written, long timeout = INFINITE);

	char m_ReadBuffer[257];
	long m_ReadPos;
	long m_ReadCount;
	long Read(void* data, long length, long* read, long timeout = INFINITE);

	long m_LastErrorCode;
	WCHAR_T *m_LastErrorMessage;
	bool SetLastError(long code, const WCHAR_T* message = NULL);

	HANDLE	m_hDebugLog;
	void DebugMessageStart();
	void DebugMessage(const char* message, size_t length = 0);
	void DebugMessageEnd(long code = 0L);
};

#endif //__MARIATPADDIN_H__
