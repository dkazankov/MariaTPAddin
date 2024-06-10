
#define _CRT_SECURE_NO_WARNINGS

#include "stdafx.h"

#include <strsafe.h>

#include "AddIn.h"
#include "MariaTPAddin.h"

const WCHAR_T *MARIATP_CLASS_NAME = L"MariaTP";

//---------------------------------------------------------------------------//
//CAddInNative
CMariaTPAddin::CMariaTPAddin()
{
	m_Connection = nullptr;
	m_MemoryManager = nullptr;
	m_LocaleName = nullptr;

	m_hFile = INVALID_HANDLE_VALUE;
	m_ReadPos = sizeof(m_ReadBuffer);
	m_ReadCount = 0;
	m_UseCRC = false;
	m_LastErrorCode = 0;
	m_LastErrorMessage = nullptr;

	m_hDebugLog = INVALID_HANDLE_VALUE;
}
//---------------------------------------------------------------------------//
CMariaTPAddin::~CMariaTPAddin()
{
	if (m_LastErrorMessage != nullptr)
	{
		//::HeapFree(::GetProcessHeap(), 0, m_LastErrorMessage);
		delete m_LastErrorMessage;
		m_LastErrorMessage = nullptr;
	}
	if (m_LocaleName != nullptr)
	{
		//::HeapFree(::GetProcessHeap(), 0, m_LocaleName);
		delete m_LocaleName;
		m_LocaleName = nullptr;
	}
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::Init(void* pConnection)
{ 
    m_Connection = (IAddInDefBase*)pConnection;

	if (m_Connection == nullptr)
		return false;

	//if (m_Capabilities >= eAppCapabilities1)
	//{
	//	IAddInDefBaseEx* conn = (IAddInDefBaseEx*)m_Connection;
	//	IMsgBox* imsgbox = (IMsgBox*)conn->GetInterface(eIMsgBox);
	//	if (imsgbox)
	//	{
	//		IPlatformInfo* pinfo = (IPlatformInfo*)conn->GetInterface(eIPlatformInfo);
	//		assert(pinfo);
	//		const IPlatformInfo::AppInfo* appinfo = pinfo->GetPlatformInfo();
	//	}
	//}
    return true;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::setMemManager(void* mem)
{
    m_MemoryManager = (IMemoryManager*)mem;
    return m_MemoryManager != nullptr;
}
//---------------------------------------------------------------------------//
void CMariaTPAddin::SetLocale(const WCHAR_T* loc)
{
	//m_LocaleName = _wsetlocale(LC_ALL, loc);

	size_t len;
	HRESULT result = ::StringCchLength(loc, STRSAFE_MAX_CCH, &len);
	if (FAILED(result))
	{
		return;
	}

	if (m_LocaleName != nullptr)
	{
		//::HeapFree(::GetProcessHeap(), 0, m_LocaleName);
		delete m_LocaleName;
		m_LocaleName = nullptr;
	}

	//m_LocaleName = (WCHAR_T *) ::HeapAlloc(::GetProcessHeap(), 0, (len + 1) * sizeof(wchar_t));
	m_LocaleName = new wchar_t[len + 1];
	if (m_LocaleName == nullptr)
		return;

	result = ::StringCchCopy(m_LocaleName, len + 1, loc);
	if (FAILED(result))
	{
		return;
	}
}
//---------------------------------------------------------------------------//
long CMariaTPAddin::GetInfo()
{ 
    return 2000; 
}
//---------------------------------------------------------------------------//
void CMariaTPAddin::Done()
{
	Close();

	if (m_hDebugLog != INVALID_HANDLE_VALUE)
	{
		DebugMessageStart();
		DebugMessage("STOP LOGGING");
		DebugMessageEnd();

		::CloseHandle(m_hDebugLog);
		m_hDebugLog = INVALID_HANDLE_VALUE;
	}
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
	return AllocateStringCopy(m_MemoryManager, wsExtensionName, MARIATP_CLASS_NAME) != 0;
}

//---------------------------------------------------------------------------//
long CMariaTPAddin::GetNProps()
{ 
    return ePropLast;
}
//---------------------------------------------------------------------------//
static const wchar_t *g_PropNames[] =
{
	L"UseCRC"
};

long CMariaTPAddin::FindProp(const WCHAR_T* wsPropName)
{ 
    return FindName(wsPropName, g_PropNames, ePropLast);
}
//---------------------------------------------------------------------------//
const WCHAR_T* CMariaTPAddin::GetPropName(long lPropNum, long lPropAlias)
{ 
    if (lPropNum >= ePropLast)
        return nullptr;

    wchar_t *name = nullptr;

    switch(lPropAlias)
    {
    case 0: // First language
    case 1: // Second language
        name = (wchar_t *) g_PropNames[lPropNum];
        break;
    default:
        return nullptr;
    }

	assert(name != nullptr);

    WCHAR_T *result = nullptr;
	if (AllocateStringCopy(m_MemoryManager, &result, name) == 0)
        return nullptr;

    return result;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::IsPropReadable(const long lPropNum)
{ 
	assert(lPropNum < ePropLast);

    switch(lPropNum)
    { 
	case ePropUseCRC:
		return true;
    }
	return false;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::IsPropWritable(const long lPropNum)
{
	assert(lPropNum < ePropLast);

	switch (lPropNum)
	{
	case ePropUseCRC:
		return true;
	}
	return false;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{ 
	assert(lPropNum < ePropLast);

	SetLastError(0);

	tVarInit(pvarPropVal);

	switch (lPropNum)
    {
	case ePropUseCRC:
		TV_BOOL(pvarPropVal) = m_UseCRC;
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		return true;
	}
	return false;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::SetPropVal(const long lPropNum, tVariant* varPropVal)
{
	assert(lPropNum < ePropLast);

	SetLastError(0);

	switch (lPropNum)
	{
	case ePropUseCRC:
		if (TV_VT(varPropVal) != VTYPE_BOOL)
		{
			SetLastError(ERROR_INVALID_DATATYPE);
			return false;
		}

		m_UseCRC = TV_BOOL(varPropVal);
		return true;
	}
	return false;
}

//---------------------------------------------------------------------------//
long CMariaTPAddin::GetNMethods()
{ 
    return eMethLast;
}
//---------------------------------------------------------------------------//
static const wchar_t *g_MethodNames[] =
{
	L"GetVersion", L"GetDescription", L"GetLastErrorCode", L"GetLastErrorDescription", L"Open", L"Close", L"SendConnect", L"SendReconnect", L"Send", L"Receive", L"SetDebugLog"
};

long CMariaTPAddin::FindMethod(const WCHAR_T* wsMethodName)
{ 
    return FindName(wsMethodName, g_MethodNames, eMethLast);
}
//---------------------------------------------------------------------------//
const WCHAR_T* CMariaTPAddin::GetMethodName(const long lMethodNum, const long lMethodAlias)
{ 
    if (lMethodNum >= eMethLast)
        return nullptr;

    wchar_t *name = nullptr;

    switch(lMethodAlias)
    {
    case 0: // First language
    case 1: // Second language
        name = (wchar_t *) g_MethodNames[lMethodNum];
        break;
    default: 
        return nullptr;
    }

	assert(name != nullptr);

    WCHAR_T *result = nullptr;
	if (AllocateStringCopy(m_MemoryManager, &result, name) == 0)
        return nullptr;

    return result;
}
//---------------------------------------------------------------------------//
long CMariaTPAddin::GetNParams(const long lMethodNum)
{ 
	assert(lMethodNum < eMethLast);

    switch(lMethodNum)
    { 
	case eMethSend:
		return 2;
	case eMethOpen:
	case eMethSendConnect:
	case eMethReceive:
	case eMethSetDebugLog:
		return 1;
    }
	return 0;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::GetParamDefValue(const long lMethodNum, const long lParamNum, tVariant *pvarParamDefValue)
{ 
	assert(lMethodNum < eMethLast);

	tVarInit(pvarParamDefValue);

	switch (lMethodNum)
	{
	case eMethSend:
		switch (lParamNum)
		{
		case 1:
			TV_I4(pvarParamDefValue) = 2000;
			TV_VT(pvarParamDefValue) = VTYPE_I4;
			return true;
		}
		break;
	case eMethReceive:
		switch (lParamNum)
		{
		case 0:
			TV_I4(pvarParamDefValue) = 6000;
			TV_VT(pvarParamDefValue) = VTYPE_I4;
			return true;
		}
		break;
	}

	return false;
} 
//---------------------------------------------------------------------------//
bool CMariaTPAddin::HasRetVal(const long lMethodNum)
{ 
	assert(lMethodNum < eMethLast);

	switch (lMethodNum)
	{
	case eMethGetVersion:
	case eMethGetDescription:
	case eMethGetLastErrorCode:
	case eMethGetLastErrorDescription:
	case eMethReceive:
		return true;
	}

	return false;
}

//---------------------------------------------------------------------------//
bool CMariaTPAddin::CallAsProc(const long lMethodNum, tVariant* paParams, const long lSizeArray)
{
	assert(lMethodNum < eMethLast);

	SetLastError(0);

	long error = 0L;
	switch (lMethodNum)
    {
    case eMethOpen:
		if (TV_VT(paParams) != VTYPE_PWSTR)
		{
			SetLastError(ERROR_INVALID_DATATYPE, L"Param #1");
			return false;
		}
		{
			WCHAR_T param0[MAX_PATH + 1];
			HRESULT result = ::StringCchCopyN(param0, MAX_PATH + 1, paParams->pwstrVal, paParams->wstrLen);
			if (FAILED(result))
			{
				SetLastError(HRESULT_CODE(result));
				return false;
			}

			error = Open(param0);
			if (error)
			{
				SetLastError(error);
				return false;
			}
		}
		return true;

	case eMethClose:
		if (error = Close())
		{
			SetLastError(error);
			return false;
		}
		return true;

	case eMethSendConnect:
		if (!TV_IS_INT(paParams))
		{
			SetLastError(ERROR_INVALID_DATATYPE, L"Param #1");
			return false;
		}

		if (error = SendConnect(TV_I4(paParams)))
		{
			SetLastError(error);
			return false;
		}
		return true;

	case eMethSendReconnect:
		if (error = SendReconnect())
		{
			SetLastError(error);
			return false;
		}
		return true;

	case eMethSend:
		if (TV_VT(paParams) != VTYPE_PWSTR)
		{
			SetLastError(ERROR_INVALID_DATATYPE, L"Param #1");
			return false;
		}
		if (!TV_IS_INT(&paParams[1]))
		{
			SetLastError(ERROR_INVALID_DATATYPE, L"Param #2");
			return false;
		}
		if (paParams->wstrLen > 252)
		{
			SetLastError(ERROR_INVALID_BLOCK_LENGTH);
			return false;
		}

		{
			char param0[252];

			// CP_OEMCP
			if (!::WideCharToMultiByte(CP_OEMCP, 0, paParams->pwstrVal, paParams->wstrLen, param0, 252, NULL, NULL))
			{
				SetLastError(::GetLastError());
				return false;
			}

			error = Send(param0, paParams->wstrLen, TV_I4(&paParams[1]));
			if (error)
			{
				SetLastError(error);
				return false;
			}
		}
		return true;
	case eMethSetDebugLog:
		if (TV_VT(paParams) != VTYPE_PWSTR)
		{
			SetLastError(ERROR_INVALID_DATATYPE);
			return false;
		}

		{
			if (m_hDebugLog != INVALID_HANDLE_VALUE)
			{
				DebugMessageStart();
				DebugMessage("STOP LOGGING");
				DebugMessageEnd();

				::CloseHandle(m_hDebugLog);
				m_hDebugLog = INVALID_HANDLE_VALUE;
			}
			if (paParams->wstrLen > 0)
			{
				WCHAR_T name[MAX_PATH + 1];
				HRESULT result = ::StringCchCopyN(name, MAX_PATH + 1, paParams->pwstrVal, paParams->wstrLen);
				if (FAILED(result))
				{
					SetLastError(HRESULT_CODE(result));
					return false;
				}

				m_hDebugLog = ::CreateFile(name, GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
				if (m_hDebugLog == INVALID_HANDLE_VALUE)
				{
					SetLastError(::GetLastError());
					return false;
				}

				::SetFilePointer(m_hDebugLog, 0, NULL, FILE_END);

				DebugMessageStart();
				DebugMessage("START LOGGING");
				DebugMessageEnd();
			}
		}
		return true;
	}
	return false;
}
//---------------------------------------------------------------------------//
bool CMariaTPAddin::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
	assert(lMethodNum < eMethLast);

	tVarInit(pvarRetValue);
	SetLastError(0);

	switch (lMethodNum)
	{
	case eMethGetVersion:
		AllocateVarStringCopy(m_MemoryManager, pvarRetValue, MARIATP_VERSION);
		return true;

	case eMethGetDescription:
		AllocateVarStringCopy(m_MemoryManager, pvarRetValue, MARIATP_DESCRIPTION);
		return true;

	case eMethGetLastErrorCode:
		TV_I4(pvarRetValue) = HRESULT_FROM_WIN32(m_LastErrorCode);
		TV_VT(pvarRetValue) = VTYPE_I4;
		return true;

	case eMethGetLastErrorDescription:
		if (m_LastErrorCode != 0 && m_LastErrorMessage != NULL)
			AllocateVarStringCopy(m_MemoryManager, pvarRetValue, m_LastErrorMessage);
		return true;

	case eMethReceive:
		if (!TV_IS_INT(paParams))
		{
			SetLastError(ERROR_INVALID_DATATYPE, L"Param #1");
			return false;
		}

		{
			long error = 0L;
			long len = 0L;
			char param0[252];

			long rerror = Receive(param0, sizeof(param0), &len, TV_I4(paParams));
			if (len == 0L)
			{
				SetLastError(rerror ? rerror : ENODATA);
				return false;
			}

			if (!m_MemoryManager->AllocMemory((void**)&pvarRetValue->pwstrVal, len*sizeof(WCHAR_T)))
			{
				SetLastError(ERROR_NOT_ENOUGH_MEMORY);
				return false;
			}

			pvarRetValue->wstrLen = len;
			TV_VT(pvarRetValue) = VTYPE_PWSTR;

			if (!::MultiByteToWideChar(CP_OEMCP, 0, param0, len, pvarRetValue->pwstrVal, pvarRetValue->wstrLen))
			{
				SetLastError(::GetLastError());
				return false;
			}

			if (rerror)
				SetLastError(rerror);
		}
		return true;
	}

	return false;
}

//---------------------------------------------------------------------------//
bool CMariaTPAddin::SetLastError(long code, const WCHAR_T* message)
{
	m_LastErrorCode = code;
	if (m_LastErrorMessage != nullptr)
	{
		//::HeapFree(::GetProcessHeap(), 0, m_LastErrorMessage);
		delete m_LastErrorMessage;
		m_LastErrorMessage = nullptr;
	}

	if (code == 0L)
		return true;

	size_t MessageLen = 0;
	HRESULT result = ::StringCchLength(message, STRSAFE_MAX_CCH, &MessageLen);
	if (FAILED(result))
	{
		return false;
	}

	LPWSTR buf;
	size_t SystemMessageLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_IGNORE_INSERTS,
		NULL,
		code,
		0,
		(LPWSTR)&buf,
		0,
		NULL);

	size_t len = SystemMessageLen;
	if (MessageLen > 0)
	{
		len += MessageLen + 2;
	}

	//m_LastErrorMessage = (wchar_t *) ::HeapAlloc(::GetProcessHeap(), 0, (len + 1) * sizeof(wchar_t));
	m_LastErrorMessage = new WCHAR_T[len + 1];
	if (m_LastErrorMessage == nullptr)
	{
		return 0;
	}

	if (MessageLen > 0)
	{
		result = ::StringCchCopy(m_LastErrorMessage, len + 1, message);
		if (FAILED(result))
		{
			return false;
		}
		result = ::StringCchCat(m_LastErrorMessage, len + 1, L": ");
		if (FAILED(result))
		{
			return false;
		}
	}
	result = ::StringCchCat(m_LastErrorMessage, len + 1, buf);
	if (FAILED(result))
	{
		return false;
	}

	if (SystemMessageLen > 0)
		::HeapFree(::GetProcessHeap(), 0, buf);

	return false;
}

#define CHAR253 ((unsigned char)253)
#define CHAR254 ((unsigned char)254)

static unsigned char buf[257];

long CMariaTPAddin::Open(WCHAR_T *device)
{
	DWORD error;

	m_hFile = ::CreateFile(device, GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0);
	if (m_hFile == INVALID_HANDLE_VALUE)
		return ::GetLastError();

	if (!::SetupComm(m_hFile, 257, 257))
	{
		error = ::GetLastError();
		::CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
		return error;
	}

	if (!::SetCommMask(m_hFile, EV_BREAK | EV_ERR | EV_RXCHAR))
	{
		error = ::GetLastError();
		::CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
		return error;
	}

	return 0L;
}

long CMariaTPAddin::SendConnect(long speed)
{
	//static const unsigned char U = 'U';

	long event;
	long bytes;
	//char buf[10];

	DWORD error;
	DWORD status;
	DWORD mask;
	DCB dcb;
	COMMTIMEOUTS timeouts;

	m_UseCRC = false;

	timeouts.ReadIntervalTimeout = 1;
	timeouts.ReadTotalTimeoutMultiplier = 0;
	timeouts.ReadTotalTimeoutConstant = 0;
	timeouts.WriteTotalTimeoutMultiplier = 0;
	timeouts.WriteTotalTimeoutConstant = 0;

	if (!::SetCommTimeouts(m_hFile, &timeouts))
	{
		error = ::GetLastError();
		::CloseHandle(m_hFile);
		m_hFile = INVALID_HANDLE_VALUE;
		return error;
	}

	if (!::EscapeCommFunction(m_hFile, CLRDTR))
		return ::GetLastError();

	::Sleep(3000);

	dcb.DCBlength = sizeof(DCB);
	if (!::GetCommState(m_hFile, &dcb))
		return ::GetLastError();

	dcb.ByteSize = 8;
	dcb.fParity = TRUE;
	dcb.Parity = EVENPARITY;
	dcb.StopBits = TWOSTOPBITS;

	dcb.fBinary = TRUE;
	dcb.fDtrControl = DTR_CONTROL_ENABLE;
	dcb.fDsrSensitivity = FALSE;
	dcb.fTXContinueOnXoff = FALSE;
	dcb.fOutX = FALSE;
	dcb.fInX = FALSE;
	dcb.fErrorChar = FALSE;
	dcb.fNull = FALSE;
	dcb.fRtsControl = RTS_CONTROL_ENABLE;
	dcb.fAbortOnError = FALSE;
	dcb.fOutxCtsFlow = FALSE;
	dcb.fOutxDsrFlow = FALSE;

	dcb.BaudRate = speed;

	DebugMessageStart();
	DebugMessage("SetCommState(");

	HRESULT result = StringCchPrintfA((char *) buf, 10, "%u", speed);
	if (SUCCEEDED(result))
		DebugMessage((char *)buf);
	DebugMessage(")");

	if (!::SetCommState(m_hFile, &dcb))
	{
		error = ::GetLastError();
		DebugMessageEnd(error);
		return error;
	}

	DebugMessageEnd();

	DebugMessageStart();
	DebugMessage("GetCommModemStatus()");

	if (!::GetCommModemStatus(m_hFile, &status))
	{
		error = ::GetLastError();
		DebugMessageEnd(error);
		return error;
	}
	DebugMessageEnd();

	if (!(status & (MS_CTS_ON | MS_DSR_ON)))
	{
		DebugMessageStart();
		DebugMessage("GetCommMask()");

		if (!::GetCommMask(m_hFile, &mask))
		{
			error = ::GetLastError();
			DebugMessageEnd(error);
			return error;
		}
		DebugMessageEnd();

		DebugMessageStart();
		DebugMessage("SetCommMask(EV_CTS|EV_DSR)");

		if (!::SetCommMask(m_hFile, EV_CTS | EV_DSR))
		{
			error = ::GetLastError();
			DebugMessageEnd(error);
			return error;
		}
		DebugMessageEnd();

		DebugMessageStart();
		DebugMessage("WaitEvent(6000)");

		error = WaitEvent(&event, 6000);
		if (error)
		{
			DebugMessageEnd(error);
			//return error;
		}
		else
		{
			DebugMessageEnd();
		}

		DebugMessageStart();
		DebugMessage("SetCommMask()");
		if (!::SetCommMask(m_hFile, mask))
		{
			error = ::GetLastError();
			DebugMessageEnd(error);
			return error;
		}
		DebugMessageEnd();
	}

	::Sleep(1000);

	DebugMessageStart();
	DebugMessage("->U");

	buf[0] = 'U';

	if (error = Write(buf, 1, &bytes, 1000))
	{
		DebugMessageEnd(error);
		return error;
	}

	::Sleep(5);

	DebugMessage("U");

	if (error = Write(buf, 1, &bytes, 1000))
	{
		DebugMessageEnd(error);
		return error;
	}

	DebugMessageEnd();
	return 0L;
}

long CMariaTPAddin::SendReconnect()
{
	long bytes = 0;
	//char buf[6];

	m_UseCRC = false;

	DWORD error;
	DWORD status;
	//if (!::EscapeCommFunction(m_hFile, SETDTR))
	//{
	//	error = ::GetLastError();
	//	return error;
	//}

	DebugMessageStart();
	DebugMessage("GetCommModemStatus()");

	if (!::GetCommModemStatus(m_hFile, &status))
	{
		error = ::GetLastError();
		DebugMessageEnd(error);
		return error;
	}

	if (!(status & (MS_CTS_ON | MS_DSR_ON)))
	{
		DebugMessageEnd(ERROR_DEVICE_NOT_CONNECTED);
		return ERROR_DEVICE_NOT_CONNECTED;
	}

	buf[0] = buf[1] = CHAR253;
	buf[2] = buf[3] = buf[4] = buf[5] = CHAR254;

	DebugMessageStart();
	DebugMessage("->");
	DebugMessage((char *)buf, 6);

	error = Write(buf, 6, &bytes, 1000);
	if (error)
	{
		DebugMessageEnd(error);
		return error;
	}

	DebugMessageEnd();
	return 0L;
}

long CMariaTPAddin::Close()
{
	if (m_hFile == INVALID_HANDLE_VALUE)
		return 0L;

	if (!::CloseHandle(m_hFile))
		return ::GetLastError();

	m_hFile = INVALID_HANDLE_VALUE;
	return 0L;
}

unsigned short CRC16(void *mem, unsigned long len)
{
	unsigned short a, crc16 = 0;
	unsigned char *pch = (unsigned char *)mem;
	while (len--)
	{
		crc16 ^= *pch++;
		a = (crc16 ^ (crc16 << 4)) & 0x00FF;
		crc16 = (crc16 >> 8) ^ (a << 8) ^ (a << 3) ^ (a >> 4);
	}
	return crc16;
}

long CMariaTPAddin::Send(const char *block, long length, long timeout)
{
	long written = 0;
	long error = 0;
	unsigned short crc;
	long len = 0;
	//unsigned char buf[257];
	unsigned char *ptr = (unsigned char *)block;

	DebugMessageStart();
	DebugMessage("->");

	if (length < 4 || length > 252)
	{
		DebugMessageEnd(ERROR_BAD_LENGTH);
		return ERROR_BAD_LENGTH;
	}

	// Block structure (253)...(block length)(254)[CRC]
	buf[len++] = CHAR253;
	for (long i = 0; i < length; i++)
	{
		if (*ptr > 252)
		{
			DebugMessageEnd(ERROR_INVALID_DATA);
			return ERROR_INVALID_DATA;
		}
		buf[len++] = *ptr++;
	}
	buf[len++] = (unsigned char)(length + 1);
	buf[len++] = CHAR254;

	if (m_UseCRC)
	{
		crc = CRC16(buf, len);
		buf[len++] = (unsigned char)crc;
		buf[len++] = (unsigned char)(crc >> 8);
	}

	DebugMessage((char *)buf, len);

	if (error = Write(buf, len, &written, timeout))
	{
		DebugMessageEnd(error);
		return error;
	}

	DebugMessageEnd();
	return 0L;
}

long CMariaTPAddin::Receive(char *block, long size, long *length, long timeout)
{
	long read = 0;
	long error = 0;
	unsigned char c;
	unsigned short crc = 0;
	unsigned short bcrc = 0;
	long blen = 0;
	long bsize = 0;
	long len = 0;
	//unsigned char buf[257];
	unsigned char *ptr = (unsigned char *)block;

	*length = 0;
	*block = '\0';

	DebugMessageStart();
	DebugMessage("<-");

	bool started = false;
	bool stopped = false;
	long CRCBytes = 0;

	// Block structure (253)...(block length)(254)[CRC]
	while (true)
	{
		if (m_ReadCount == 0)
		{
			m_ReadPos = 0;
			if (error = Read(m_ReadBuffer, sizeof(m_ReadBuffer), &m_ReadCount, timeout))
			{
				if (len > 0)
					DebugMessage((char *)buf, len);
				DebugMessageEnd(error);
				return error;
			}
			if (m_ReadCount == 0)
			{
				if (len > 0)
					DebugMessage((char *)buf, len);
				DebugMessageEnd(ENODATA);
				return ENODATA;
			}
		}
		while (m_ReadCount > 0)
		{
			m_ReadCount--;
			c = m_ReadBuffer[m_ReadPos++];
			if (c == CHAR253)
			{
				if (len > 0)
				{
					DebugMessage((char *)buf, len);
					DebugMessageEnd(ENODATA);
					DebugMessageStart();
					DebugMessage("<-");
				}
				started = true;
				stopped = false;
				len = 0;
				buf[len++] = c;
			}
			else if (!started)
			{
				buf[len++] = c;
				if (len > 253)
				{
					DebugMessage((char *)buf, len);
					DebugMessageEnd(ENODATA);
					return ENODATA;
				}
			}
			else if (!stopped)
			{
				buf[len++] = c;
				if (c == CHAR254)
				{
					stopped = true;
					if (!m_UseCRC)
						break;
				}
				if (len > 255)
				{
					DebugMessage((char *)buf, len);
					DebugMessageEnd(ERROR_BAD_LENGTH);
					return ERROR_BAD_LENGTH;
				}
			}
			else if (m_UseCRC)
			{
				CRCBytes++;
				if (CRCBytes == 1)
					bcrc = c;
				else if (CRCBytes == 2)
				{
					bcrc |= ((unsigned short)c) << 8;
					DebugMessage((char *)&bcrc, 2);
					break;
				}
				else
					break;
			}
			else
			{
				// Must never reach
				if (len > 0)
					DebugMessage((char *)buf, len);
				DebugMessageEnd(EOTHER);
				return EOTHER;
			}
		}
		if (!m_UseCRC && stopped)
			break;
		if (m_UseCRC && CRCBytes > 1)
			break;
	}

	DebugMessage((char *)buf, len);
	if (len < 2)
	{
		DebugMessageEnd(ERROR_BAD_LENGTH);
		return ERROR_BAD_LENGTH;
	}

	blen = buf[len - 2];

	if (len - 2 == blen)
		bsize = len - 2;
	else
		bsize = len - 1;

	if (bsize > 253)
		bsize = 253;

	for (long i = 1; i<bsize; i++)
	{
		if ((*length) + 1 > size)
		{
			DebugMessageEnd(ERROR_INSUFFICIENT_BUFFER);
			return ERROR_INSUFFICIENT_BUFFER;
		}
		if (buf[i] > 252)
		{
			DebugMessageEnd(ERROR_INVALID_DATA);
			return ERROR_INVALID_DATA;
		}
		ptr[i - 1] = buf[i];
		(*length)++;
	}

	if (len - 2 != blen || blen < 4 || blen > 253)
	{
		DebugMessageEnd(ERROR_BAD_LENGTH);
		return ERROR_BAD_LENGTH;
	}

	if (m_UseCRC)
	{
		crc = CRC16(buf, len);
		if (crc != bcrc)
		{
			DebugMessageEnd(ERROR_CRC);
			return ERROR_CRC;
		}
	}

	DebugMessageEnd();
	return 0L;
}

long CMariaTPAddin::WaitEvent(long *Event, long timeout)
{
	if (m_hFile == INVALID_HANDLE_VALUE)
		return ERROR_INVALID_HANDLE;

	long error = 0L;

	DWORD bytes;
	OVERLAPPED ovInternal = {0};
	ovInternal.hEvent = ::CreateEvent(0, TRUE, FALSE, 0);
	if (ovInternal.hEvent == NULL)
		return ::GetLastError();

	if (!::WaitCommEvent(m_hFile, (DWORD *)Event, &ovInternal))
	{
		error = ::GetLastError();

		if (error == ERROR_IO_PENDING)
		{
			switch (::WaitForSingleObject(ovInternal.hEvent, timeout))
			{
			case WAIT_OBJECT_0:
				if (!::GetOverlappedResult(m_hFile, &ovInternal, &bytes, TRUE))
					error = ::GetLastError();
				else
					error = NOERROR;
				break;
			case WAIT_TIMEOUT:
				::CancelIo(m_hFile);
				error = ERROR_TIMEOUT;
				break;
			default:
				error = ::GetLastError();
			}
		}
	}

	::CloseHandle(ovInternal.hEvent);

	return error;
}

long CMariaTPAddin::Write(const void* data, long length, long* written, long Timeout)
{
	if (m_hFile == INVALID_HANDLE_VALUE)
		return ERROR_INVALID_HANDLE;

	long error = 0L;
	*written = 0;

	OVERLAPPED ovInternal = {0};
	ovInternal.hEvent = ::CreateEvent(0, TRUE, FALSE, 0);
	if (ovInternal.hEvent == NULL)
		return ::GetLastError();

	if (!::WriteFile(m_hFile, data, length, NULL, &ovInternal))
	{
		error = ::GetLastError();
		if (error == ERROR_IO_PENDING)
		{
			switch (::WaitForSingleObject(ovInternal.hEvent, Timeout))
			{
			case WAIT_OBJECT_0:
				if (!::GetOverlappedResult(m_hFile, &ovInternal, (DWORD *)written, TRUE))
					error = ::GetLastError();
				else
					error = NOERROR;
				break;

			case WAIT_TIMEOUT:
				::CancelIo(m_hFile);
				error = ERROR_TIMEOUT;
				break;

			default:
				error = ::GetLastError();
			}
		}
	}

	::CloseHandle(ovInternal.hEvent);

	return error;
}

long CMariaTPAddin::Read(void* data, long length, long* read, long timeout)
{
	if (m_hFile == INVALID_HANDLE_VALUE)
		return ERROR_INVALID_HANDLE;

	long error = 0L;
	*read = 0;

	OVERLAPPED ovInternal = {0};
	ovInternal.hEvent = ::CreateEvent(0, TRUE, FALSE, 0);
	if (ovInternal.hEvent == NULL)
		return ::GetLastError();

	if (!::ReadFile(m_hFile, data, length, NULL, &ovInternal))
	{
		error = ::GetLastError();
		if (error == ERROR_IO_PENDING)
		{
			switch (::WaitForSingleObject(ovInternal.hEvent, timeout))
			{
			case WAIT_OBJECT_0:
				if (!::GetOverlappedResult(m_hFile, &ovInternal, (DWORD *)read, TRUE))
					error = ::GetLastError();
				else
					error = NOERROR;
				break;

			case WAIT_TIMEOUT:
				::CancelIo(m_hFile);
				error = ERROR_TIMEOUT;
				break;

			default:
				error = ::GetLastError();
			}
		}
	}

	::CloseHandle(ovInternal.hEvent);

	return error;
}

void CMariaTPAddin::DebugMessageStart()
{
	if (m_hDebugLog == INVALID_HANDLE_VALUE)
		return;

	SYSTEMTIME t;
	GetLocalTime(&t);

	char buf[19];

	HRESULT result = StringCchPrintfA(buf, 19, "%02u/%02u/%02u %02u:%02u:%02u:", t.wYear, t.wMonth, t.wDay, t.wHour, t.wMinute, t.wMonth);
	if (SUCCEEDED(result))
	{
		DWORD written;
		::WriteFile(m_hDebugLog, buf, 18, &written, NULL);
	}
}

void CMariaTPAddin::DebugMessage(const char* message, size_t len)
{
	if (m_hDebugLog == INVALID_HANDLE_VALUE)
		return;

	if (message == nullptr)
		return;

	if (len == 0)
	{
		HRESULT result = ::StringCbLengthA(message, STRSAFE_MAX_CCH, &len);
		if (FAILED(result))
			return;
	}

	DWORD written;
	::WriteFile(m_hDebugLog, message, len, &written, NULL);
}

void CMariaTPAddin::DebugMessageEnd(long code)
{
	if (m_hDebugLog == INVALID_HANDLE_VALUE)
		return;

	DWORD written;
	if (code != 0)
	{
		::WriteFile(m_hDebugLog, "0x", 2, &written, NULL);

		char buf[9];

		HRESULT result = StringCchPrintfA(buf, 9, "%08X", code);
		if (SUCCEEDED(result))
			::WriteFile(m_hDebugLog, buf, 8, &written, NULL);
	}
	::WriteFile(m_hDebugLog, "\n", 1, &written, NULL);
}
