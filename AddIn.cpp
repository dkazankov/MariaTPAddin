
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS

#include "stdafx.h"

#include <strsafe.h>

#include "AddIn.h"
#include "MariaTPAddin.h"

//---------------------------------------------------------------------------//
const WCHAR_T* GetClassNames()
{
	static WCHAR_T *g_ClassNames = L"MariaTP";
	return g_ClassNames;
}
//---------------------------------------------------------------------------//
long GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
	if (*pInterface != nullptr)
		return 0L;

	if (CompareStringEx(LOCALE_NAME_INVARIANT, NORM_IGNORECASE, wsName, -1, MARIATP_CLASS_NAME, -1, NULL, NULL, 0) != CSTR_EQUAL)
		return 0L;

	*pInterface = new CMariaTPAddin();
	return (long)*pInterface;
}
//---------------------------------------------------------------------------//
long DestroyObject(IComponentBase** pInterface)
{
	if (*pInterface == nullptr)
		return -1L;

	//((CMariaTPAddin *)*pInterface)->~CMariaTPAddin();
	delete (CMariaTPAddin *)*pInterface;
	*pInterface = nullptr;
	return 0L;
}
//---------------------------------------------------------------------------//
AppCapabilities m_Capabilities = eAppCapabilitiesInvalid;
AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities)
{
    m_Capabilities = capabilities;
    return eAppCapabilitiesLast;
}

//---------------------------------------------------------------------------//
long FindName(const WCHAR_T* name, const WCHAR_T* names[], size_t size)
{
    for (size_t i = 0; i < size; i++)
		if (CompareStringEx(LOCALE_NAME_INVARIANT, NORM_IGNORECASE, name, -1, names[i], -1, NULL, NULL, 0) == CSTR_EQUAL)
			return i;

    return -1L;
}

size_t AllocateStringCopy(IMemoryManager *mm, WCHAR_T **dst, const WCHAR_T *src)
{
	assert(mm != nullptr);
	assert(dst != nullptr);

	size_t len = 0;
	HRESULT result = ::StringCchLength(src, STRSAFE_MAX_CCH, &len);
	if (FAILED(result))
	{
		return 0;
	}

	if( !mm->AllocMemory((void**)dst, (len + 1) * sizeof(WCHAR_T)) )
		return 0;

	result = ::StringCchCopyN(*dst, len + 1, src, len);
	if (FAILED(result))
	{
		mm->FreeMemory((void**)dst);
		return 0;
	}

	return len+1;
}

size_t AllocateVarStringCopy(IMemoryManager *mm, tVariant* dst, const WCHAR_T* src)
{
	assert(mm != nullptr);
	assert(dst != nullptr);

	tVarInit(dst);

	size_t len = 0;
	HRESULT result = ::StringCchLength(src, STRSAFE_MAX_CCH, &len);
	if (FAILED(result))
	{
		return 0;
	}

	if (!mm->AllocMemory((void**)&dst->pwstrVal, len * sizeof(WCHAR_T)))
		return 0;

	result = ::StringCchCopyN(dst->pwstrVal, len, src, len);
	if (FAILED(result))
	{
		mm->FreeMemory((void**)&dst->pwstrVal);
		return 0;
	}

	TV_VT(dst) = VTYPE_PWSTR;
	dst->wstrLen = len;

	return len + 1;
}
