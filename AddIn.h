
#ifndef __ADDINNATIVE_H__
#define __ADDINNATIVE_H__

#ifndef __linux__
#include <wtypes.h>
#endif //__linux__

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"

#define TV_IS_INT(X)  ( (X)->vt == VTYPE_I4 || (X)->vt == VTYPE_I2 || (X)->vt == VTYPE_UI1 )

extern AppCapabilities m_Capabilities;

long FindName(const WCHAR_T* name, const WCHAR_T* names[], size_t size);

size_t AllocateStringCopy(IMemoryManager *mm, WCHAR_T **dst, const WCHAR_T *src);
size_t AllocateVarStringCopy(IMemoryManager *mm, tVariant* dst, const WCHAR_T* src);

#endif //__ADDINNATIVE_H__
