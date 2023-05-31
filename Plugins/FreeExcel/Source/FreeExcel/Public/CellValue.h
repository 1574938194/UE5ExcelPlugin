/**
*        Copyright(c) 2022  Weijian Tian
*
*    Permission is hereby granted, free of charge, to any person obtaining a copy
*    of this softwareand associated documentation files(the "Software"), to
*    deal in the Software without restriction, including without limitation the
*    rights to use, copy, modify, merge, publish, distribute, sublicense, and /or
*    sell copies of the Software, and to permit persons to whom the Software is
*    furnished to do so, subject to the following conditions :
*
*    The above copyright noticeand this permission notice shall be included in
*    all copies or substantial portions of the Software.
*
*   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
*    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
*    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
*    IN THE SOFTWARE.
*/

#pragma once 
#include "CoreMinimal.h"
#include <cstdint>
#include <cmath>
#include <string>
#include "CellValue.generated.h"

UENUM(BlueprintType)
enum class EValueType :uint8 {
    Empty UMETA(DisplayName = "Empty"),
    Boolean UMETA(DisplayName = "Boolean"),
    Integer UMETA(DisplayName = "Integer"),
    Float UMETA(DisplayName = "Float"),
    Error,
    String UMETA(DisplayName = "String")
};



class UCell;
/**
 * 
 */
USTRUCT(BlueprintType, Meta = (HasNativeMake, HasNativeBreak))
struct FREEEXCEL_API FCellValue
{
    GENERATED_BODY()
public:
    FCellValue():RawText(TEXT("")),ival(0),_Type(EValueType::Empty) {}
    FCellValue(std::string sval) : RawText(sval.c_str())  ,_Type(EValueType::String)
    {
        try
        {

        }
        catch (...)
        {

        }
    }

    template<
        typename _Ty,
        typename std::enable_if<
        std::is_same_v<_Ty, FCellValue>||
        std::is_integral_v<_Ty> ||
        std::is_floating_point_v<_Ty> || 
        std::is_same_v<std::decay_t<_Ty>, std::string> ||
        std::is_same_v<std::decay_t<_Ty>, std::wstring> ||
#if __cplusplus >=201703L
        std::is_same_v<std::decay_t<T>, std::string_view> ||
        std::is_same_v<std::decay_t<T>, std::wstring_view> ||
#endif
        std::is_same_v<std::decay_t<_Ty>, FString> ||
        std::is_same_v<std::decay_t<_Ty>, const wchar_t*> ||
        std::is_same_v<std::decay_t<_Ty>, wchar_t*> ||
        std::is_same_v<std::decay_t<_Ty>, const char*> ||
        std::is_same_v<std::decay_t<_Ty>, char*> ||
        std::is_same_v<_Ty, FDateTime>>::type* = nullptr>
        FCellValue(_Ty _Val)
    {
        if constexpr (std::is_integral_v<_Ty> && std::is_same_v<_Ty, bool>) {
            _Type = EValueType::Boolean;
            ival = int(_Val);
            RawText = _Val? TEXT("true"):TEXT("false");
        }
        else if constexpr (std::is_integral_v<_Ty> && !std::is_same_v<_Ty, bool>) {
            _Type = EValueType::Integer;
            ival = int(_Val);
            RawText = FString::FromInt(_Val);
        }
        else if constexpr (std::is_same_v<std::decay_t<_Ty>, std::string> ||
            std::is_same_v<std::decay_t<_Ty>, std::wstring> ||
#if __cplusplus >=201703L
             std::is_same_v<std::decay_t<T>, std::string_view> ||
             std::is_same_v<std::decay_t<T>, std::wstring_view>||
#endif
            std::is_same_v<std::decay_t<_Ty>, const char*> ||
            std::is_same_v<std::decay_t<_Ty>, const wchar_t*> ||
            std::is_same_v<std::decay_t<_Ty>, FString> ||
            (std::is_same_v<std::decay_t<_Ty>, char*> && !std::is_same_v<_Ty, bool>)||
            (std::is_same_v<std::decay_t<_Ty>, wchar_t*> && !std::is_same_v<_Ty, bool>))
        {
            _Type = EValueType::String;
            ival = 0;
            RawText = _Val;
        }  
        else if constexpr (std::is_same_v<_Ty, FDateTime>) {
            _Type = EValueType::Float;
            fval = datetime_to_serial(_Val);
            RawText = _Val.ToString();
        } 
        else if constexpr (std::is_same_v<_Ty, FCellValue>)
        {
            _Type = _Val._Type;
            ival = _Val.ival;
            RawText = _Val.RawText;
        }
        else {
            static_assert(std::is_floating_point_v<_Ty>, "Invalid argument for constructing FCellValue object");
            _Type = EValueType::Float;
            fval = _Val;
            RawText = FString::SanitizeFloat(_Val);
        }
    }
    template<
        typename _Ty,
        typename std::enable_if<
        std::is_same_v<_Ty, FCellValue> ||
        std::is_integral_v<_Ty> ||
        std::is_floating_point_v<_Ty> ||
        std::is_same_v<std::decay_t<_Ty>, std::string> ||
        std::is_same_v<std::decay_t<_Ty>, std::wstring> ||
#if __cplusplus >=201703L
        std::is_same_v<std::decay_t<T>, std::string_view> ||
        std::is_same_v<std::decay_t<T>, std::wstring_view> ||
#endif
        std::is_same_v<std::decay_t<_Ty>, FString> ||
        std::is_same_v<std::decay_t<_Ty>, const wchar_t*> ||
        std::is_same_v<std::decay_t<_Ty>, wchar_t*> ||
        std::is_same_v<std::decay_t<_Ty>, const char*> ||
        std::is_same_v<std::decay_t<_Ty>, char*> ||
        std::is_same_v<_Ty, FDateTime>>::type* = nullptr>
        FCellValue& operator=(const _Ty& _Val)
    { 
        FCellValue temp(_Val);
        std::swap(*this, temp);
        return *this;
    }
      
    
    template<
        typename _Ty,
        typename std::enable_if<
        std::is_same_v<_Ty, FCellValue> ||
        std::is_integral_v<_Ty> ||
        std::is_floating_point_v<_Ty> ||
        std::is_same_v<std::decay_t<_Ty>, std::string> ||
        std::is_same_v<std::decay_t<_Ty>, std::wstring> ||
#if __cplusplus >=201703L
        std::is_same_v<std::decay_t<T>, std::string_view> ||
        std::is_same_v<std::decay_t<T>, std::wstring_view> ||
#endif
        std::is_same_v<std::decay_t<_Ty>, FString> ||
        std::is_same_v<std::decay_t<_Ty>, const wchar_t*> ||
        std::is_same_v<std::decay_t<_Ty>, wchar_t*> ||
        std::is_same_v<std::decay_t<_Ty>, const char*> ||
        std::is_same_v<std::decay_t<_Ty>, char*> ||
        std::is_same_v<_Ty, FDateTime>>::type* = nullptr>
        operator _Ty () const
    {

        if constexpr (std::is_integral_v<_Ty> && std::is_same_v<_Ty, bool>)
            return (bool)(ival);
        else  if constexpr (std::is_integral_v<_Ty> && !std::is_same_v<_Ty, bool>)
            return ival;
        else if constexpr (std::is_floating_point_v<_Ty>)
            return fval;
        else if constexpr (std::is_same_v<std::decay_t<_Ty>, std::string> ||
#if __cplusplus >=201703L
            std::is_same_v<std::decay_t<T>, std::string_view> ||
#endif
            std::is_same_v<std::decay_t<_Ty>, const char*> ||
            (std::is_same_v<std::decay_t<_Ty>, char*> && !std::is_same_v<_Ty, bool>))
            return std::string(*RawText);
        else if constexpr (std::is_same_v<std::decay_t<_Ty>, std::wstring> ||
#if __cplusplus >=201703L
            std::is_same_v<std::decay_t<T>, std::wstring_view> ||
#endif
            std::is_same_v<std::decay_t<_Ty>, FString> ||
            std::is_same_v<std::decay_t<_Ty>, const wchar_t*> ||
            (std::is_same_v<std::decay_t<_Ty>, wchar_t*> && !std::is_same_v<_Ty, bool>))
            return *RawText;
        else if constexpr (std::is_same_v<std::decay_t<_Ty>, FString>)
        {
            return RawText;
        }
        else if constexpr (std::is_same_v<_Ty, FDateTime>)
        { 
            return serial_to_datetime(fval);
        }
        else if constexpr (std::is_same_v<_Ty, FCellValue>)
        {
            return *this;
        }

    }
  
    FString ToString() const
    {
        return RawText;
    }

    void clear()
    {
        _Type = EValueType::Empty;
        ival = 0;
        RawText.Empty();
    }
      
    EValueType type()const
    {
        return _Type;
    }


public:
   

    static double datetime_to_serial(FDateTime& dt);

    static double unix_to_serial(time_t unixtime)
    {
        // There are 86400 seconds in a day
        // There are 25569 days between 1/1/1970 and 30/12/1899 (the epoch used by Excel)
        return static_cast<double>(unixtime) / 86400 + 25569;
    }
     
    static FDateTime serial_to_datetime(double serial);

protected:
    FString RawText;
    union
    {
        double fval;
        int64 ival;
    };
    EValueType _Type ;                            
	friend class UCell;
    friend class UFreeExcelLibrary;
};

 