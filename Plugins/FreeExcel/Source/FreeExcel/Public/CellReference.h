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
#include "UObject/NoExportTypes.h" 
#include <string>
#include "CellReference.generated.h"

 
class UFreeExcelLibrary;
/**
 * 
 */
USTRUCT(BlueprintType,Meta = (HasNativeMake, HasNativeBreak))
struct FREEEXCEL_API FCellReference 
{
	GENERATED_BODY()
public:
	FCellReference() {}
	FCellReference(const FCellReference&) = default;
	FCellReference(int32 r, int32 c) :Row(r),Col(c){ }
	FCellReference(FString ref);
	FCellReference(const std::string& ref) :FCellReference(FString(ref.c_str())){}
	FCellReference(const char* ref) :FCellReference(FString(ref)) {}
	FCellReference(const std::wstring& ref) :FCellReference(FString(ref.c_str())) {}
	FCellReference(const wchar_t* ref) :FCellReference(FString(ref)) {}

	UPROPERTY(EditAnywhere)
		int32 Row = { 1 };

	UPROPERTY(EditAnywhere)
		int32 Col {1};
	 
	FString ToString() const
	{
		return FString(to_string().c_str());
	}
	std::string to_string() const;

	inline static bool address_is_valid(uint32_t row, uint16_t column)  {
		return !(row < 1 || row >  max_rows || column < 1 || column >  max_cols);
	}

	inline bool operator==(const FCellReference& right) const
	{
		return Row == right.Row && Col == right.Col;
	}
	inline bool operator!=(const FCellReference& right)const
	{
		return Row != right.Row || Col != right.Col;
	}
protected: 
	static constexpr uint8_t alphabetSize = 26;
	static constexpr uint8_t asciiOffset = 64;
	static constexpr uint16_t max_cols = 16384;
	static constexpr uint32_t max_rows = 1048576;


	friend class USheet;
	friend class UFreeExcelLibrary;
};
