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
#include "CellReference.h"
#include "Cell.h"
#include "CellIterator.h"
#include "CellRange.generated.h"
  
class UFreeExcelLibrary;
/**
 * 
 */
USTRUCT(BlueprintType, Meta = (HasNativeMake, HasNativeBreak))
struct FREEEXCEL_API FCellRange
{ 
    GENERATED_BODY()
public: 
    FCellRange() = default;
    FCellRange(FCellReference min, FCellReference max)
        :Min(min),Max(max)
    {

    } 
    FCellRange(const FCellRange& right)
        :dataNode(right.dataNode),Min(right.Min),Max(right.Max){  }
    FCellRange& operator=(const FCellRange& right)
    {
        if (&right != this)
        {
            dataNode = right.dataNode;
            Min = right.Min;
            Max = right.Max;
        }
        return *this;
    }

    inline int32 rows() const
    {
        return Max.Row + 1 - Min.Row;
    }

    inline int16 cols()
    {
        return Max.Col + 1 - Min.Col;
    }

    FCellIterator begin() const;

    FCellIterator end() const;


   
protected:
    XMLNode dataNode;   
    /**< The cell reference of the first cell in the range */
    UPROPERTY(EditAnywhere)
        FCellReference Min;
    /**< The cell reference of the last cell in the range */
    UPROPERTY(EditAnywhere)
        FCellReference Max;

    friend class USheet; 
    friend struct FCellIterator;
    friend class UFreeExcelLibrary;
};
 