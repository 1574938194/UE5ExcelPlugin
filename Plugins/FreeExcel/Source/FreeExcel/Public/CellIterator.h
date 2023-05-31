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
#include "CellIterator.generated.h"
 
struct FCellRange;
class UFreeExcelLibrary;

USTRUCT(BlueprintType, Meta = (HasNativeMake, HasNativeBreak))
struct FREEEXCEL_API FCellIterator
{
    GENERATED_BODY()
public:
    using iterator_category = std::forward_iterator_tag;
    using value_type = UCell;
    using difference_type = int64_t;
    using pointer = UCell*;
    using reference = UCell&;
  
    FCellIterator() = default;
    FCellIterator(const FCellIterator& right)
        : dataNode(right.dataNode),
        Current(right.Current) ,
        min(right.min),
        max(right.max),
        reached(right.reached)
    {  }
    FCellIterator& operator =(const FCellIterator& right)
    {
        if (&right != this) {
            dataNode = right.dataNode; 
            min = right.min;
            max = right.max;
            Current = right.Current;
            reached = right.reached;
        }
        return *this;
    }

    FCellIterator& operator++();
     
    FCellIterator operator++(int);    // NOLINT

    pointer get()const;

    pointer operator->()const
    {
        return get();
    }

    void next_row();

    int32 rowCellCount() const
    {
        return FCellReference(dataNode.parent().last_child().attribute("r").value()).Col;
    }

    bool operator==(const FCellIterator& right)const;

    bool operator!=(const FCellIterator& right) const
    { 
        return !(*this == right);
    }
     
    
protected:
    XMLNode dataNode;
    FCellReference Current;
    FCellReference min;   
    FCellReference max;    
    bool reached{ false };

    friend struct FCellRange;
    friend class UFreeExcelLibrary;
};

  