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
#include "Cell.h"
#include "CellReference.h"
#include "CellRange.h" 
#include "XmlData.h"
#include "Sheet.generated.h"

 
class UDocument;
class UFreeExcelLibrary;
/**
 * 
 */
UCLASS(BlueprintType)
class FREEEXCEL_API USheet : public UObject
{
	GENERATED_BODY()
public:
     
    // Get the title of sheet
    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    FString Name() const;

    // Set the title for the sheet
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
    void SetName(FString sheetName);
   
    // Get a pointer to the Cell object for the given cell reference.
    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    UCell* Cell( FCellReference ref) const;
   
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
        int32 GetRowCellCount(int32 row)const;

    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
     TArray<FCellValue> GetRowData(int32 row)const;

    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
    TArray<float> GetRowFloatData(int32 row , bool skip0= true)const;
   
  
    void reset(const _XML_DATA* xmlData);

    bool isSelected() const
    {
        return data_->xmlDocument().first_child().child("sheetViews").first_child().attribute("tabSelected").as_bool();
    }
 
    void setSelected(bool selected)
    { 
        unsigned int value = (selected ? 1 : 0);
        data_->xmlDocument().first_child().child("sheetViews").first_child().attribute("tabSelected").set_value(value);

    }
    // Get an CellReference to the last (bottom right) cell in the worksheet.
    UFUNCTION(BlueprintPure, Category = "FreeExcel")
        FCellReference  LastCell() const noexcept;
 
    // update cell reference sheet name
    void updateSheetName(const std::string& oldName, const std::string& newName);
protected:
    UDocument* doc_;
    _XML_DATA* data_;
    std::string id_;

    friend class UDocument; 
    friend class UFreeExcelLibrary;
    friend class ADemo1;
    friend class UCell;
};
