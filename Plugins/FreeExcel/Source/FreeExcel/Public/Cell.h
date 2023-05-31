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
#include "CellValue.h"
#include "XmlData.h"
#include "Cell.generated.h"

class USheet;

UENUM(BlueprintType)
enum class ECellType :uint8 {
	Empty UMETA(DisplayName = "Empty"),
	Boolean UMETA(DisplayName = "Boolean"),
	Integer UMETA(DisplayName = "Integer"),
	Float UMETA(DisplayName = "Float"),
	Error,
	String UMETA(DisplayName = "String"),
	Formula UMETA(DisplayName = "Formula")
};

/**
 * 
 */
UCLASS(BlueprintType)
class FREEEXCEL_API UCell : public UObject
{
	GENERATED_BODY()
	
public:

    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    FCellReference GetReference()const;
	 
    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    bool HasFormula() const ;
	  
    UFUNCTION(BlueprintCallable, Category = "FreeExcel")
    void SetFormula(FString formula);

    UFUNCTION(BlueprintPure, Category = "FreeExcel")
		FCellValue Value()const
	{
		return to_value(sheet, cellNode);
	}

	static FCellValue to_value(const USheet* sheet, const XMLNode& cellNode);


    UFUNCTION(BlueprintCallable, meta = (DisplayName = "Set Cell Value (Cell)"), Category = "FreeExcel")
    void SetCellValue(const FCellValue& val)
	{
		switch (val.type())
		{
		case EValueType::Empty:
			Clear();
			break;
		case EValueType::Boolean:
			SetBool((bool)val.ival);
			break;
		case EValueType::Integer:
			SetInt(val.ival);
			break;
		case EValueType::Float:
			Generic_SetFloat(val.fval);
			break;
		case EValueType::Error:
			break;
		case EValueType::String:
			SetString(val.RawText);
			break;
		default:
			break;
		}
	}

    UFUNCTION(BlueprintPure, Category = "FreeExcel")
    bool HasValue() const;

    UFUNCTION(BlueprintCallable, meta = (DisplayName = "Clear (Cell)"), Category = "FreeExcel")
    void Clear();
	 
	UFUNCTION(BlueprintCallable, meta = (DisplayName = "Set Bool (Cell)"), Category = "FreeExcel")
		void SetBool(bool val);

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "Set Int (Cell)"), Category = "FreeExcel")
		void SetInt(int32 val);

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "Set String (Cell)"), Category = "FreeExcel")
		void SetString(FString val);

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "Set DateTime (Cell)"), Category = "FreeExcel")
    void SetDateTime(FDateTime val)
    {
		Generic_SetFloat(FCellValue::datetime_to_serial(val));
    }

	UFUNCTION(BlueprintCallable, CustomThunk, meta = (CustomStructureParam = "Value",DisplayName = "Set Float (Cell)"), Category = "FreeExcel")
		void SetFloat(const int32& Value) {}
#if ENGINE_MAJOR_VERSION >=5
	 void Generic_SetFloat(double val);
#else
	void Generic_SetFloat(float val);
#endif
	DECLARE_FUNCTION(execSetFloat)
	{
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY(FDoubleProperty, Z_Param_val);
#else 
		P_GET_PROPERTY(FFloatProperty, Z_Param_val);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		P_THIS->Generic_SetFloat(Z_Param_val);
		P_NATIVE_END;
	}


	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "Value", DisplayName = "To Float (Cell)"), Category = "FreeExcel")
		void ToFloat(int32& ReturnValue)const {}
#if ENGINE_MAJOR_VERSION >=5
	void Generic_ToFloat(double& val)const;
#else
	void Generic_ToFloat(float& val)const;
#endif
	DECLARE_FUNCTION(execToFloat)
	{
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_val);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_val);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		P_THIS->Generic_ToFloat(Z_Param_val);
		P_NATIVE_END;
	}
#if ENGINE_MAJOR_VERSION >=5
	double ToFloat()const 
	{ 
		double val = 0.0;
		Generic_ToFloat(val);
		return val;
	}
#else 
	float ToFloat()const
	{
		float val = 0.0f;
		Generic_ToFloat(val);
		return val;
	}
#endif
	UFUNCTION(BlueprintPure, meta = (DisplayName = "To Bool (Cell)"), Category = "FreeExcel")
		bool ToBool()const;

	UFUNCTION(BlueprintPure, meta = (DisplayName = "To Int (Cell)"), Category = "FreeExcel")
		int32 ToInt()const;



	UFUNCTION(BlueprintPure, meta = (DisplayName = "To String (Cell)"), Category = "FreeExcel")
		FString ToString()const;
	 
	UFUNCTION(BlueprintPure, meta = (DisplayName = "To DateTime (Cell)"), Category = "FreeExcel")
		FDateTime ToDateTime()const
	{
#if ENGINE_MAJOR_VERSION >=5
		double val = 0.0;
#else
		float val = 0.0;
#endif
		Generic_ToFloat(val);
		return FCellValue::serial_to_datetime(val);
	}
  
	UFUNCTION(BlueprintPure, meta = (DisplayName = "Get Type (Cell)"), Category = "FreeExcel")
		ECellType Type()const;
  
    operator bool() const
    {
        return (bool)cellNode;
    }
 
	void offset(int32 rowOffset, int32 colOffset);
   
    //void setError()
    //{
    //    // ===== Check that the m_cellNode is valid.
    //    assert(m_cellNode);              // NOLINT
    //    assert(!m_cellNode->empty());    // NOLINT

    //    // ===== If the cell node doesn't have a type attribute, create it.
    //    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");

    //    // ===== Set the type to "e", i.e. error
    //    m_cellNode->attribute("t").set_value("e");

    //    // ===== Remove the value node, if it exists
    //    m_cellNode->remove_child("v");

    //    // ===== Disable space preservation (only relevant for strings).
    //    m_cellNode->remove_attribute(" xml:space");

    //    // ===== Remove the value node.
    //    m_cellNode->remove_child("v");
    //}
  
   
protected:
    USheet* sheet;
    XMLNode cellNode;

    friend class USheet;
    friend struct FCellIterator;  
	friend class UCell;
	friend class UFreeExcelLibrary;
};
