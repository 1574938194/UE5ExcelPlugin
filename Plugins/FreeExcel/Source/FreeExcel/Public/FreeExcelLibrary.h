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
#include "Kismet/BlueprintFunctionLibrary.h"
#include "CellReference.h"
#include "CellRange.h"
#include "CellValue.h"
#include "FreeExcelLibrary.generated.h"


class UDocument;
class USheet;
class UCell;
 
/**
 *
 */
UCLASS(Meta=(BlueprintThreadSafe))
class FREEEXCEL_API UFreeExcelLibrary : public UBlueprintFunctionLibrary
{
	GENERATED_BODY()

public:


	UFUNCTION(BlueprintPure, Category = "FreeExcel", meta = (NativeBreakFunc))
		static void BreakCellReference(FCellReference Ref, int& Row, int& Col);

	UFUNCTION(BlueprintCallable, Category = "FreeExcel")
		static void Offset(FCellReference Ref, int32 row, int32 col);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Equal (CellReference)", CompactNodeTitle = "==", ScriptMethod = "Equals", ScriptOperator = "==", Keywords = "== equal"), Category = "FreeExcel")
		static bool Equal_CellReference(FCellReference A, FCellReference B);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Not Equal (CellReference)", CompactNodeTitle = "!=", ScriptMethod = "NotEqual", ScriptOperator = "==", Keywords = "== not equal"), Category = "FreeExcel")
		static bool NotEqual_CellReference(FCellReference A, FCellReference B);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Equal (CellRange)", CompactNodeTitle = "==", ScriptMethod = "Equals", ScriptOperator = "==", Keywords = "== equal"), Category = "FreeExcel")
		static bool Equal_CellRange(FCellRange A, FCellRange B);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Not Equal (CellRange)", CompactNodeTitle = "!=", ScriptMethod = "NotEqual", ScriptOperator = "==", Keywords = "== not equal"), Category = "FreeExcel")
		static bool NotEqual_CellRange(FCellRange A, FCellRange B);

	// The constructor creates a new XLCellReference from a string, e.g. 'A1'. If there's no input,
	// the default reference will be cell A1.
	UFUNCTION(BlueprintPure, meta = (NativeMakeFunc), Category = "FreeExcel")
		static FCellReference MakeCellReferenceWithString(FString ref);

	// Creates a new FCellReference from a given row and column number, e.g. 1,1 (=A1)
	UFUNCTION(BlueprintPure, meta = (NativeMakeFunc), Category = "FreeExcel")
		static FCellReference MakeCellReference(int32 row, int32 col);


	UFUNCTION(BlueprintPure, meta = (NativeMakeFunc), Category = "FreeExcel")
		static FCellRange MakeCellRangeWithString(FString ref);

	UFUNCTION(BlueprintPure, meta = (NativeMakeFunc), Category = "FreeExcel")
		static FCellRange MakeCellRange(int32 MinRow, int32 MinCol, int32 MaxRow, int32 MaxCol);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static FCellRange MakeCellRange_Internal(USheet* sheet, const FCellRange& range);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "To String (CellReference)"), Category = "FreeExcel")
		static FString ToString_CellReference(FCellReference ref);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Make CellValue (Bool)", NativeMakeFunc), Category = "FreeExcel")
		static FCellValue MakeCellValue_Bool(bool val);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Make CellValue (Int)", NativeMakeFunc), Category = "FreeExcel")
		static FCellValue  MakeCellValue_Int(int32 val);


	UFUNCTION(BlueprintPure, meta = (DisplayName = "Make CellValue (String)", NativeMakeFunc), Category = "FreeExcel")
		static FCellValue  MakeCellValue_String(FString val);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Make CellValue (DateTime)", NativeMakeFunc), Category = "FreeExcel")
		static FCellValue  MakeCellValue_DateTime(FDateTime val);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "To Bool (CellValue)"), Category = "FreeExcel")
		static	bool  ToBool_CellValue(const FCellValue& val);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "To Int (CellValue)"), Category = "FreeExcel")
		static	int32  ToInt_CellValue(const FCellValue& val);



	UFUNCTION(BlueprintPure, meta = (DisplayName = "To String (CellValue)"), Category = "FreeExcel")
		static	FString  ToString_CellValue(const FCellValue& val);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "To DateTime (CellValue)"), Category = "FreeExcel")
		static	FDateTime  ToDateTime_CellValue(const FCellValue& val);

	UFUNCTION(BlueprintCallable, meta = (DisplayName = "Clear (CellValue)"), Category = "FreeExcel")
		static	void  Clear_CellValue(FCellValue& val);

	UFUNCTION(BlueprintPure, meta = (DisplayName = "Get Type (CellValue)"), Category = "FreeExcel")
		static	EValueType  Type_CellValue(const FCellValue& val);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void CellIterator_Forward(const FCellIterator& Target);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void Range_Begin(const FCellRange& Target, FCellIterator& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void Range_End(const FCellRange& Target, FCellIterator& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void CellIterator_CellReference(const FCellIterator& Target, FCellReference& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void CellIterator_Cell(const FCellIterator& Target, UCell*& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void CellIterator_NotEqual(const FCellIterator& A, const FCellIterator& B, bool& ReturnValue);

	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "Target,ReturnValue", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void ToTemplateArray(const int32& Target, TArray<int32>& ReturnValue);
	static void Generic_ToTemplateArray(FObjectProperty* SelfProperty, void* Self, FArrayProperty* RetProperty, void* Ret);
	DECLARE_FUNCTION(execToTemplateArray)
	{
		Stack.StepCompiledIn<FStructProperty>(NULL);
		FObjectProperty* SelfProperty = CastField<FObjectProperty>(Stack.MostRecentProperty);
		void* SelfPtr = Stack.MostRecentPropertyAddress;

		Stack.StepCompiledInRef<FStructProperty, void*>(NULL);
		FArrayProperty* RetProperty = CastField<FArrayProperty>(Stack.MostRecentProperty);
		void* RetPtr = Stack.MostRecentPropertyAddress;
		if (!RetProperty)
		{
			Stack.bArrayContextFailed = true;
			return;
		}

		P_FINISH;
		P_NATIVE_BEGIN;
		P_THIS->Generic_ToTemplateArray(SelfProperty, SelfPtr, RetProperty, RetPtr);
		P_NATIVE_END;
	}

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_DocInt(UDocument* Target, const FCellReference& Ref, const int32& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_DocBool(UDocument* Target, const FCellReference& Ref, const bool& Value);


	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_DocString(UDocument* Target, const FCellReference& Ref, const FString& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_DocDateTime(UDocument* Target, const FCellReference& Ref, const FDateTime& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_DocCellValue(UDocument* Target, const FCellReference& Ref, const FCellValue& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_SheetInt(USheet* Target, const FCellReference& Ref, const int32& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_SheetBool(USheet* Target, const FCellReference& Ref, const bool& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_SheetString(USheet* Target, const FCellReference& Ref, const FString& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_SheetDateTime(USheet* Target, const FCellReference& Ref, const FDateTime& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_SheetCellValue(USheet* Target, const FCellReference& Ref, const FCellValue& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellInt(UCell* Target, const int32& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellBool(UCell* Target, const bool& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellString(UCell* Target, const FString& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellDateTime(UCell* Target, const FDateTime& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellCellValue(UCell* Target, const FCellValue& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellValueInt(const FCellValue& Target, const int32& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellValueBool(const FCellValue& Target, const bool& Value);


	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellValueString(const FCellValue& Target, const FString& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellValueDateTime(const FCellValue& Target, const FDateTime& Value);

	UFUNCTION(BlueprintCallable, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellValueCellValue(const FCellValue& Target, const FCellValue& Value);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_DocInt(UDocument* Target, const FCellReference& Ref, int32& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_DocBool(UDocument* Target, const FCellReference& Ref, bool& ReturnValue);



	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_DocString(UDocument* Target, const FCellReference& Ref, FString& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_DocDateTime(UDocument* Target, const FCellReference& Ref, FDateTime& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_DocCellValue(UDocument* Target, const FCellReference& Ref, FCellValue& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_SheetInt(USheet* Target, const FCellReference& Ref, int32& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_SheetBool(USheet* Target, const FCellReference& Ref, bool& ReturnValue);


	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_SheetString(USheet* Target, const FCellReference& Ref, FString& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_SheetDateTime(USheet* Target, const FCellReference& Ref, FDateTime& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_SheetCellValue(USheet* Target, const FCellReference& Ref, FCellValue& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellInt(UCell* Target, int32& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellBool(UCell* Target, bool& ReturnValue);


	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellString(UCell* Target, FString& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellDateTime(UCell* Target, FDateTime& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellCellValue(UCell* Target, FCellValue& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellValueInt(const FCellValue& Target, int32& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellValueBool(const FCellValue& Target, bool& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellValueString(const FCellValue& Target, FString& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellValueDateTime(const FCellValue& Target, FDateTime& ReturnValue);

	UFUNCTION(BlueprintPure, meta = (BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellValueCellValue(const FCellValue& Target, FCellValue& ReturnValue);



	UFUNCTION(BlueprintCallable, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_DocFloat(UDocument* Target, const FCellReference& Ref, const int32& Value) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_SetCellValue_DocFloat(UDocument* Target, const FCellReference& Ref, const double& Value);
#else
	static void Generic_SetCellValue_DocFloat(UDocument* Target, const FCellReference& Ref, const float& Value);
#endif
	DECLARE_FUNCTION(execSetCellValue_DocFloat)
	{
		P_GET_OBJECT(UDocument, Z_Param_Target);
		P_GET_STRUCT_REF(FCellReference, Z_Param_Out_Ref);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_SetCellValue_DocFloat(Z_Param_Target, Z_Param_Out_Ref, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}

	UFUNCTION(BlueprintCallable, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_SheetFloat(USheet* Target, const FCellReference& Ref, const int32& Value) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_SetCellValue_SheetFloat(USheet* Target, const FCellReference& Ref, const double& Value);
#else
	static void Generic_SetCellValue_SheetFloat(USheet* Target, const FCellReference& Ref, const float& Value);
#endif
	DECLARE_FUNCTION(execSetCellValue_SheetFloat)
	{
		P_GET_OBJECT(USheet, Z_Param_Target);
		P_GET_STRUCT_REF(FCellReference, Z_Param_Out_Ref);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_SetCellValue_SheetFloat(Z_Param_Target, Z_Param_Out_Ref, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}

	UFUNCTION(BlueprintCallable, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellFloat(UCell* Target, const int32& Value) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_SetCellValue_CellFloat(UCell* Target, const double& Value);
#else
	static void Generic_SetCellValue_CellFloat(UCell* Target, const float& Value);
#endif
	DECLARE_FUNCTION(execSetCellValue_CellFloat)
	{
		P_GET_OBJECT(UCell, Z_Param_Target);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_SetCellValue_CellFloat(Z_Param_Target, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}
	
	UFUNCTION(BlueprintCallable, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void SetCellValue_CellValueFloat(const FCellValue& Target, const int32& Value) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_SetCellValue_CellValueFloat(const FCellValue& Target,const double& Value);
#else
	static void Generic_SetCellValue_CellValueFloat(const FCellValue& Target, const float& Value);
#endif
	DECLARE_FUNCTION(execSetCellValue_CellValueFloat)
	{
		P_GET_STRUCT_REF(FCellValue, Z_Param_Out_Target);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_SetCellValue_CellValueFloat(Z_Param_Out_Target, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}

	
	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_DocFloat(UDocument* Target, const FCellReference& Ref, int32& ReturnValue) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_GetCellValue_DocFloat(UDocument* Target, const FCellReference& Ref, double& Value);
#else
	static void Generic_GetCellValue_DocFloat(UDocument* Target, const FCellReference& Ref, float& Value);
#endif
	DECLARE_FUNCTION(execGetCellValue_DocFloat)
	{
		P_GET_OBJECT(UDocument, Z_Param_Target);
		P_GET_STRUCT_REF(FCellReference, Z_Param_Out_Ref);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_GetCellValue_DocFloat(Z_Param_Target, Z_Param_Out_Ref, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}
	
	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_SheetFloat(USheet* Target, const FCellReference& Ref, int32& ReturnValue) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_GetCellValue_SheetFloat(USheet* Target, const FCellReference& Ref, double& Value);
#else
	static void Generic_GetCellValue_SheetFloat(USheet* Target, const FCellReference& Ref, float& Value);
#endif
	DECLARE_FUNCTION(execGetCellValue_SheetFloat)
	{
		P_GET_OBJECT(USheet, Z_Param_Target);
		P_GET_STRUCT_REF(FCellReference, Z_Param_Out_Ref);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_GetCellValue_SheetFloat(Z_Param_Target, Z_Param_Out_Ref, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}
	
	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "ReturnValue", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellFloat(UCell* Target, int32& ReturnValue) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_GetCellValue_CellFloat(UCell* Target, double& Value);
#else
	static void Generic_GetCellValue_CellFloat(UCell* Target,  float& Value);
#endif
	DECLARE_FUNCTION(execGetCellValue_CellFloat)
	{
		P_GET_OBJECT(UCell, Z_Param_Target);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_GetCellValue_CellFloat(Z_Param_Target, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}
	
	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "Value", BlueprintInternalUseOnly = "true"), Category = "FreeExcel")
		static void GetCellValue_CellValueFloat(const FCellValue& Target, int32& ReturnValue) {}
#if ENGINE_MAJOR_VERSION >=5
	static void Generic_GetCellValue_CellValueFloat(const FCellValue& Target, double& Value);
#else
	static void Generic_GetCellValue_CellValueFloat(const FCellValue& Target, float& Value);
#endif
	DECLARE_FUNCTION(execGetCellValue_CellValueFloat)
	{
		P_GET_STRUCT_REF(FCellValue, Z_Param_Target);
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_FINISH;
		P_NATIVE_BEGIN;
		UFreeExcelLibrary::Generic_GetCellValue_CellValueFloat(Z_Param_Target, Z_Param_Out_ReturnValue);
		P_NATIVE_END;
	}

	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "Value", DisplayName = "Make CellValue (Float)", NativeMakeFunc), Category = "FreeExcel")
		static FCellValue  MakeCellValue_Float(const int32& Value) { return{}; }
#if ENGINE_MAJOR_VERSION >=5
	static FCellValue Generic_MakeCellValue_Float(double Value);
#else
	static FCellValue Generic_MakeCellValue_Float(float Value);
#endif
	DECLARE_FUNCTION(execMakeCellValue_Float)
	{ 
#if ENGINE_MAJOR_VERSION >=5
		P_GET_PROPERTY_REF(FDoubleProperty, Z_Param_Out_ReturnValue);
#else 
		P_GET_PROPERTY_REF(FFloatProperty, Z_Param_Out_ReturnValue);
#endif
		P_GET_PROPERTY(FDoubleProperty, Z_Param_val);
		P_FINISH;
		P_NATIVE_BEGIN;
		*(FCellValue*)Z_Param__Result = UFreeExcelLibrary::Generic_MakeCellValue_Float(Z_Param_val);
		P_NATIVE_END;
	}


	UFUNCTION(BlueprintPure, CustomThunk, meta = (CustomStructureParam = "ReturnValue", DisplayName = "To Float (CellValue)"), Category = "FreeExcel")
		static	void  ToFloat_CellValue(const FCellValue& Target, int32& ReturnValue) {}
#if ENGINE_MAJOR_VERSION >=5
	static double Generic_ToFloat_CellValue(const FCellValue& Target);
#else
	static float Generic_ToFloat_CellValue(const FCellValue& Target);
#endif
	DECLARE_FUNCTION(execToFloat_CellValue)
	{ 
		P_GET_STRUCT_REF(FCellValue, Z_Param_Out_val);
		P_FINISH;
		P_NATIVE_BEGIN;
		
#if ENGINE_MAJOR_VERSION >=5
		* (double*)Z_Param__Result = UFreeExcelLibrary::Generic_ToFloat_CellValue(Z_Param_Out_val);
#else 
		* (float*)Z_Param__Result = UFreeExcelLibrary::Generic_ToFloat_CellValue(Z_Param_Out_val);
#endif
		P_NATIVE_END;
	}

	 
};