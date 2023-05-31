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


#include "K2Node_GetCellValue.h"
#include "EdGraphSchema_K2.h"
#include "BlueprintActionDatabaseRegistrar.h"
#include "BlueprintNodeSpawner.h"
#include "EditorCategoryUtils.h"
#include "K2Node_CallFunction.h"
#include "KismetCompiler.h"
#include "FreeExcelLibrary.h"
#include "K2Node_Self.h"
#include "Document.h"
#include "Sheet.h"
#include "Cell.h"
#include "CellValue.h"

#define LOCTEXT_NAMESPACE "K2Node_GetCellValue"

namespace ArraySortHelper
{
	const FName TargetArrayPinName = "TargetArray";
	const FName PropertyNamePinName = "PropertyName";
	const FName AscendingPinName = "bAscending";
}




UK2Node_GetCellValue::UK2Node_GetCellValue(const FObjectInitializer& ObjectInitializer)
	: Super(ObjectInitializer)
{
	NodeTooltip = LOCTEXT("GetCellValue_NodeTooltip", "Get cell value from given context");
	InputCheck = LOCTEXT("GetCellValue_InputCheck", "Cannot input this type!");
}

void UK2Node_GetCellValue::SetPinToolTip(UEdGraphPin& MutatablePin, const FText& PinDescription) const
{
	MutatablePin.PinToolTip = UEdGraphSchema_K2::TypeToText(MutatablePin.PinType).ToString();

	UEdGraphSchema_K2 const* const K2Schema = Cast<const UEdGraphSchema_K2>(GetSchema());
	if (K2Schema != nullptr)
	{
		MutatablePin.PinToolTip += TEXT(" ");
		MutatablePin.PinToolTip += K2Schema->GetPinDisplayName(&MutatablePin).ToString();
	}

	MutatablePin.PinToolTip += FString(TEXT("\n")) + PinDescription.ToString();
}

void UK2Node_GetCellValue::AllocateDefaultPins()
{
	  

	const UEdGraphSchema_K2* K2Schema = GetDefault<UEdGraphSchema_K2>();

	// Add execution pins
	/*CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Exec, UEdGraphSchema_K2::PN_Execute);
	CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Exec, UEdGraphSchema_K2::PN_Then);

	*/
	 
	auto SelfPin = CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Wildcard, UEdGraphSchema_K2::PN_Self);
	SetPinToolTip(*SelfPin, LOCTEXT("GetCellValue_SelfPinDesc", "[Document,Sheet,Cell,CellValue] In"));

	auto RefPin = CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Struct,FCellReference::StaticStruct(),  "Ref");
	SetPinToolTip(*RefPin, LOCTEXT("GetCellValue_RefPinDesc", "Cell Reference In"));

	auto ValPin = CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Wildcard,  "ReturnValue");
	SetPinToolTip(*ValPin, LOCTEXT("GetCellValue_ReturnValuePinDesc", "[Integer,Boolean,Float,String,DateTime,CellValue] Out"));

	Super::AllocateDefaultPins();
}

 

 

void UK2Node_GetCellValue::GetMenuActions(FBlueprintActionDatabaseRegistrar& ActionRegistrar) const
{
	// actions get registered under specific object-keys; the idea is that 
	// actions might have to be updated (or deleted) if their object-key is  
	// mutated (or removed)... here we use the node's class (so if the node 
	// type disappears, then the action should go with it)
	UClass* ActionKey = GetClass();
	// to keep from needlessly instantiating a UBlueprintNodeSpawner, first   
	// check to make sure that the registrar is looking for actions of this type
	// (could be regenerating actions for a specific asset, and therefore the 
	// registrar would only accept actions corresponding to that asset)
	if (ActionRegistrar.IsOpenForRegistration(ActionKey))
	{
		UBlueprintNodeSpawner* NodeSpawner = UBlueprintNodeSpawner::Create(GetClass());
		check(NodeSpawner != nullptr);

		ActionRegistrar.AddBlueprintAction(ActionKey, NodeSpawner);
	}
}

FText UK2Node_GetCellValue::GetMenuCategory() const
{
	return LOCTEXT("GetCellValue_GetMenuCategory", "FreeExcel");
}


void UK2Node_GetCellValue::PinDefaultValueChanged(UEdGraphPin* ChangedPin)
{
	TArray<FName> ls = { TEXT("Self"), TEXT("ReturnValue") };
	if (ls.Contains(ChangedPin->GetFName()))
	{
		PropagatePinType(ChangedPin);
	}
}
 

FText UK2Node_GetCellValue::GetTooltipText() const
{
	return NodeTooltip;
}


FText UK2Node_GetCellValue::GetNodeTitle(ENodeTitleType::Type TitleType) const
{
	return LOCTEXT("GetCellValue_NodeTitle", "Get Cell Value");
}

void UK2Node_GetCellValue::ExpandNode(class FKismetCompilerContext& CompilerContext, UEdGraph* SourceGraph)
{
	Super::ExpandNode(CompilerContext, SourceGraph);

	if (FindPin(TEXT("ReturnValue"))->PinType.PinCategory== UEdGraphSchema_K2::PC_Wildcard )
	{
		CompilerContext.MessageLog.Error(*NSLOCTEXT("K2Node", "GetCellValue_ExpandError", "GetCellValue Expand error in \"ReturnValue\"").ToString());
		return;
	}
	const UEdGraphSchema_K2* Schema = CompilerContext.GetSchema();

	UK2Node_CallFunction* CallFuncNode = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(this, SourceGraph);
	//CallFuncNode->SetFromFunction(Function);
	const FName DstFunctionName = GetDestCall();
	CallFuncNode->FunctionReference.SetExternalMember(DstFunctionName, UFreeExcelLibrary::StaticClass());
	CallFuncNode->AllocateDefaultPins();
	UEdGraphPin* CallFuncSelfPin = Schema->FindSelfPin(*CallFuncNode, EGPD_Input);
 
	auto _Self = FindPin(UEdGraphSchema_K2::PN_Self);
	if (_Self->PinType.PinCategory == UEdGraphSchema_K2::PC_Wildcard)
	{
		UK2Node_Self* self = CompilerContext.SpawnIntermediateNode<UK2Node_Self>(this, SourceGraph);
		self->AllocateDefaultPins();
		auto selfPin = self->Pins[0];

		UEdGraphPin* DestPin = CallFuncNode->FindPin(TEXT("Target"));
		if (DestPin != NULL)
		{
			DestPin->PinType = selfPin->PinType;
			DestPin->PinType.PinSubCategoryObject = selfPin->PinType.PinSubCategoryObject;
			CompilerContext.MovePinLinksToIntermediate(*selfPin, *DestPin);
		}
	}
	else
	{
		UEdGraphPin* DestPin = CallFuncNode->FindPin(TEXT("Target"));
		if (DestPin != NULL)
		{
			DestPin->PinType = _Self->PinType;
			DestPin->PinType.PinSubCategoryObject = _Self->PinType.PinSubCategoryObject;
			CompilerContext.MovePinLinksToIntermediate(*_Self, *DestPin);
		}
	}

	// Now move the rest of the connections (including exec connections...)
	for (int32 SrcPinIdx = 0; SrcPinIdx < Pins.Num(); SrcPinIdx++)
	{
		UEdGraphPin* SrcPin = Pins[SrcPinIdx];
		if (SrcPin != NULL && SrcPin->PinName!= UEdGraphSchema_K2::PN_Self) // check its not the self pin
		{
			if(SrcPin->PinName == "Ref" &&SrcPin->bHidden == true)
			{  
			}
			else
			{
				UEdGraphPin* DestPin = CallFuncNode->FindPin(SrcPin->PinName);
				if (DestPin != NULL)
				{
					DestPin->PinType = SrcPin->PinType;
					DestPin->PinType.PinSubCategoryObject = SrcPin->PinType.PinSubCategoryObject;
					CompilerContext.MovePinLinksToIntermediate(*SrcPin, *DestPin); // Source node is assumed to be owner...
				}
			}
		}
	}

	// Finally, break any remaining links on the 'call func on member' node
	BreakAllNodeLinks();

  
}

FSlateIcon UK2Node_GetCellValue::GetIconAndTint(FLinearColor& OutColor) const
{
	OutColor = GetNodeTitleColor();
	static FSlateIcon Icon("EditorStyle", "Kismet.AllClasses.FunctionIcon");
	return Icon;
}

void UK2Node_GetCellValue::PostReconstructNode()
{
	Super::PostReconstructNode();
	for (auto& it : { FindPin(TEXT("Self")) ,FindPin(TEXT("ReturnValue")) })
	{
		PropagatePinType(it);
	}
}

void UK2Node_GetCellValue::EarlyValidation(class FCompilerResultsLog& MessageLog) const
{
	Super::EarlyValidation(MessageLog);
	 
}

 
void UK2Node_GetCellValue::NotifyPinConnectionListChanged(UEdGraphPin* Pin)
{
	Super::NotifyPinConnectionListChanged(Pin);
 
	if (Pin == FindPin(TEXT("Self")))
	{
		static TArray<FName> names = { "CellValue"  , "Cell" };

		auto cond = Pin->PinType.PinSubCategoryObject.IsValid() && names.Contains( Pin->PinType.PinSubCategoryObject->GetFName());

		auto _Ref = FindPin(TEXT("Ref"));
		if (_Ref->bHidden != cond)
		{
			_Ref->bHidden = cond;
			ReconstructNode();
		}
	}
	else if (Pin == FindPin(TEXT("ReturnValue"))) PropagatePinType(Pin);
} 
 
void UK2Node_GetCellValue::PropagatePinType(UEdGraphPin* Pin)
{
	if (Pin)
	{
		if (Pin->LinkedTo.Num() == 0)
		{
			Pin->ResetDefaultValue();
			Pin->PinType.PinCategory = UEdGraphSchema_K2::PC_Wildcard;
			return;
		}
		if (Pin->LinkedTo[0]->PinType.PinCategory == UEdGraphSchema_K2::PC_Wildcard)
		{
			 
			Pin->PinType.PinCategory = UEdGraphSchema_K2::PC_Struct;
			Pin->PinType.PinSubCategoryObject = FCellValue::StaticStruct();

			Pin->LinkedTo[0]->PinType.PinCategory = UEdGraphSchema_K2::PC_Struct;
			Pin->LinkedTo[0]->PinType.PinSubCategoryObject = FCellValue::StaticStruct();
		}
		else
		{
			Pin->PinType = Pin->LinkedTo[0]->PinType;
		}
	}
	 
}


bool UK2Node_GetCellValue::IsConnectionDisallowed(const UEdGraphPin* MyPin, const UEdGraphPin* OtherPin, FString& OutReason) const
{
	if (!ensure(OtherPin))
	{
		return true;
	}

	if (MyPin->PinName == "Self")
	{
		if (OtherPin->PinType.IsContainer())
		{
			OutReason = InputCheck.ToString();
			return true;
		}
		else if (OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Object)
		{
			static TArray<FName> names = { "Document","Sheet","Cell" };
			if (OtherPin->PinType.PinSubCategoryObject.IsValid() && names.Contains(OtherPin->PinType.PinSubCategoryObject->GetFName()))
			{
				return false;
			}
		}
		else if (OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Struct)
		{
			static TArray<FName> names = { "CellValue" };
			if (OtherPin->PinType.PinSubCategoryObject.IsValid() && names.Contains(OtherPin->PinType.PinSubCategoryObject->GetFName()))
			{
				return false;
			}
		}
		OutReason = InputCheck.ToString();
		return true;
	}
	else if (MyPin->PinName == "ReturnValue")
	{
		if (OtherPin->PinType.IsContainer())
		{
			OutReason = InputCheck.ToString();
			return true;
		}
		else if (OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_String ||
			OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Int ||
			OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Boolean ||
#if ENGINE_MAJOR_VERSION >=5
			OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Real
#else
			OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Float
#endif
			)
		{
			return false;
		}
		else if (OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Struct)
		{
			static TArray<FName> names = { "DateTime", "CellValue" };
			if (OtherPin->PinType.PinSubCategoryObject.IsValid() && names.Contains(OtherPin->PinType.PinSubCategoryObject->GetFName()))
			{
				return false;
			}
		}
		else if (OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Wildcard)
		{
			auto Pin = const_cast<UEdGraphPin*>(MyPin);
			Pin->PinType.PinCategory = UEdGraphSchema_K2::PC_Struct;
			Pin->PinType.PinSubCategoryObject = FCellValue::StaticStruct();
			return false; 
		}
		OutReason = InputCheck.ToString();
		return true;
	}
	return false;
}
FName UK2Node_GetCellValue::GetDestCall()const
{
	auto Target = FindPin(UEdGraphSchema_K2::PN_Self);
	auto Value = FindPin(TEXT("ReturnValue"));

	FString FunctionName = TEXT("GetCellValue_");
	if (Target->PinType.PinSubCategoryObject == nullptr)
	{
		auto type = this->GetBlueprint()->StaticClass();
		if (type->IsChildOf<UDocument>())
		{
			FunctionName.Append(TEXT("Doc"));
		}
		else if (type->IsChildOf<USheet>())
		{
			FunctionName.Append(TEXT("Sheet"));
		}
		else if (type->IsChildOf<UCell>())
		{
			FunctionName.Append(TEXT("Cell"));
		} 
	}
	else
	{
		if (Target->PinType.PinSubCategoryObject==UDocument::StaticClass())
		{
			FunctionName.Append(TEXT("Doc"));
		}
		else if (Target->PinType.PinSubCategoryObject==USheet::StaticClass())
		{
			FunctionName.Append(TEXT("Sheet"));
		}
		else if (Target->PinType.PinSubCategoryObject==UCell::StaticClass())
		{
			FunctionName.Append(TEXT("Cell"));
		}
		else  if (Target->PinType.PinSubCategoryObject == FCellValue::StaticStruct())
		{
			FunctionName.Append(TEXT("CellValue"));
		}
	}

	if (Value->PinType.PinCategory == UEdGraphSchema_K2::PC_Boolean)
	{
		FunctionName.Append(TEXT("Bool"));
	}
	else if (Value->PinType.PinCategory == UEdGraphSchema_K2::PC_Int)
	{
		FunctionName.Append(TEXT("Int"));
	}
	else if (
#if ENGINE_MAJOR_VERSION >=5
		Value->PinType.PinCategory == UEdGraphSchema_K2::PC_Real
#else
		Value->PinType.PinCategory == UEdGraphSchema_K2::PC_Float
#endif
		)
	{
		FunctionName.Append(TEXT("Float"));
	}
	else  if (Value->PinType.PinCategory == UEdGraphSchema_K2::PC_String)
	{
		FunctionName.Append(TEXT("String"));
	}
	else  if (Value->PinType.PinSubCategoryObject!=nullptr &&
		Value->PinType.PinSubCategoryObject->GetFName() == "FDateTime")
	{
		FunctionName.Append(TEXT("DateTime"));
	}
	else  if (Value->PinType.PinSubCategoryObject == FCellValue::StaticStruct())
	{
		FunctionName.Append(TEXT("CellValue"));
	}

	return *FunctionName;
}

#undef LOCTEXT_NAMESPACE