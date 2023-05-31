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

#include "K2Node_ToTemplateArray.h"
#include "EdGraphSchema_K2.h"
#include "BlueprintActionDatabaseRegistrar.h"
#include "BlueprintNodeSpawner.h"
#include "EditorCategoryUtils.h"
#include "K2Node_CallFunction.h"
#include "KismetCompiler.h"
#include "FreeExcelLibrary.h"
#include "K2Node_Self.h"
#include "Document.h"
 
#define LOCTEXT_NAMESPACE "K2Node_ToTemplateArray"

 

UK2Node_ToTemplateArray::UK2Node_ToTemplateArray(const FObjectInitializer& ObjectInitializer)
	: Super(ObjectInitializer)
{
	NodeTooltip = LOCTEXT("ToTemplateArray_NodeTooltip", "Deserialize Sheet to a Struct list");
	InputCheck = LOCTEXT("ToTemplateArray_InputCheck", "Cannot input this type!");
}

void UK2Node_ToTemplateArray::SetPinToolTip(UEdGraphPin& MutatablePin, const FText& PinDescription) const
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

void UK2Node_ToTemplateArray::AllocateDefaultPins()
{
	  

	const UEdGraphSchema_K2* K2Schema = GetDefault<UEdGraphSchema_K2>();

	// Add execution pins
	/*CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Exec, UEdGraphSchema_K2::PN_Execute);
	CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Exec, UEdGraphSchema_K2::PN_Then);
	 */
	auto SelfPin = CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Wildcard, UEdGraphSchema_K2::PN_Self);
	SetPinToolTip(*SelfPin, LOCTEXT("ToTemplateArray_SelfPinDesc", "[Document,Sheet] In"));
	 
	auto ReturnValuePin = CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Wildcard, UEdGraphSchema_K2::PN_ReturnValue);
	ReturnValuePin->PinType.ContainerType = EPinContainerType::Array;
	SetPinToolTip(*ReturnValuePin, LOCTEXT("ToTemplateArray_ReturnValuePinDesc", "Array<T> Out"));

	Super::AllocateDefaultPins();
}

  
void UK2Node_ToTemplateArray::GetMenuActions(FBlueprintActionDatabaseRegistrar& ActionRegistrar) const
{ 
	UClass* ActionKey = GetClass(); 
	if (ActionRegistrar.IsOpenForRegistration(ActionKey))
	{
		UBlueprintNodeSpawner* NodeSpawner = UBlueprintNodeSpawner::Create(GetClass());
		check(NodeSpawner != nullptr);

		ActionRegistrar.AddBlueprintAction(ActionKey, NodeSpawner);
	}
}

FText UK2Node_ToTemplateArray::GetMenuCategory() const
{
	return LOCTEXT("ToTemplateArray_GetMenuCategory", "FreeExcel");
}


void UK2Node_ToTemplateArray::PinDefaultValueChanged(UEdGraphPin* ChangedPin)
{
	TArray<FName> ls = { UEdGraphSchema_K2::PN_Self, UEdGraphSchema_K2::PN_ReturnValue };
	if (ls.Contains(ChangedPin->GetFName()))
	{
		PropagatePinType(ChangedPin);
	}
	 
}
 

FText UK2Node_ToTemplateArray::GetTooltipText() const
{
	return NodeTooltip;
}


FText UK2Node_ToTemplateArray::GetNodeTitle(ENodeTitleType::Type TitleType) const
{
	return LOCTEXT("ToTemplateArray_NodeTitle", "To Template Array");
}

void UK2Node_ToTemplateArray::ExpandNode(class FKismetCompilerContext& CompilerContext, UEdGraph* SourceGraph)
{
	Super::ExpandNode(CompilerContext, SourceGraph);

 
	const UEdGraphSchema_K2* Schema = CompilerContext.GetSchema();

	UK2Node_CallFunction* CallFuncNode = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(this, SourceGraph);
	//CallFuncNode->SetFromFunction(Function);
	const FName DstFunctionName = GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, ToTemplateArray);
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
		if (SrcPin != NULL && SrcPin->PinName != UEdGraphSchema_K2::PN_Self) // check its not the self pin
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

	// Finally, break any remaining links on the 'call func on member' node
	BreakAllNodeLinks();

  
}

FSlateIcon UK2Node_ToTemplateArray::GetIconAndTint(FLinearColor& OutColor) const
{
	OutColor = GetNodeTitleColor();
	static FSlateIcon Icon("EditorStyle", "Kismet.AllClasses.FunctionIcon");
	return Icon;
}

void UK2Node_ToTemplateArray::PostReconstructNode()
{
	Super::PostReconstructNode();
	for (auto& it : { FindPin(UEdGraphSchema_K2::PN_Self) ,FindPin(UEdGraphSchema_K2::PN_ReturnValue) })
	{
		PropagatePinType(it);
	}
}

void UK2Node_ToTemplateArray::EarlyValidation(class FCompilerResultsLog& MessageLog) const
{
	Super::EarlyValidation(MessageLog);
 
}

 
void UK2Node_ToTemplateArray::NotifyPinConnectionListChanged(UEdGraphPin* Pin)
{
	Super::NotifyPinConnectionListChanged(Pin);
  
	if (Pin->PinName == UEdGraphSchema_K2::PN_Self)
	{
		PropagatePinType(Pin);
	}
	else if (Pin->PinName == UEdGraphSchema_K2::PN_ReturnValue)
	{
		PropagatePinType(Pin); 
	}
} 
 
void UK2Node_ToTemplateArray::PropagatePinType(UEdGraphPin* Pin)
{ 
	if (Pin) 
	{
		if (Pin->LinkedTo.Num() == 0)
		{
			Pin->ResetDefaultValue();
			Pin->PinType.PinCategory = UEdGraphSchema_K2::PC_Wildcard;
			return;
		}
		Pin->PinType = Pin->LinkedTo[0]->PinType;

	} 
}

 bool UK2Node_ToTemplateArray::IsConnectionDisallowed(const UEdGraphPin* MyPin, const UEdGraphPin* OtherPin, FString& OutReason) const
{
	if (!ensure(OtherPin))
	{
		return true;
	}
	 
	if (MyPin->PinName== UEdGraphSchema_K2::PN_Self)
	{
		if (OtherPin->PinType.IsContainer())
		{
			OutReason = InputCheck.ToString();
			return true;
		}
		else if (OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Object)
		{ 
			static TArray<FName> names = {"Document","Sheet"};
			if (OtherPin->PinType.PinSubCategoryObject.IsValid() && names.Contains(OtherPin->PinType.PinSubCategoryObject->GetFName()))
			{ 
				return false;
			}
		} 
		OutReason = InputCheck.ToString();
		return true;
	}
	else if (MyPin->PinName == UEdGraphSchema_K2::PN_ReturnValue)
	{
		if (OtherPin->PinType.ContainerType == EPinContainerType::Array &&
			(
				OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Struct||
				OtherPin->PinType.PinCategory == UEdGraphSchema_K2::PC_Object
				))
		{ 
			return false;
		} 
		OutReason = InputCheck.ToString();
		return true;
	}
	 
	return false;
}
#undef LOCTEXT_NAMESPACE