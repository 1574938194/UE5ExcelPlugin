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


#include "K2Node_ForEachCellInRange.h"
#include "EdGraphSchema_K2.h"
#include "K2Node_CallFunction.h" 
#include "K2Node_ExecutionSequence.h"
#include "K2Node_IfThenElse.h" 
#include "K2Node_TemporaryVariable.h"
#include "KismetCompiler.h"
#include "Kismet/KismetNodeHelperLibrary.h"   
#include "BlueprintFieldNodeSpawner.h"
#include "EditorCategoryUtils.h"
#include "BlueprintActionDatabaseRegistrar.h"
#include "CellIterator.h"
#include "K2Node_AssignmentStatement.h"
#include "FreeExcelLibrary.h"
#include "Sheet.h"

#define LOCTEXT_NAMESPACE "K2Node"

struct FForExpandNodeHelper
{
	UEdGraphPin* StartLoopExecInPin;
	UEdGraphPin* TargetInPin;
	UEdGraphPin* RangeInPin;
	UEdGraphPin* InsideLoopExecOutPin;
	UEdGraphPin* LoopCompleteOutExecPin;

	UEdGraphPin* CellReferenceOutPin;
	UEdGraphPin* CellOutPin; 
	UEdGraphPin* IteratorOutPin;
	UEdGraphPin* IteratorEndOutPin;

	FForExpandNodeHelper()
		: StartLoopExecInPin(nullptr)
		, InsideLoopExecOutPin(nullptr)
		, LoopCompleteOutExecPin(nullptr)
		, CellReferenceOutPin(nullptr)
		, IteratorOutPin(nullptr)
		, IteratorEndOutPin(nullptr)
	{ }

	bool BuildLoop(UK2Node* Node, FKismetCompilerContext& CompilerContext, UEdGraph* SourceGraph)
	{
		const UEdGraphSchema_K2* Schema = CompilerContext.GetSchema();
		check(Node && SourceGraph && Schema);

		bool bResult = true;
		 
		// FCellIterator it
		UK2Node_TemporaryVariable* IteratorNode = CompilerContext.SpawnIntermediateNode<UK2Node_TemporaryVariable>(Node, SourceGraph);
		IteratorNode->VariableType.PinCategory = UEdGraphSchema_K2::PC_Struct;
		IteratorNode->VariableType.PinSubCategoryObject = FCellIterator::StaticStruct();
		IteratorNode->AllocateDefaultPins(); 
		IteratorOutPin = IteratorNode->GetVariablePin();
		check(IteratorOutPin);

		// FCellIterator end
		UK2Node_TemporaryVariable* IteratorEndNode = CompilerContext.SpawnIntermediateNode<UK2Node_TemporaryVariable>(Node, SourceGraph);
		IteratorEndNode->VariableType.PinCategory = UEdGraphSchema_K2::PC_Struct;
		IteratorEndNode->VariableType.PinSubCategoryObject = FCellIterator::StaticStruct();
		IteratorEndNode->AllocateDefaultPins();
		IteratorEndOutPin = IteratorEndNode->GetVariablePin();
		check(IteratorEndOutPin);

		// Make Range Internal
		UK2Node_CallFunction* RangeNode = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph);
		RangeNode->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, MakeCellRange_Internal)));
		RangeNode->AllocateDefaultPins();
		TargetInPin = RangeNode->FindPinChecked(TEXT("Sheet"));
		RangeInPin = RangeNode->FindPinChecked(TEXT("Range"));
		check(IteratorEndOutPin && RangeInPin); 

		// range.begin()
		UK2Node_CallFunction* RangeBegin = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph);
		RangeBegin->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, Range_Begin)));
		RangeBegin->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(RangeNode->GetReturnValuePin(), RangeBegin->FindPinChecked(TEXT("Target")));

		// range.end()
		UK2Node_CallFunction* RangeEnd = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph);
		RangeEnd->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, Range_End))); 
		RangeEnd->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(RangeNode->GetReturnValuePin(), RangeEnd->FindPinChecked(TEXT("Target")));

		// it = range.begin()
		UK2Node_AssignmentStatement* InitializeIterator = CompilerContext.SpawnIntermediateNode<UK2Node_AssignmentStatement>(Node, SourceGraph);
		InitializeIterator->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(IteratorOutPin, InitializeIterator->GetVariablePin());
		bResult &= Schema->TryCreateConnection(RangeBegin->FindPinChecked(UEdGraphSchema_K2::PN_ReturnValue), InitializeIterator->GetValuePin());
		StartLoopExecInPin = InitializeIterator->GetExecPin();
		check(StartLoopExecInPin);

		// end = range.end()
		UK2Node_AssignmentStatement* InitializeIteratorEnd = CompilerContext.SpawnIntermediateNode<UK2Node_AssignmentStatement>(Node, SourceGraph);
		InitializeIteratorEnd->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(IteratorEndOutPin, InitializeIteratorEnd->GetVariablePin());
		bResult &= Schema->TryCreateConnection(RangeEnd->FindPinChecked(UEdGraphSchema_K2::PN_ReturnValue), InitializeIteratorEnd->GetValuePin());
		bResult &= Schema->TryCreateConnection(InitializeIterator->GetThenPin(), InitializeIteratorEnd->GetExecPin());
		 

		// it->CellReference
		UK2Node_CallFunction* CellReferenceNode = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph);
		CellReferenceNode->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, CellIterator_CellReference)));
		CellReferenceNode->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(IteratorOutPin, CellReferenceNode->FindPinChecked(TEXT("Target"))); // should be Target?
		CellReferenceOutPin = CellReferenceNode->FindPinChecked(UEdGraphSchema_K2::PN_ReturnValue);
		check(CellReferenceOutPin);
		 
		// *it
		UK2Node_CallFunction* CellNode = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph);
		CellNode->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, CellIterator_Cell)));
		CellNode->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(IteratorOutPin, CellNode->FindPinChecked(TEXT("Target"))); // should be Target?
		CellOutPin = CellNode->FindPinChecked(UEdGraphSchema_K2::PN_ReturnValue);
		check(CellOutPin);
		 
		// for{...; if ;...}
		UK2Node_IfThenElse* Branch = CompilerContext.SpawnIntermediateNode<UK2Node_IfThenElse>(Node, SourceGraph);
		Branch->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(InitializeIteratorEnd->GetThenPin(), Branch->GetExecPin());
		LoopCompleteOutExecPin = Branch->GetElsePin();
		check(LoopCompleteOutExecPin);

		// for{...; <> ;...}
		UK2Node_CallFunction* Condition = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph); 
		Condition->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, CellIterator_NotEqual)));
		Condition->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(Condition->GetReturnValuePin(), Branch->GetConditionPin());
		bResult &= Schema->TryCreateConnection(Condition->FindPinChecked(TEXT("A")), IteratorOutPin);
		bResult &= Schema->TryCreateConnection(Condition->FindPinChecked(TEXT("B")), IteratorEndOutPin);
		 
		 
		// for{...} { <> }
		UK2Node_ExecutionSequence* Sequence = CompilerContext.SpawnIntermediateNode<UK2Node_ExecutionSequence>(Node, SourceGraph);
		Sequence->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(Branch->GetThenPin(), Sequence->GetExecPin());
		InsideLoopExecOutPin = Sequence->GetThenPinGivenIndex(0);
		check(InsideLoopExecOutPin);

		// ++it
		UK2Node_CallFunction* Forward = CompilerContext.SpawnIntermediateNode<UK2Node_CallFunction>(Node, SourceGraph); 
		Forward->SetFromFunction(UFreeExcelLibrary::StaticClass()->
			FindFunctionByName(GET_FUNCTION_NAME_CHECKED(UFreeExcelLibrary, CellIterator_Forward)));
		Forward->AllocateDefaultPins();
		bResult &= Schema->TryCreateConnection(Forward->FindPinChecked(TEXT("Target")), IteratorOutPin);
		bResult &= Schema->TryCreateConnection(Forward->GetExecPin(), Sequence->GetThenPinGivenIndex(1));
		bResult &= Schema->TryCreateConnection(Forward->GetThenPin(), Branch->GetExecPin());
  
		return bResult;
	}
};

UK2Node_ForEachCellInRange::UK2Node_ForEachCellInRange(const FObjectInitializer& ObjectInitializer)
	: Super(ObjectInitializer)
{
	NodeTooltip = LOCTEXT("ForEachCellInRange_NodeTooltip", "For each Cell in range do action");
}

const FName UK2Node_ForEachCellInRange::InsideLoopPinName(TEXT("LoopBody"));
const FName UK2Node_ForEachCellInRange::EnumOuputPinName(TEXT("EnumValue"));
const FName UK2Node_ForEachCellInRange::SkipHiddenPinName(TEXT("SkipHidden"));

void UK2Node_ForEachCellInRange::AllocateDefaultPins()
{
	UEdGraphSchema_K2 const* K2Schema = GetDefault<UEdGraphSchema_K2>();

	CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Exec, UEdGraphSchema_K2::PN_Execute);
	CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Exec, InsideLoopPinName);
	CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Object, USheet::StaticClass(), "Target");
	CreatePin(EGPD_Input, UEdGraphSchema_K2::PC_Struct, FCellRange::StaticStruct(), "Range");
	CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Struct, FCellReference::StaticStruct(), "Ref");
	CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Object, UCell::StaticClass(), "Cell");


	if (UEdGraphPin* CompletedPin = CreatePin(EGPD_Output, UEdGraphSchema_K2::PC_Exec, UEdGraphSchema_K2::PN_Then))
	{
		CompletedPin->PinFriendlyName = LOCTEXT("Completed", "Completed");
	}
}

void UK2Node_ForEachCellInRange::ValidateNodeDuringCompilation(FCompilerResultsLog& MessageLog) const
{
	Super::ValidateNodeDuringCompilation(MessageLog);
	 

}

FText UK2Node_ForEachCellInRange::GetNodeTitle(ENodeTitleType::Type TitleType) const
{
	return LOCTEXT("ForEachCellInRange_NodeTitle", "ForEach Cell In Range");
}

FSlateIcon UK2Node_ForEachCellInRange::GetIconAndTint(FLinearColor& OutColor) const
{
	static FSlateIcon Icon("EditorStyle", "GraphEditor.Macro.Loop_16x");
	return Icon;
}

void UK2Node_ForEachCellInRange::ExpandNode(class FKismetCompilerContext& CompilerContext, UEdGraph* SourceGraph)
{
	Super::ExpandNode(CompilerContext, SourceGraph);
	 
	FForExpandNodeHelper ForLoop;
	if (!ForLoop.BuildLoop(this, CompilerContext, SourceGraph))
	{
		CompilerContext.MessageLog.Error(*NSLOCTEXT("K2Node", "ForEachCellInRange_ForError", "For Expand error in @@").ToString(), this);
	}

	const UEdGraphSchema_K2* Schema = CompilerContext.GetSchema();

	bool bResult = true;

	CompilerContext.MovePinLinksToIntermediate(*GetExecPin(), *ForLoop.StartLoopExecInPin);
	CompilerContext.MovePinLinksToIntermediate(*FindPinChecked(UEdGraphSchema_K2::PN_Then), *ForLoop.LoopCompleteOutExecPin);
	CompilerContext.MovePinLinksToIntermediate(*FindPinChecked(InsideLoopPinName), *ForLoop.InsideLoopExecOutPin);
	CompilerContext.MovePinLinksToIntermediate(*FindPinChecked(TEXT("Ref")), *ForLoop.CellReferenceOutPin);
	CompilerContext.MovePinLinksToIntermediate(*FindPinChecked(TEXT("Cell")), *ForLoop.CellOutPin);
	CompilerContext.MovePinLinksToIntermediate(*FindPinChecked(TEXT("Target")), *ForLoop.TargetInPin);
	CompilerContext.MovePinLinksToIntermediate(*FindPinChecked(TEXT("Range")), *ForLoop.RangeInPin);

	if (!bResult)
	{
		CompilerContext.MessageLog.Error(*NSLOCTEXT("K2Node", "ForEachCellInRange_ExpandError", "Expand error in @@").ToString(), this);
	}

	BreakAllNodeLinks();
}

void UK2Node_ForEachCellInRange::GetMenuActions(FBlueprintActionDatabaseRegistrar& ActionRegistrar) const
{ 
	UClass* ActionKey = GetClass(); 
	if (ActionRegistrar.IsOpenForRegistration(ActionKey))
	{
		UBlueprintNodeSpawner* NodeSpawner = UBlueprintNodeSpawner::Create(GetClass());
		check(NodeSpawner != nullptr);

		ActionRegistrar.AddBlueprintAction(ActionKey, NodeSpawner);
	}
}

FText UK2Node_ForEachCellInRange::GetMenuCategory() const
{
	return LOCTEXT("ForEachCellInRange_GetMenuCategory", "FreeExcel");
}

void UK2Node_ForEachCellInRange::PostPlacedNewNode()
{
	Super::PostPlacedNewNode();
}

FText UK2Node_ForEachCellInRange::GetTooltipText() const
{
	return NodeTooltip;
}

void UK2Node_ForEachCellInRange::PinDefaultValueChanged(UEdGraphPin* ChangedPin)
{ 
}
void UK2Node_ForEachCellInRange::PostReconstructNode()
{
	Super::PostReconstructNode(); 
}

void UK2Node_ForEachCellInRange::EarlyValidation(class FCompilerResultsLog& MessageLog) const
{
	Super::EarlyValidation(MessageLog);
}


void UK2Node_ForEachCellInRange::NotifyPinConnectionListChanged(UEdGraphPin* Pin)
{
	Super::NotifyPinConnectionListChanged(Pin);
 
}
#undef LOCTEXT_NAMESPACE
