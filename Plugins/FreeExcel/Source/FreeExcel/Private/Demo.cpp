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
 
#include "Demo.h"
#include "Kismet/KismetSystemLibrary.h" 
#include "Sheet.h"
#include <string>
#include <iostream>
#include "Document.h"
#include <random>
#include <algorithm>
#include <numeric>
#include "FreeExcelLibrary.h"
#include "CellValue.h" 


using namespace std; 

// Sets default values
ADemo::ADemo()
{
 	// Set this actor to call Tick() every frame.  You can turn this off to improve performance if you don't need it.
	PrimaryActorTick.bCanEverTick = true;

}

// Called when the game starts or when spawned
void ADemo::BeginPlay()
{
	Super::BeginPlay();
	
}

// Called every frame
void ADemo::Tick(float DeltaTime)
{
	Super::Tick(DeltaTime);

}
 
void ADemo::RunDemo()
{
    // Create excel document and get first sheet 
    auto doc = NewObject<UDocument>();
    doc->Create(TEXT("D:\\proj\\ueTest\\Demo01.xlsx"));
    auto wks = doc->GetOrCreateSheetWithName(TEXT("Sheet1"));

    // Basic Usage
    wks->Cell({ 1,1 })->SetFloat(3.1415926);
    wks->Cell(UFreeExcelLibrary::MakeCellReferenceWithString("B1"))->SetInt(42);
    wks->Cell(UFreeExcelLibrary::MakeCellReference(1, 3))->SetString(TEXT("Hello FreeExcel"));
    wks->Cell("D1")->SetBool(true);

    auto A1 = wks->Cell("A1")->ToFloat();
    auto B1 = wks->Cell("B1")->ToInt();
    auto C1 = wks->Cell("C1")->ToString();
    auto D1 = wks->Cell("D1")->ToBool();

    UKismetSystemLibrary::PrintString(this,
        FString::SanitizeFloat(A1) + "," +
        FString::FromInt(B1) + "," +
        C1 + "," +
        FString(D1 ? TEXT("true") : TEXT("flase"))
    );

    
    wks->Cell("A2")->SetCellValue(wks->Cell("C1")->Value());
    UKismetSystemLibrary::PrintString(this, "," + wks->Cell("A2")->ToString());

    // DateTime Value
    wks->Cell(TEXT("B2"))->SetDateTime(FDateTime::Now());
    UKismetSystemLibrary::PrintString(this, wks->Cell("B2")->ToDateTime().ToString(TEXT("%Y-%m-%d-%H-%M-%S")));
 
    // Formula Usage
    wks->Cell("C2")->SetFormula( "SQRT(B1)");

    // Sheet handling 
    doc->GetOrCreateSheetWithName("Sheet2");
    doc->CloneSheet("Sheet1", "Sheet3");
     
    doc->RemoveSheet("Sheet2"); 
    //Unicode
    wks->Cell(TEXT("D2"))->SetString( TEXT("こんにちは世界"));

    // Row Handling
    auto values = wks->GetRowData(1);
    for (auto it : values)
    {
        switch (it.type())  
        {
        case EValueType::Boolean: UKismetSystemLibrary::PrintString(this, FString((bool)it?TEXT("true"):TEXT("false"))); break;
        case EValueType::Integer: UKismetSystemLibrary::PrintString(this, FString::FromInt((int32)it)); break;
        case EValueType::Float: UKismetSystemLibrary::PrintString(this, FString::SanitizeFloat((float)it)); break;
        case EValueType::String: UKismetSystemLibrary::PrintString(this,it.operator FString()); break;
        default:
            break;
        }
    }

    // Ranges and Iterators 
    std::random_device                 rand_dev;
    std::mt19937                       generator(rand_dev());
    std::uniform_int_distribution<int> distr(0, 99);

//    auto range = UFreeExcelLibrary::MakeCellRangeWithString("A3:D4");
 
    //for (auto it = range.begin() ; it!= range.end() ; ++it)
    //{
    //    auto ref = it.Current;

    //    int i = 3;
    //}
    
    // save and close doc
    doc->Save();
    doc->Close();
}


void ADemo::RunDemo2()
{ 
    auto doc = NewObject<UDocument>();
    doc->Open(TEXT("D:\\proj\\ueTest\\Demo02.xlsx"));
    auto wks = doc->GetOrCreateSheetWithName(TEXT("Sheet1"));
      
    

    doc->Close();
}
    