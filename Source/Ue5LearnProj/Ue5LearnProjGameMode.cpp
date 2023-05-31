// Copyright Epic Games, Inc. All Rights Reserved.

#include "Ue5LearnProjGameMode.h"
#include "Ue5LearnProjCharacter.h"
#include "UObject/ConstructorHelpers.h"

AUe5LearnProjGameMode::AUe5LearnProjGameMode()
{
	// set default pawn class to our Blueprinted character
	static ConstructorHelpers::FClassFinder<APawn> PlayerPawnBPClass(TEXT("/Game/ThirdPerson/Blueprints/BP_ThirdPersonCharacter"));
	if (PlayerPawnBPClass.Class != NULL)
	{
		DefaultPawnClass = PlayerPawnBPClass.Class;
	}
}
