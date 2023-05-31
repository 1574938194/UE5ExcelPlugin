// Copyright Epic Games, Inc. All Rights Reserved.

using UnrealBuildTool;

public class Ue5LearnProj : ModuleRules
{
	public Ue5LearnProj(ReadOnlyTargetRules Target) : base(Target)
	{
		PCHUsage = PCHUsageMode.UseExplicitOrSharedPCHs;

		PublicDependencyModuleNames.AddRange(new string[] { "Core", "CoreUObject", "Engine", "InputCore", "HeadMountedDisplay" });
	}
}
