// Copyright Epic Games, Inc. All Rights Reserved.

using UnrealBuildTool;

public class FreeExcelEditor : ModuleRules
{
	public FreeExcelEditor(ReadOnlyTargetRules Target) : base(Target)
	{
		PCHUsage = ModuleRules.PCHUsageMode.UseExplicitOrSharedPCHs;

		bEnableExceptions = true;

		PublicIncludePaths.AddRange(
			new string[] {
				// ... add public include paths required here ...
				 
			}
			);
				
		
		PrivateIncludePaths.AddRange(
			new string[] {
				// ... add other private include paths required here ...
			}
			);
			
		
		PublicDependencyModuleNames.AddRange(
			new string[]
			{
				"Core", 
				"Projects",
				 "CoreUObject", 
				"Engine",
                 
				// ... add other public dependencies that you statically link with here ...
			}
			);
			
		
		PrivateDependencyModuleNames.AddRange(
			new string[]
			{
				"Slate",
                 "SlateCore",
                 "KismetCompiler",
                 "UnrealEd",
                 "BlueprintGraph",
                 "GraphEditor",
				 "FreeExcel",
			}
			);
		
		
		DynamicallyLoadedModuleNames.AddRange(
			new string[]
			{
				// ... add any modules that your module loads dynamically here ...
			}
			);
			
		 
	}
}
