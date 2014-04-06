// include Fake lib
#r @"tools\FAKE\tools\FakeLib.dll"
open Fake
 
RestorePackages()

// Properties
let buildDir = @".\build\"
let testDir  = @".\test\"
let packagesDir = @".\packages"
let packagingRoot = "./packaging/"
let packagesVersion = "1.0.5"

// tools
let nunitVersion = GetPackageVersion packagesDir "NUnit.Runners"
let nunitPath = sprintf @"./packages/NUnit.Runners.%s/tools/" nunitVersion

 
// Targets
Target "Clean" (fun _ ->
    CleanDirs [buildDir; testDir; packagingRoot]
)

Target "BuildApp" (fun _ ->
    !! @"ExcelPackageF\*.fsproj"
      |> MSBuildRelease buildDir "Build"
      |> Log "AppBuild-Output: "
)

Target "BuildTest" (fun _ ->
   !! @"ExcelPackageF.Tests\*.fsproj"
      |> MSBuildDebug testDir "Build"
      |> Log "TestBuild-Output: "
)

Target "Test" (fun _ ->
    !! (testDir + @"\ExcelPackageF.Tests.dll") 
      |> NUnit (fun p ->
          {p with
             ToolPath = nunitPath;
             DisableShadowCopy = true;
             OutputFile = testDir + @"TestResults.xml" })
)

Target "CreateNugetPackage" (fun _ ->    
    NuGet (fun p -> 
        {p with                  
            Project = "ExcelPackageF"          
            OutputPath = packagingRoot
            WorkingDir = buildDir
            Version = packagesVersion
            Dependencies =
                ["EPPlus", GetPackageVersion "./packages/" "EPPlus"]
            Publish = false
            }) "ExcelPackageF.nuspec"
)

Target "Default" (fun _ ->
    trace "Build completed"
)
 
// Dependencies
"Clean"
  ==> "BuildApp"
  ==> "BuildTest"
  ==> "Test"
  ==> "CreateNugetPackage"
  ==> "Default"
 
// start build
Run "Default"