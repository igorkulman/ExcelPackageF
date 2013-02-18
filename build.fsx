// include Fake lib
#r @"tools\FAKE\tools\FakeLib.dll"
open Fake
 
// Properties
let buildDir = @".\build\"
 
// Targets
Target "Clean" (fun _ ->
    CleanDir buildDir
)

Target "BuildApp" (fun _ ->
    !! @"ExcelPackageF\*.fsproj"
      |> MSBuildRelease buildDir "Build"
      |> Log "AppBuild-Output: "
)
 
Target "Default" (fun _ ->
    trace "Hello World from FAKE"
)
 
// Dependencies
"Clean"
  ==> "BuildApp"
  ==> "Default"
 
// start build
Run "Default"