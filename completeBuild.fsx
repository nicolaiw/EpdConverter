// include Fake lib
#r @"packages\FAKE\tools\FakeLib.dll"

open Fake
open Fake.AssemblyInfoFile

RestorePackages()

// Directories
let buildDir  = @".\build\"
let testDir   = @".\test\"
let deployDir = @".\deploy\"
let packagesDir = @".\packages"

// tools
let fxCopRoot = @".\Tools\FxCop\FxCopCmd.exe"

// version info
let version = "0.2"  // or retrieve from CI server

// Targets
Target "Clean" (fun _ ->
    CleanDirs [buildDir; testDir; deployDir]
)

Target "SetVersions" (fun _ ->
    CreateCSharpAssemblyInfo "./src/app/EpdConverter.Core/Properties/AssemblyInfo.cs"
        [Attribute.Title "EpdToExcel.Core"
         Attribute.Description "Epd to Excel converter"
         Attribute.Guid "4daa8a08-9b1e-42e6-b2a3-8bb8a7f71199"
         Attribute.Product "EpdToExcel.Core"
         Attribute.Version version
         Attribute.FileVersion version]

    CreateCSharpAssemblyInfo "./src/app/EpdConverter.Console.Test/Properties/AssemblyInfo.cs"
        [Attribute.Title "EpdToExcel.Console.Test"
         Attribute.Description "EpdToExcel.Console.Test"
         Attribute.Guid "09c1bce8-d5dd-4b42-b548-16ad8a67fba8"
         Attribute.Product "EpdToExcel.Console.Test"
         Attribute.Version version
         Attribute.FileVersion version]
)

Target "CompileApp" (fun _ ->
    !! @"src\**\*.csproj"
      |> MSBuildRelease buildDir "Build"
      |> Log "AppBuild-Output: "
)

Target "CompileTest" (fun _ ->
    !! @"src\test\**\*.csproj"
      |> MSBuildDebug testDir "Build"
      |> Log "TestBuild-Output: "
)

Target "NUnitTest" (fun _ ->
    !! (testDir + @"\NUnit.Test.*.dll")
      |> NUnit (fun p ->
                 {p with
                   DisableShadowCopy = true;
                   OutputFile = testDir + @"TestResults.xml"})
)

Target "FxCop" (fun _ ->
    !! (buildDir + @"\**\*.dll")
      ++ (buildDir + @"\**\*.exe")
        |> FxCop (fun p ->
            {p with
                ReportFileName = testDir + "FXCopResults.xml";
                ToolPath = fxCopRoot})
)

Target "Zip" (fun _ ->
    !! (buildDir + "\**\*.*")
        -- "*.zip"
        |> Zip buildDir (deployDir + "Calculator." + version + ".zip")
)

// Dependencies
"Clean"
  //==> "SetVersions"
  ==> "CompileApp"
  //==> "CompileTest"
  //==> "FxCop"
  //==> "NUnitTest"
  //==> "Zip"

// start build
RunTargetOrDefault "CompileApp"