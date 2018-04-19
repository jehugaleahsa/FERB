&dotnet pack "..\FERB\FERB.csproj" --configuration Release --output $PWD

.\NuGet.exe push FERB.*.nupkg -Source https://www.nuget.org/api/v2/package

Remove-Item FERB.*.nupkg