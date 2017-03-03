C:\Windows\Microsoft.NET\Framework64\v4.0.30319\msbuild ../FERB.sln /p:Configuration=Release
nuget pack ../FERB/FERB.csproj -Prop Configuration=Release
nuget push *.nupkg -Source https://www.nuget.org/api/v2/package
del *.nupkg