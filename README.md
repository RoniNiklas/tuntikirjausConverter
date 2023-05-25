Development:    
Export Tuntiraportti.xlsx from eKieku and put it in the Project folder where program.cs is   
Press start in Visual Studio 2022    

Publishing:    
dotnet publish -c Release    
Requires C++ workload for desktop to be installed with Visual Studio 2022 installer. Because AOT publishing does not work without it.     
