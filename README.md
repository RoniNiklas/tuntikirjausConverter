Development:    
Export Tuntiraportti.xlsx from eKieku and put it in the Project folder where program.cs is   
Press start in Visual Studio 2022    
Go to tuntikirjaukset/bin/debug/net7.0/win-x64 and check that the outputted .csv file matches the desired outcome   

Publishing:    
dotnet publish -c Release    
Requires C++ workload for desktop to be installed with Visual Studio 2022 installer. Because AOT publishing does not work without it.     

Use:   
Run the published .exe in the same folder as a Tuntiraportti.xlsx file.
The program will output a .csv file with the tuntiraportti-data grouped by project for easier tracking of work hours and workdays spent on each project.