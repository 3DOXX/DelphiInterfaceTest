# DelphiInterfaceTest

Delphi: Konsolenanwendung, die sich aus der DLL das Interface holt und eine Testfunktion ausführt

C#: Vollständige Solution, die die DLL zusammenbaut.

Der Problematische Code:

	SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open("G:\\Vorlagendatei.xlsm", true); // Pfad ändern!
	WorkbookStylesPart stylesPart = spreadSheet.WorkbookPart.WorkbookStylesPart;
	Stylesheet stylesheet = stylesPart.Stylesheet;
	spreadSheet.Close();

Führt man den C#-Code in einer .NET-Umgebung, kommt der Fehler nicht. Der Fehler tritt genau dann auf, wenn versucht wird, das Stylesheet auszulesen. Zurückgegeben wird eine StackOverflowException, man erfährt aber nie den zugehörigen Stack.

Ihr braucht in C# die Packages

OpenXML
DllExport