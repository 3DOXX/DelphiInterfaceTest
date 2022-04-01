program Project1;

{$APPTYPE CONSOLE}
{$R *.res}

uses
  System.SysUtils,
  ActiveX;

type
  IExcelLib2 = interface
  ['{07056E7A-183B-49B5-8391-F32DBC68ECC4}']
    procedure test ( var Cardinal ) ; stdcall ;
  end;

procedure OpenExcelFile ( var IExcelLib2 ) ; stdcall ; external 'openxml2.dll' ;

var
  intf : IExcelLib2;
  res : Cardinal ;

begin
  try
    intf := nil;
    OpenExcelFile(intf);

    if Assigned(intf) then

    begin
      intf.test(res);
      // res muss 23 ausgeben, sonst kam es in der test-Methode zu einer Exception
      WriteLn(res);
      WriteLn('ok');
      intf := nil;
      ReadLn;
    end;
  except
    on E: Exception do
    begin
      Writeln(E.ClassName, ': ', E.Message);
    ReadLn;
    end;
  end;
end.
