library Webinar;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

uses
  ShareMem,
  System.SysUtils,
  System.Classes,
  Vcl.Themes,
  Vcl.Styles,
  Windows,
  vcl.Dialogs,
  Vcl.Forms,
  Vcl.Controls,
  frmInterfaz_u in 'frmInterfaz_u.pas' {Interfaz},
  dataModule in 'dataModule.pas' {DataModule1: TDataModule},
  frmComplementario_u in 'frmComplementario_u.pas';

{$R *.res}

function FunctionCaptionW():PWideChar; export;stdcall;
  // returns a module caption which will be displayed in extension menu
 var
    test: string;
    mychar: PWideChar;
 begin
    test := 'My extension';
    myChar := @test[1];
    RESULT := myChar;
 end;




function FunctionName():PAnsiChar; export;stdcall;
  var
    text: AnsiString;
    mychar: PAnsiChar;

begin
    //returns name of the main function in the module which performs the required
    //task
//    RESULT := PAnsiChar('test');
    text := 'Main';
    myChar := @text[1];
    RESULT := myChar;
end;



function FunctionType():integer; export;stdcall;
begin
    //returns a number in range 0 – 3 indicating which set of additional parameters
    //should be passed to a function specified by FunctionName.
    //The function types are described later.
    RESULT := 0;
end;



procedure BasicParametersW(const DLLfilename :PWideChar;
    const ActiveLayerName :PWideChar;
    const FMDB_FileName :PWideChar;
    const DatabaseDir :PWideChar;
    const LicensedFieldMap :integer;
    const UseVirtualKeyboard :integer;
    const Int2, Int3, Int4, Int5, Int6, Int7, Int8, Int9 :integer;
    const CurrentUsername :PWideChar;
    const Reserved2, Reserved3, Reserved4, Reserved5 :PWideChar
    ); export;stdcall;

    begin
        {
            Variables declaradas en dmConexionBDFieldMap_u
            Son para configurar la conexion a la base de datos
        }
        DirectorioProyecto := WideCharToString(DatabaseDir);
        SetCurrentDir(DirectorioProyecto);
        NombreArchivoBD := 'FieldMapData_' +
                            ExtractFileName(DirectorioProyecto) +
                            '.';

        RutaAbsolutaBDFieldMap := DirectorioProyecto +
                                  '\' +
                                  NombreArchivoBD;

//        NombreUsuarioFieldMap := WideCharToString(CurrentUsername);

    end;




procedure Main(Parent: HWND); export;

begin
    Application.Initialize;
    Application.MainFormOnTaskbar := True;
    Application.CreateForm(TInterfaz, Interfaz);
  Application.CreateForm(Tmiscelaneos, miscelaneos);
  Application.CreateForm(TDataModule1, DataModule1);
  Interfaz.ShowModal;
end;



exports FunctionCaptionW, FunctionName, FunctionType, BasicParametersW, Main;

begin
end.
