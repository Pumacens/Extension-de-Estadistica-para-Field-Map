unit dataModule;

interface

uses
  System.SysUtils, System.Classes, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.VCLUI.Wait,
  FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt,
  FireDAC.Phys.FBDef, FireDAC.Phys.MSAccDef, FireDAC.Phys.ODBCBase,
  FireDAC.Phys.MSAcc, FireDAC.Phys.IBBase, FireDAC.Phys.FB, Data.DB,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, Vcl.Dialogs,
  Generics.Collections, AdvSmoothListBox, Contnrs, StrUtils;

type
  TDataModule1 = class(TDataModule)
    ConexionBD: TFDConnection;
    SQLQuery: TFDQuery;
    FDPhysFBDriverLink1: TFDPhysFBDriverLink;
    FDPhysMSAccessDriverLink1: TFDPhysMSAccessDriverLink;
    procedure DataModuleCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
    DataModule1: TDataModule1;
    TipoBDEncontrada: TSearchRec;
    RutaAbsolutaBDFieldMap,
    NombreCapaParcela,
    NombreArchivoBD: String;
    num_parcelas: Integer;
    numIndvPorCapa: TStringList;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

uses
    frmInterfaz_u,
    frmComplementario_u;
{$R *.dfm}

procedure TDataModule1.DataModuleCreate(Sender: TObject);
var
    ListaCapasFieldMap: TStrings;
    i, j: Integer;

begin
    ListaCapasFieldMap := TStringList.Create;
    numIndvPorCapa := TStringList.Create;

    if FindFirst(NombreArchivoBD + '*', faAnyFile, TipoBDEncontrada) = 0 then begin
        // Reestableciendo la conexion a la base de datos
        ConexionBD.Connected := False;
        ConexionBD.Params.Clear;
        // Si en la carpeta del proyecto encontró una base de datos Access
        if ExtractFileExt(TipoBDEncontrada.Name) = '.accdb' then begin
            RutaAbsolutaBDFieldMap := RutaAbsolutaBDFieldMap + 'accdb';

            try
                // Configurando base de datos para Access
                ConexionBD.Params.Add('DriverID=MSAcc');
                ConexionBD.Params.Add('Database=' + RutaAbsolutaBDFieldMap);
                FindClose(TipoBDEncontrada);
                // Conectarse a la base de datos
                ConexionBD.Connected := True;
//                SQLQuery.Active := True;
//                ShowMessage('Conectado a la base de datos = ' + BoolToStr(ConexionBD.Connected, True));
//                ShowMessage('Está activado el SQLQuery = ' + BoolToStr(SQLQuery.Active, True));


            // Mostrar cual es el catrehpta error si algo falla con las bases de datos
            except on E: Exception do
                begin
                    ShowMessage('Para Access - Exception class name = ' + E.ClassName);
                    ShowMessage('Para Access - Exception message = ' + E.Message);
                end;
            end;

        // Si en la carpeta del proyecto encontró una base de datos FireBird
        end else if ExtractFileExt(TipoBDEncontrada.Name) = '.FDB' then begin
            RutaAbsolutaBDFieldMap := RutaAbsolutaBDFieldMap + 'FDB';

            try
                // Configurando base de datos para FireBird
                ConexionBD.Params.Add('DriverID=FB');
                ConexionBD.Params.Add('Database=' + RutaAbsolutaBDFieldMap);
                ConexionBD.Params.Add('User_Name=' + 'sysdba');
                ConexionBD.Params.Add('Password=' + 'masterkey');
                ConexionBD.Params.Add('OSAuthent=No');
                FindClose(TipoBDEncontrada);
                // Conectarse a la base de datos
                ConexionBD.Connected := True;

//                ShowMessage('Conectado a la base de datos = ' + BoolToStr(ConexionBD.Connected, True));
//                ShowMessage('Está activado el SQLQuery = ' + BoolToStr(SQLQuery.Active, True));

            // Mostrar cual es el catrehpta error si algo falla con las bases de datos
            except on E : Exception do
                begin
                    ShowMessage('Para FDB - Exception class name = ' + E.ClassName);
                    ShowMessage('Para FDB - Exception message = ' + E.Message);
                end;
            end;
        end;
    end;

    // Guarda el nombre de todas las tablas de la BD en ListaCapasFieldMap
    ConexionBD.GetTableNames('', '', '', ListaCapasFieldMap);

    for i := 0 to ListaCapasFieldMap.Count-1 do begin
        {
            Abre capa por capa buscando cuales tienen especies por dentro
            ya que con esas es que se realizará el análisis estadístico
        }
        SQLQuery.Open('SELECT * FROM ' + ListaCapasFieldMap[i]);
        SQLQuery.FetchAll;

        if (AnsiStartsStr('ID', SQLQuery.Fields[0].FieldName)) and
           (Length(SQLQuery.Fields[0].FieldName) > 2)                 then begin
           // Almaceno el nombre de la capa de parcelas
           NombreCapaParcela := AnsiReplaceStr(SQLQuery.Fields[0].FieldName, 'ID', '');
        end;


        try
            // Atributos que deben tener las capas para ser analizables
            // en la extension
            SQLQuery.FieldByName('Species');
            SQLQuery.FieldByName('X_m');

            {
              Comprobar que la capa arbol a analizar tiene almenos una fila con informacion
              Sino, no se agrega la capa.
            }

            if (SQLQuery.Table.Rows.Count = 0) then begin
                continue

            end;

            numIndvPorCapa.Add(IntToStr(SQLQuery.Table.Rows.Count));

            with Interfaz.capasArbolIVI.Items.Add do begin
                Caption := ListaCapasFieldMap[i];
//                // Pone el icono de arbol de la lista de miscelaneos
                GraphicLeft := miscelaneos._ItemsIconos.Items[0].GraphicLeft;
                GraphicLeftType := miscelaneos._ItemsIconos.Items[0].GraphicLeftType;
                GraphicLeftWidth := miscelaneos._ItemsIconos.Items[0].GraphicLeftWidth;
                GraphicLeftMargin := miscelaneos._ItemsIconos.Items[0].GraphicLeftMargin;
                GraphicLeftHeight := miscelaneos._ItemsIconos.Items[0].GraphicLeftHeight;
                CaptionFont.Size := miscelaneos._ItemsIconos.Items[0].CaptionFont.Size ;
                CaptionSelectedFont.Size := miscelaneos._ItemsIconos.Items[0].CaptionSelectedFont.Size ;
                CaptionMargin := miscelaneos._ItemsIconos.Items[0].CaptionMargin;
            end;

            with Interfaz.capasArbolDiversidad.Items.Add do begin
                Caption := ListaCapasFieldMap[i];
//                // Pone el icono de arbol de la lista de miscelaneos
                GraphicLeft := miscelaneos._ItemsIconos.Items[0].GraphicLeft;
                GraphicLeftType := miscelaneos._ItemsIconos.Items[0].GraphicLeftType;
                GraphicLeftWidth := miscelaneos._ItemsIconos.Items[0].GraphicLeftWidth;
                GraphicLeftMargin := miscelaneos._ItemsIconos.Items[0].GraphicLeftMargin;
                GraphicLeftHeight := miscelaneos._ItemsIconos.Items[0].GraphicLeftHeight;
                CaptionFont.Size := miscelaneos._ItemsIconos.Items[0].CaptionFont.Size ;
                CaptionSelectedFont.Size := miscelaneos._ItemsIconos.Items[0].CaptionSelectedFont.Size ;
                CaptionMargin := miscelaneos._ItemsIconos.Items[0].CaptionMargin;
            end;

            with Interfaz.capasArbolRiqueza.Items.Add do begin
                Caption := ListaCapasFieldMap[i];
//                // Pone el icono de arbol de la lista de miscelaneos
                GraphicLeft := miscelaneos._ItemsIconos.Items[0].GraphicLeft;
                GraphicLeftType := miscelaneos._ItemsIconos.Items[0].GraphicLeftType;
                GraphicLeftWidth := miscelaneos._ItemsIconos.Items[0].GraphicLeftWidth;
                GraphicLeftMargin := miscelaneos._ItemsIconos.Items[0].GraphicLeftMargin;
                GraphicLeftHeight := miscelaneos._ItemsIconos.Items[0].GraphicLeftHeight;
                CaptionFont.Size := miscelaneos._ItemsIconos.Items[0].CaptionFont.Size ;
                CaptionSelectedFont.Size := miscelaneos._ItemsIconos.Items[0].CaptionSelectedFont.Size ;
                CaptionMargin := miscelaneos._ItemsIconos.Items[0].CaptionMargin;
            end;

        except on EDatabaseError do
            continue;
        end;

    end;

//    ShowMessage(RutaAbsolutaBDFieldMap);
end;

end.
