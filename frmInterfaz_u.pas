unit frmInterfaz_u;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, AdvMetroCategoryList,
  Vcl.CategoryButtons, AdvSmoothPanel, AdvSmoothTabPager, Vcl.Imaging.pngimage,
  Vcl.ExtCtrls, AdvSmoothLabel, tmsAdvGridExcel, AdvUtil, Vcl.Grids, AdvObj,
  BaseGrid, AdvGrid, AdvSprd, advmultibuttonedit, VclTee.TeeGDIPlus,
  VCLTee.TeEngine, VCLTee.TeeProcs, VCLTee.Chart, AdvSmoothListBox,
  AdvSmoothComboBox, VCLTee.Series, AdvSmoothScrollBar, dataModule,
  AdvSmoothToggleButton, AdvScrollBox, StrUtils, Generics.Collections,
  AdvGlowButton, Vcl.StdCtrls;

type
  TInterfaz = class(TForm)
    Menu: TAdvSmoothTabPager;
    PageIVI: TAdvSmoothTabPage;
    PageDiversidad: TAdvSmoothTabPage;
    PageRiqueza: TAdvSmoothTabPage;
    AdvSmoothPanel1: TAdvSmoothPanel;
    Image1: TImage;
    AdvSmoothLabel1: TAdvSmoothLabel;
    exportadorIVI: TAdvGridExcelIO;
    AdvGridUndoRedo1: TAdvGridUndoRedo;
    IVIExcelData: TAdvSpreadGrid;
    capasArbolIVI: TAdvSmoothComboBox;
    GraficoIVI: TChart;
    panelParaScroll: TAdvSmoothPanel;
    panelExcel: TAdvSmoothPanel;
    IVIExcelHeader: TAdvSpreadGrid;
    AdvSmoothLabel2: TAdvSmoothLabel;
    AdvSmoothLabel3: TAdvSmoothLabel;
    AdvSmoothLabel4: TAdvSmoothLabel;
    botonIVIEspecies: TAdvSmoothPanel;
    botonIVIFamilias: TAdvSmoothPanel;
    PanelIVIInterno: TAdvScrollBox;
    SerieFrecuencia: TBarSeries;
    SerieAbundancia: TBarSeries;
    PanelDiversidadInterno: TAdvScrollBox;
    ShannonSimpsonDatosBase: TAdvSpreadGrid;
    capasArbolDiversidad: TAdvSmoothComboBox;
    AdvSmoothLabel5: TAdvSmoothLabel;
    ShannonSimpsonExcel: TAdvSpreadGrid;
    AdvSmoothPanel2: TAdvSmoothPanel;
    tituloHeader: TAdvSmoothLabel;
    Image2: TImage;
    Image3: TImage;
    AdvSmoothLabel6: TAdvSmoothLabel;
    panelRiquezaInterno: TAdvScrollBox;
    AdvSmoothLabel7: TAdvSmoothLabel;
    capasArbolRiqueza: TAdvSmoothComboBox;
    RiquezaExcel: TAdvSpreadGrid;
    Image5: TImage;
    Image4: TImage;
    GraficoMargalef: TChart;
    datosMargalef: TBarSeries;
    GraficoMenhinick: TChart;
    datosMenhinick: TBarSeries;
    SerieDominancia: TBarSeries;
    exportadorRiqueza: TAdvGridExcelIO;
    exportadorDiversidadBase: TAdvGridExcelIO;
    exportadorDiversidadExcel: TAdvGridExcelIO;
    exportadorIVIHeader: TAdvGridExcelIO;
    TeeGDIPlus1: TTeeGDIPlus;
    AdvGlowButton1: TAdvGlowButton;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure calcularIVI(itemIndex: Integer; tipo: String);
    procedure calcularDiversidad(itemIndex: Integer);
    procedure calcularRiqueza(itemIndex: Integer);
    procedure capasArbolIVIItemSelected(Sender: TObject; itemindex: Integer);
    procedure botonIVIEspeciesMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure capasArbolDiversidadItemSelected(Sender: TObject;
      itemindex: Integer);
    procedure MenuChanging(Sender: TObject; FromPage, ToPage: Integer;
      var AllowChange: Boolean);
    procedure capasArbolRiquezaItemClick(Sender: TObject; itemindex: Integer);
    procedure AdvGlowButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;


var
  Interfaz: TInterfaz;
  DirectorioProyecto,
  capa,
  frecQuery,
  abundQuery,
  domQuery,
  field: String;



implementation

{$R *.dfm}

procedure TInterfaz.capasArbolIVIItemSelected(Sender: TObject; itemindex: Integer);
begin
    // Si está seleccionado el boton de especies (Está de color verde)
    if botonIVIEspecies.Fill.Color = $0055B065 then begin
        calcularIVI(itemindex, 'Especies');

    end else if botonIVIFamilias.Fill.Color = $0055B065 then begin
        calcularIVI(itemindex, 'Familias');
    end;
end;



procedure TInterfaz.capasArbolRiquezaItemClick(Sender: TObject;
  itemindex: Integer);
begin
    calcularRiqueza(itemIndex);
end;



procedure TInterfaz.capasArbolDiversidadItemSelected(Sender: TObject;
  itemindex: Integer);
begin
    calcularDiversidad(itemIndex);
end;



procedure TInterfaz.FormCreate(Sender: TObject);

begin
    capasArbolIVI.Items.Clear;
    capasArbolDiversidad.Items.Clear;
    capasArbolRiqueza.Items.Clear;

    IVIExcelHeader.MergeCells(0, 0, 2, 2);
    IVIExcelHeader.Cells[1, 1] := 'Especie';
    IVIExcelHeader.MergeCells(2, 0, 2, 1);
    IVIExcelHeader.Cells[2, 0] := 'Frecuencia';
    IVIExcelHeader.Cells[2, 1] := 'Abs';
    IVIExcelHeader.Alignments[2, 1] := TAlignment.taCenter;
    IVIExcelHeader.Cells[3, 1] := 'Rel';
    IVIExcelHeader.Alignments[3, 1] := TAlignment.taCenter;
    IVIExcelHeader.MergeCells(4, 0, 2, 1);
    IVIExcelHeader.Cells[4, 0] := 'Dominancia';
    IVIExcelHeader.Cells[4, 1] := 'Abs';
    IVIExcelHeader.Alignments[4, 1] := TAlignment.taCenter;
    IVIExcelHeader.Cells[5, 1] := 'Rel';
    IVIExcelHeader.Alignments[5, 1] := TAlignment.taCenter;
    IVIExcelHeader.MergeCells(6, 0, 2, 1);
    IVIExcelHeader.Cells[6, 0] := 'Abundancia';
    IVIExcelHeader.Cells[6, 1] := 'Abs';
    IVIExcelHeader.Alignments[6, 1] := TAlignment.taCenter;
    IVIExcelHeader.Cells[7, 1] := 'Rel';
    IVIExcelHeader.Alignments[7, 1] := TAlignment.taCenter;
    IVIExcelHeader.MergeCells(8, 0, 1, 2);
    IVIExcelHeader.Cells[8, 0] := 'IVI (300%)';
    IVIExcelHeader.ColWidths[8] := 84;

    SerieDominancia.Clear;
    SerieFrecuencia.Clear;
    SerieAbundancia.Clear;

    IVIExcelData.SelectionRectangleColor := $00418D4F;
    ShannonSimpsonDatosBase.SelectionRectangleColor := $00418D4F;
    ShannonSimpsonExcel.SelectionRectangleColor := $00418D4F;
    RiquezaExcel.SelectionRectangleColor := $00418D4F;

    ShannonSimpsonDatosBase.Cells[1, 0] := 'IDParcela';
    ShannonSimpsonDatosBase.Cells[2, 0] := 'Especie';
    ShannonSimpsonDatosBase.Cells[3, 0] := '# Individuos';
    ShannonSimpsonDatosBase.Cells[4, 0] := 'Pi';
    ShannonSimpsonDatosBase.Cells[5, 0] := 'Pi^2';
    ShannonSimpsonDatosBase.Cells[6, 0] := 'Pi*Ln(Pi)';

    ShannonSimpsonExcel.Cells[1, 0] := 'IDParcela';
    ShannonSimpsonExcel.Cells[2, 0] := 'Indice de Simpson (D)';
    ShannonSimpsonExcel.Cells[3, 0] := 'Indice de Shannon (H)';

    RiquezaExcel.Cells[1, 0] := 'IDParcela';
    RiquezaExcel.Cells[2, 0] := 'N (# Indv)';
    RiquezaExcel.Cells[3, 0] := 'S (# Sp)';
    RiquezaExcel.Cells[4, 0] := 'Indice de Margalef (Dmg)';
    RiquezaExcel.Cells[5, 0] := 'Indice de Menhinick (Dmn)';


end;



procedure TInterfaz.MenuChanging(Sender: TObject; FromPage, ToPage: Integer;
  var AllowChange: Boolean);
begin
    if ToPage = 0 then begin
        tituloHeader.Caption.Text := 'Indice de valor de importancia';

    end else if ToPage = 1 then begin
        tituloHeader.Caption.Text := 'Shannon | Simpson';

    end else if ToPage = 2 then begin
        tituloHeader.Caption.Text := 'Margalef | Menhinick';
    end;

end;




procedure TInterfaz.AdvGlowButton1Click(Sender: TObject);
var
    archivo, carpeta, directorioExcel: String;
    imagenGrafic: TBitMap;

begin
    carpeta := 'Resultados indices ecologicos';
    CreateDir(carpeta);
    DeleteFile(carpeta + '\Resultados.xls');
    DeleteFile(carpeta + '\Grafico IVI.bmp');
    DeleteFile(carpeta + '\Grafico Margalef.bmp');
    DeleteFile(carpeta + '\Grafico Menhinick.bmp');
    archivo := carpeta + '\Resultados.xls';

    GraficoIVI.SaveToBitmapFile(carpeta + '\Grafico IVI.bmp');
    GraficoMargalef.SaveToBitmapFile(carpeta + '\Grafico Margalef.bmp');
    GraficoMenhinick.SaveToBitmapFile(carpeta + '\Grafico Menhinick.bmp');

    imagenGrafic := IVIExcelData.CreateBitmap(9, 0, True, TCellHAlign.haCenter, TCellVAlign.vaCenter);
    imagenGrafic.LoadFromFile(carpeta + '\Grafico IVI.bmp');

    imagenGrafic := RiquezaExcel.CreateBitmap(6, 0, True, TCellHAlign.haCenter, TCellVAlign.vaCenter);
    imagenGrafic.LoadFromFile(carpeta + '\Grafico Margalef.bmp');

    imagenGrafic := RiquezaExcel.CreateBitmap(14, 0, True, TCellHAlign.haCenter, TCellVAlign.vaCenter);
    imagenGrafic.LoadFromFile(carpeta + '\Grafico Menhinick.bmp');

    exportadorIVI.XLSExport(archivo, 'IVI', -1, 1, TInsertInSheet.InsertInSheet_OverwriteCells);
    exportadorIVIHeader.XLSExport(archivo, 'IVI', -1, 1, TInsertInSheet.InsertInSheet_OverwriteCells);
    exportadorRiqueza.XLSExport(archivo, 'Riqueza',-1, 1, TInsertInSheet.InsertInSheet_OverwriteCells);
    exportadorDiversidadBase.XLSExport(archivo, 'Diversidad',-1, 1, TInsertInSheet.InsertInSheet_OverwriteCells);
    exportadorDiversidadExcel.XLSExport(archivo, 'Diversidad',-1, 1, TInsertInSheet.InsertInSheet_OverwriteCells);


    messagedlg('El análisis ha sido exportado a: ' + #13#10#13#10 +
                DirectorioProyecto + '\' + carpeta
     ,mtInformation,
                              [mbOK], 0);
end;




procedure TInterfaz.botonIVIEspeciesMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

var
    Btn: TAdvSmoothPanel;
begin

    Btn := Sender as TAdvSmoothPanel;

    if (Btn.Fill.Color = clWhite) then begin
        // activando boton de especies
        if (Btn.Caption.Text = 'Especies') then begin
        // cambiando primera columna de la tabla a Especies
            IVIExcelHeader.Cells[1, 1] := 'Especies';
            GraficoIVI.Title.Caption := 'IVI para las especies encontradas en el muestreo';
            GraficoIVI.BottomAxis.Title.Caption := 'Especies';
    //      Cambiando boton Especies
            BotonIVIEspecies.Fill.Color := $0055B065;
            BotonIVIEspecies.Fill.ColorTo := $0055B065;
            BotonIVIEspecies.Fill.BorderColor := $00418D4F;

            BotonIVIEspecies.Caption.Font.Name := 'Roboto bk';
            BotonIVIEspecies.Caption.ColorStart := clWhite;
            BotonIVIEspecies.Caption.ColorEnd := clWhite;

    //      Cambiando boton de familias
            BotonIVIFamilias.Fill.Color := clWhite;
            BotonIVIFamilias.Fill.ColorTo := clWhite;
            BotonIVIFamilias.Fill.BorderColor := $00C4C6C8;

            BotonIVIFamilias.Caption.Font.Name := 'Roboto';
            BotonIVIFamilias.Caption.ColorStart := clBlack;
            BotonIVIFamilias.Caption.ColorEnd := clBlack;

            if capasArbolIVI.SelectedItemIndex <> -1 then begin
                calcularIVI(capasArbolIVI.SelectedItemIndex, 'Especies');

            end;

        // activando boton de familias
        end else if (Btn.Caption.Text = 'Familias') then begin
            // cambiando primera columna de la tabla a Familias
            IVIExcelHeader.Cells[1, 1] := 'Familias';
            GraficoIVI.Title.Caption := 'IVI para las familias encontradas en el muestreo';
            GraficoIVI.BottomAxis.Title.Caption := 'Familias';

    //      Cambiando boton Especies
            BotonIVIFamilias.Fill.Color := $0055B065;
            BotonIVIFamilias.Fill.ColorTo := $0055B065;
            BotonIVIFamilias.Fill.BorderColor := $00418D4F;

            BotonIVIFamilias.Caption.Font.Name := 'Roboto bk';
            BotonIVIFamilias.Caption.ColorStart := clWhite;
            BotonIVIFamilias.Caption.ColorEnd := clWhite;

    //      Cambiando boton de familias
            BotonIVIEspecies.Fill.Color := clWhite;
            BotonIVIEspecies.Fill.ColorTo := clWhite;
            BotonIVIEspecies.Fill.BorderColor := $00C4C6C8;

            BotonIVIEspecies.Caption.Font.Name := 'Roboto';
            BotonIVIEspecies.Caption.ColorStart := clBlack;
            BotonIVIEspecies.Caption.ColorEnd := clBlack;

            if capasArbolIVI.SelectedItemIndex <> -1 then begin
                calcularIVI(capasArbolIVI.SelectedItemIndex, 'Familias');

            end;

        end;
    end;

end;




procedure TInterfaz.calcularIVI(itemindex: Integer; tipo: String);
var
    fila, numSp, i: Integer;

begin
    // capa que seleccíono la persona
    capa := capasArbolIVI.Items[itemIndex].Caption;

    if tipo = 'Especies' then begin
        field := 'Especie';

    end else if tipo = 'Familias' then begin
        field := 'familia';

    end;

    // PENDIENTE EN LO DE WHERE IS NOT NULL CARAJO
    frecQuery := Format('SELECT %s, COUNT(*) AS presencia_parcelas '+
                        'FROM ( SELECT DISTINCT %s, %s FROM %s WHERE %s IS NOT NULL) ' +
                        'GROUP BY %s '
                        ,
                        [field,
                        field,
                        'ID'+NombreCapaParcela,
                        capa,
                        field,
                        field]);


    abundQuery := Format('SELECT %s, Count(%s) AS num_ind '+
                        'FROM %s  WHERE %s IS NOT NULL '+
                        'GROUP BY %s '
                         ,
                        [field,
                        field,
                        capa,
                        field,
                        field]);

    domQuery := Format('SELECT %s, SUM(  (DBH_mm/1000)^2 * (3.14159265358979/4)  )  as area_basal_m2 '+
                        'FROM %s '+
                        'GROUP BY %s ',
                        [field,
                        capa,
                        field]);


    with DataModule1 do begin
        SQLQuery.Open(frecQuery);
        SQLQuery.FetchAll;
        numSp := SQLQuery.Table.Rows.Count;
        IVIExcelData.RowCount := numSp;

        for fila := 1 to numSp - 1 do begin
            // Lleando la columna de especies
            IVIExcelData.Cells[1, fila] := SQLQuery.Table.Rows[fila-1].GetData(field);
            // Frecuencia absoluta
            IVIExcelData.Cells[2, fila] := SQLQuery.Table.Rows[fila-1].GetData('presencia_parcelas');

        end;

        for fila := 1 to numSp - 1 do begin
            // Frecuencia relativa
            IVIExcelData.Cells[3, fila] := (SQLQuery.Table.Rows[fila-1].GetData('presencia_parcelas') / IVIExcelData.ColumnSum(2)) * 100;

        end;

        SQLQuery.Open(abundQuery);
        SQLQuery.FetchAll;

        for fila := 1 to numSp - 1 do begin
            // Dominancia absoluta
            IVIExcelData.Cells[4, fila] := SQLQuery.Table.Rows[fila-1].GetData('num_ind');

        end;

        for fila := 1 to numSp - 1 do begin
            // Dominancia relativa
            IVIExcelData.Cells[5, fila] := (SQLQuery.Table.Rows[fila-1].GetData('num_ind') / IVIExcelData.ColumnSum(4)) * 100;

        end;

        SQLQuery.Open(domQuery);
        SQLQuery.FetchAll;

        for fila := 1 to numSp - 1 do begin
            // abundancia absoluta
            IVIExcelData.Cells[6, fila] := SQLQuery.Table.Rows[fila-1].GetData('area_basal_m2');

        end;

        for fila := 1 to numSp - 1 do begin
            // abundancia relativa
            IVIExcelData.Cells[7, fila] := (SQLQuery.Table.Rows[fila-1].GetData('area_basal_m2') / IVIExcelData.ColumnSum(6)) * 100;

        end;

        for fila := 1 to numSp - 1 do begin
            // IVI
            IVIExcelData.Cells[8, fila] := (IVIExcelData.CellValue[3, fila] + IVIExcelData.CellValue[5, fila] + IVIExcelData.CellValue[7, fila]);
        end;
    end;

    // Graficar
    SerieDominancia.Clear;
    SerieFrecuencia.Clear;
    SerieAbundancia.Clear;

//    GraficoIVI.Height := Trunc((IVIExcelData.Cols[1].Count - 1) * 100 / 6);
    GraficoIVI.Height := Trunc(281 * 1.8);  // Tamaño normal x 1.6
    panelParaScroll.Height := GraficoIVI.Height + 10;

    IVIExcelData.Sort(8, sdDescending);

    for i := 1 to IVIExcelData.Cols[1].Count - 1 do begin
        SerieDominancia.Add(IVIExcelData.CellToReal(5, i), '', clTeeColor);
        SerieFrecuencia.Add(IVIExcelData.CellToReal(3, i), '', clTeeColor);
        SerieAbundancia.Add(IVIExcelData.CellToReal(7, i), '', clTeeColor);
        SerieAbundancia.Labels[i-1] := IVIExcelData.Cells[1, i];
    end;

end;




procedure TInterfaz.calcularDiversidad(itemIndex: Integer);
var
    i, numParcelas: Integer;
    condAccessString, field: String;

begin
    field := 'Especie';
    capa := capasArbolDiversidad.Items[itemIndex].Caption;

    with dataModule1 do begin
        // Calcula numero de individuos por cada ID de parcela
        SQLQuery.Open(Format('SELECT ID%s, Count(%s) AS num_indv ', [NombreCapaParcela, field])+
                       'FROM ' + capa +
                       Format(' GROUP BY ID%s', [NombreCapaParcela]));

        SQLQuery.FetchAll;

        // Esto es para las condicionales anidadas en el SQL final de Simpson y Shannon
        condAccessString := Format('IIf(ID%s=%s, %s, ', [NombreCapaParcela,
                                                       SQLQuery.Table.Rows[0].GetData('ID'+NombreCapaParcela),
                                                       SQLQuery.Table.Rows[0].GetData('num_indv')
                                                       ]);

        numParcelas := SQLQuery.Table.Rows.Count;

        if numParcelas > 1 then begin
            for i := 1 to numParcelas - 2 do begin
                condAccessString := condAccessString +
                                    Format('IIf(ID%s=%s, %s, ', [NombreCapaParcela,
                                                               SQLQuery.Table.Rows[i].GetData('ID'+NombreCapaParcela),
                                                               SQLQuery.Table.Rows[i].GetData('num_indv')
                                                               ]);
            end;

            condAccessString := condAccessString + Format(' %s', [SQLQuery.Table.Rows[numParcelas-1].GetData('num_indv')]);

        end else begin
            condAccessString := condAccessString + ')';
        end;

        // Terminando de configurar la condicional
        condAccessString := condAccessString + DupeString(')', numParcelas-1);

//        ShowMessage(condAccessString);

        // Query que calcula la tabla base para el calculo de Shannon y Simpson
        SQLQuery.Open(
                Format('SELECT ID%s, %s, Count(%s) as num_indv,', [NombreCapaParcela, field, field])+
                                                               Format('Count(%s) / %s AS Pi,', [field, condAccessString])+
                                                               'Pi*Pi AS pi_al_2,' +
                                                               'Pi * Log(pi) AS pi_x_ln_pi' +
                Format(' FROM %s', [capa]) +
                Format(' GROUP BY ID%s, %s ', [NombreCapaParcela, field]) +
                Format(' ORDER BY ID%s', [NombreCapaParcela])
        );

        SQLQuery.FetchAll;

        ShannonSimpsonDatosBase.RowCount := SQLQuery.Table.Rows.Count + 1;
        ShannonSimpsonDatosBase.Cells[1, 0] := 'IDParcela';
        ShannonSimpsonDatosBase.Cells[2, 0] := 'Especie';
        ShannonSimpsonDatosBase.Cells[3, 0] := '# Individuos';
        ShannonSimpsonDatosBase.Cells[4, 0] := 'Pi';
        ShannonSimpsonDatosBase.Cells[5, 0] := 'Pi^2';
        ShannonSimpsonDatosBase.Cells[6, 0] := 'Pi*Ln(Pi)';

        // Llenando los datos de la tabla base de los indices de shannon y smpson
        for i := 0 to SQLQuery.Table.Rows.Count - 1 do begin
            ShannonSimpsonDatosBase.Cells[1, i+1] := SQLQuery.Table.Rows[i].GetData(Format('ID%s', [NombreCapaParcela]));
            ShannonSimpsonDatosBase.Cells[2, i+1] := SQLQuery.Table.Rows[i].GetData('Especie');
            ShannonSimpsonDatosBase.Cells[3, i+1] := SQLQuery.Table.Rows[i].GetData(2); // '# Individuos'
            ShannonSimpsonDatosBase.Cells[4, i+1] := SQLQuery.Table.Rows[i].GetData(3); // 'Pi'
            ShannonSimpsonDatosBase.Cells[5, i+1] := SQLQuery.Table.Rows[i].GetData(4); // 'Pi^2'
            ShannonSimpsonDatosBase.Cells[6, i+1] := SQLQuery.Table.Rows[i].GetData(5); // 'Pi*Ln(Pi)'
        end;

        // Query que calcula los indices de shannon y simpson de la base de datos
        SQLQuery.Open(
            Format('SELECT ID%s, SUM(pi_al_2) AS Simpson_Sumatoria_pi_as_2, -SUM(pi_x_ln_pi) AS Shannon_menos_Sumatoria_ln_pi', [NombreCapaParcela]) +
            ' FROM (' +
                Format('SELECT ID%s, %s, Count(%s) as num_indv,', [NombreCapaParcela, field, field])+
                                                               Format('Count(%s) / %s AS Pi,', [field, condAccessString])+
                                                               'Pi*Pi AS pi_al_2,' +
                                                               'Pi * Log(pi) AS pi_x_ln_pi' +
                Format(' FROM %s', [capa]) +
                Format(' GROUP BY ID%s, %s ', [NombreCapaParcela, field]) +
                Format(' ORDER BY ID%s', [NombreCapaParcela])
            + ' ) ' +
            Format(' GROUP BY ID%s', [NombreCapaParcela])
        );

        SQLQuery.FetchAll;

        // Configurando nombre de columnas en excel de Shannon y Simpson
        ShannonSimpsonExcel.RowCount := SQLQuery.Table.Rows.Count + 7;
        ShannonSimpsonExcel.Cells[1, 0] := 'IDParcela';
        ShannonSimpsonExcel.Cells[2, 0] := 'Indice de Simpson (D)';
        ShannonSimpsonExcel.Cells[3, 0] := 'Indice de Shannon (H)';

        for i := 0 to SQLQuery.Table.Rows.Count - 1 do begin
            // Llenando excel con los indices
            ShannonSimpsonExcel.Cells[1, i+1] := SQLQuery.Table.Rows[i].GetData(Format('ID%s', [NombreCapaParcela]));
            ShannonSimpsonExcel.Cells[2, i+1] := SQLQuery.Table.Rows[i].GetData(1); // Indice de Simpson
            ShannonSimpsonExcel.Cells[3, i+1] := SQLQuery.Table.Rows[i].GetData(2); // Indice de Shannon
        end;
    end;

end;




procedure TInterfaz.calcularRiqueza(itemIndex: Integer);
var
    i: Integer;

begin
    field := 'Especie';
    capa := capasArbolDiversidad.Items[itemIndex].Caption;

    with dataModule1 do begin

        SQLQuery.Open(
            Format('SELECT S.ID%s AS IDParcela, num_indv, num_sp, (num_sp - 1) / Log(num_indv) as Margalef, (num_sp / num_indv^0.5) as Menhinick ', [NombreCapaParcela]) +
            ' FROM (' +
                Format(' SELECT * FROM ( SELECT ID%s, Count(%s) as num_sp ', [NombreCapaParcela, field]) +
                Format(' FROM ( SELECT DISTINCT ID%s, %s FROM %s ) ', [NombreCapaParcela, field, capa]) +
                Format(' GROUP BY ID%s ) AS S ', [NombreCapaParcela]) +

                ' INNER JOIN ' +

                Format(' ( SELECT ID%s, Count(ID) as num_indv ', [NombreCapaParcela]) +
                Format(' FROM %s ', [capa]) +
                Format(' GROUP BY ID%s ) AS N ', [NombreCapaParcela]) +

                Format(' ON S.ID%s = N.ID%s ', [NombreCapaParcela, NombreCapaParcela]) +
            ')'
        );

        SQLQuery.FetchAll;

        // Configurando nombre de columnas en excel de Riqueza
        RiquezaExcel.RowCount := SQLQuery.Table.Rows.Count + 7;
        RiquezaExcel.Cells[1, 0] := 'IDParcela';
        RiquezaExcel.Cells[2, 0] := 'N (# Indv)';
        RiquezaExcel.Cells[3, 0] := 'S (# Sp)';
        RiquezaExcel.Cells[4, 0] := 'Indice de Margalef (Dmg)';
        RiquezaExcel.Cells[5, 0] := 'Indice de Menhinick (Dmn)';


        for i := 0 to SQLQuery.Table.Rows.Count - 1 do begin
            // Llenando excel con los indices
            RiquezaExcel.Cells[1, i+1] := SQLQuery.Table.Rows[i].GetData('IDParcela');
            RiquezaExcel.Cells[2, i+1] := SQLQuery.Table.Rows[i].GetData(1); // Num indv
            RiquezaExcel.Cells[3, i+1] := SQLQuery.Table.Rows[i].GetData(2); // Num Sp
            RiquezaExcel.Cells[4, i+1] := SQLQuery.Table.Rows[i].GetData(3); // Margalef

            datosMargalef.Add(RiquezaExcel.CellToReal(4, i+1), '', clTeeColor);
            datosMargalef.Labels[i] := SQLQuery.Table.Rows[i].GetData('IDParcela');

            RiquezaExcel.Cells[5, i+1] := SQLQuery.Table.Rows[i].GetData(4); // Menhinick
            datosMenhinick.Add(RiquezaExcel.CellToReal(5, i+1), '', clTeeColor);
            datosMenhinick.Labels[i] := SQLQuery.Table.Rows[i].GetData('IDParcela');

        end;
    end;

end;

end.
