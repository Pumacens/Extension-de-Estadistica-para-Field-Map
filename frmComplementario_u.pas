unit frmComplementario_u;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  AdvSmoothListBox;

type
  Tmiscelaneos = class(TForm)
    _ItemsIconos: TAdvSmoothListBox;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  miscelaneos: Tmiscelaneos;

implementation

{$R *.dfm}

end.
