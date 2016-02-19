unit Unit2;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Word_tlb, VBIDE_tlb, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Samples.Spin;

type
  TForm2 = class(TForm)
    Button1: TButton;
    inn: TLabeledEdit;
    kpp: TLabeledEdit;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    LabeledEdit3: TLabeledEdit;
    LabeledEdit4: TLabeledEdit;
    LabeledEdit5: TLabeledEdit;
    DateTimePicker1: TDateTimePicker;
    Label1: TLabel;
    SpinEdit1: TSpinEdit;
    Label2: TLabel;
    ComboBox1: TComboBox;
    Label3: TLabel;
    LabeledEdit6: TLabeledEdit;
    LabeledEdit7: TLabeledEdit;
    LabeledEdit8: TLabeledEdit;
    DateTimePicker2: TDateTimePicker;
    Label4: TLabel;
    LabeledEdit9: TLabeledEdit;
    LabeledEdit1: TLabeledEdit;
    SpinEdit2: TSpinEdit;
    Label5: TLabel;
    DateTimePicker3: TDateTimePicker;
    Label6: TLabel;
    LabeledEdit2: TLabeledEdit;
    LabeledEdit10: TLabeledEdit;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.Button1Click(Sender: TObject);
var
wa:WordApplication;
begin
wa:=CoWordApplication.Create;
end;

end.
