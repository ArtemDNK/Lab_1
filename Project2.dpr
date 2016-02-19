program Project2;

uses
  Vcl.Forms,
  Unit2 in 'Unit2.pas' {Form2},
  Word_TLB in 'Word_TLB.pas',
  VBIDE_TLB in 'VBIDE_TLB.pas';

{$R *.res}

begin
  Vcl.Forms.Application.Initialize;
  Vcl.Forms.Application.MainFormOnTaskbar := True;
  Vcl.Forms.Application.CreateForm(TForm2, Form2);
  Vcl.Forms.Application.Run;
end.
