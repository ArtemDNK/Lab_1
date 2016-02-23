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
    LabeledEdit11: TLabeledEdit;
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
Doc: WordDocument;
begin
WA:=CoWordApplication.Create;
WA.Visible := True;
Doc := WA.Documents.Add('Normal', False, EmptyParam, True);

  Doc.Paragraphs.Item(1).LeftIndent:=WA.CentimetersToPoints(5) ;
  Doc.Paragraphs.Item(1).Range.Text :=
  'ИНН: ' + inn.Text + #13 +
  'КПП: ' + kpp.Text + 'Стр. 001'+ #13;
  Doc.Paragraphs.Item(3).Alignment := wdAlignParagraphRight;
  Doc.Paragraphs.Item(3).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(3).Range.Text :=
  'Форма № ТС-2'+ #13 +
  'Форма по КНД 1110051' + #13;
  Doc.Paragraphs.Item(5).Format.Alignment:=wdAlignParagraphCenter;
  Doc.Paragraphs.Item(5).LeftIndent:=WA.CentimetersToPoints(0) ;
  Doc.Paragraphs.Item(5).Range.Text :=
  'Уведомление '+ #13 +
  'о снятии с учета организации или индивидуального предпринимателя в качестве ' + #13 +
  'плательщика торгового сбора в налоговом органе по объекту осуществления вида '+ #13 +
  'предпринимательской деятельности, в отношении которого установлен торговый сбор (1)'+ #13;
  Doc.Paragraphs.Item(9).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(9).Alignment := wdAlignParagraphRight;
  Doc.Paragraphs.Item(9).Range.Text :=
  'Представляется в налоговый орган (код): ' + labelededit3.Text + #13;
  Doc.Paragraphs.Item(10).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(10).Range.Text :=
  'Сведения о плательщике сбора '+ #13 +
  'Организация / индивидуальный предприниматель:'+ #13 +
  labelededit4.Text + #13 +
  '(полное наименование организации / фамилия, имя, отчество (2) индивидуального предпринимателя)' + #13;
  Doc.Paragraphs.Item(13).Format.Alignment:=wdAlignParagraphCenter;
  Doc.Paragraphs.Item(14).Range.Text :=
  'ОГРН (3): '+ labelededit5.Text +' /ОГРНИП (4): '+ labelededit11.Text + #13;
  Doc.Paragraphs.Item(15).Range.Text :=
  'Дата прекращения осуществления предпринимательской деятельности, в отношении которой установлен торговый сбор: ' + DateToStr(DateTimePicker1.DateTime)+ #13;
  Doc.Paragraphs.Item(16).Range.Text :=
  'Настоящее уведомление составлено на 1 странице с приложением подтверждающих документов или их копий (5) на ' +inttostr(spinedit1.value)+ ' листах'+ #13;
  Doc.Tables.Add(Doc.Paragraphs.Item(17).Range,1,2,wdWord9TableBehavior,wdAutoFitFixed);
    Doc.Tables.Item(1).Cell(1,1).Range.text:=
    'Достоверность и полноту сведений, указанных в настоящем уведомлении, подтверждаю: '+ #13 +
    combobox1.Text+ #13 +
    labelededit6.Text + #13 +
    '(фамилия, имя, отчество (2) руководителя организации либо представителя полностью)'+ #13+
    'ИНН (6): '+ inn.Text + #13 +
    'Номер контактного телефона: '+ labelededit7.text + #13 +
    'E-Mail: '+ labelededit8.Text + #13 +
    'Подпись: __________ '+ 'Дата: '+ DateToStr(DateTimePicker2.DateTime)+ #13 +
    'Наименование документа, подтверждающего полномочия представителя: '+ #13 +
    labelededit9.text;
    Doc.Paragraphs.Item(17).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Paragraphs.Item(17).Range.Font.Bold := 1;
    Doc.Paragraphs.Item(20).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Paragraphs.Item(25).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Tables.Item(1).Cell(1,2).Range.text:=
    'Заполняется работником налогового органа'+ #13 +
    'Сведения о предоставлении уведомления'+ #13 +
    'Данное уведомление предоставлено (код): '+ labelededit1.Text + #13 +
    'на 1 странице с приложением копии документа (5)'+ #13 +
    'на '+inttostr(spinedit2.value)+ ' листах'+ #13 +
    'Дата представления уведомления: '+ DateToStr(DateTimePicker3.DateTime)+ #13+
    'Зарегистрировано за №: '+ labelededit2.Text+ #13 +
    'ФИО: '+ labelededit10.text + ' Подпись: __________'+ #13;
    Doc.Paragraphs.Item(27).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Paragraphs.Item(27).Range.Font.Bold := 1;
    Doc.Paragraphs.Item(28).Format.Alignment:=wdAlignParagraphCenter;
    wa.Selection.EndKey(wdstory,emptyparam);
  Doc.Paragraphs.item(37).Range.Text:=   #13+
  '1) Представляется в случае прекращения осуществления всех видов предпринимательской деятельности с использованием объектов осуществления торговли, в отношении которых установлен торговый сбор.'+ #13+
  '2) Отчество указывается при наличии.'+ #13+
  '3) Заполняется российской организацией.'+ #13+
  '4) Заполняется индивидуальным предпринимателем.'+ #13+
  '5) К уведомлению прилагается копия документа, подтверждающего полномочия представителя.'+ #13+
  '6) Заполняется в отношении физических лиц, имеющих документ, подтверждающий присвоение ИНН (Свидетельство о постановке на учет в налоговом органе, отметка в паспорте гражданина Российской Федерации), и использующих ИНН наряду с персональными данными.';

  WA.Selection.WholeStory;  //выделить все
  WA.Selection.ParagraphFormat.LineSpacing := WA.LinesToPoints(0.9);
  WA.Selection.Font.Name:= 'Calibri';
  WA.Selection.Font.Size:= 11.5;
  WA.Selection.Font.Color:=clBlack;

  //WA.Selection.Range()
  //WA.Selection.Font.Superscript = wdToggle;

  Doc.Paragraphs.Item(13).Range.Font.Size:=9;
  Doc.Paragraphs.Item(20).Range.Font.Size:=9;
  Doc.Paragraphs.Item(43).Range.Font.Size:=9;
  Doc.Paragraphs.Item(38).Range.Font.Size:=9;
  Doc.Paragraphs.Item(39).Range.Font.Size:=9;
  Doc.Paragraphs.Item(40).Range.Font.Size:=9;
  Doc.Paragraphs.Item(41).Range.Font.Size:=9;
  Doc.Paragraphs.Item(42).Range.Font.Size:=9;
end;

end.
