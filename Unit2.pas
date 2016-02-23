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
  '���: ' + inn.Text + #13 +
  '���: ' + kpp.Text + '���. 001'+ #13;
  Doc.Paragraphs.Item(3).Alignment := wdAlignParagraphRight;
  Doc.Paragraphs.Item(3).Range.Font.Bold := 1;
  Doc.Paragraphs.Item(3).Range.Text :=
  '����� � ��-2'+ #13 +
  '����� �� ��� 1110051' + #13;
  Doc.Paragraphs.Item(5).Format.Alignment:=wdAlignParagraphCenter;
  Doc.Paragraphs.Item(5).LeftIndent:=WA.CentimetersToPoints(0) ;
  Doc.Paragraphs.Item(5).Range.Text :=
  '����������� '+ #13 +
  '� ������ � ����� ����������� ��� ��������������� ��������������� � �������� ' + #13 +
  '����������� ��������� ����� � ��������� ������ �� ������� ������������� ���� '+ #13 +
  '������������������� ������������, � ��������� �������� ���������� �������� ���� (1)'+ #13;
  Doc.Paragraphs.Item(9).Range.Font.Bold := 0;
  Doc.Paragraphs.Item(9).Alignment := wdAlignParagraphRight;
  Doc.Paragraphs.Item(9).Range.Text :=
  '�������������� � ��������� ����� (���): ' + labelededit3.Text + #13;
  Doc.Paragraphs.Item(10).Alignment := wdAlignParagraphLeft;
  Doc.Paragraphs.Item(10).Range.Text :=
  '�������� � ����������� ����� '+ #13 +
  '����������� / �������������� ���������������:'+ #13 +
  labelededit4.Text + #13 +
  '(������ ������������ ����������� / �������, ���, �������� (2) ��������������� ���������������)' + #13;
  Doc.Paragraphs.Item(13).Format.Alignment:=wdAlignParagraphCenter;
  Doc.Paragraphs.Item(14).Range.Text :=
  '���� (3): '+ labelededit5.Text +' /������ (4): '+ labelededit11.Text + #13;
  Doc.Paragraphs.Item(15).Range.Text :=
  '���� ����������� ������������� ������������������� ������������, � ��������� ������� ���������� �������� ����: ' + DateToStr(DateTimePicker1.DateTime)+ #13;
  Doc.Paragraphs.Item(16).Range.Text :=
  '��������� ����������� ���������� �� 1 �������� � ����������� �������������� ���������� ��� �� ����� (5) �� ' +inttostr(spinedit1.value)+ ' ������'+ #13;
  Doc.Tables.Add(Doc.Paragraphs.Item(17).Range,1,2,wdWord9TableBehavior,wdAutoFitFixed);
    Doc.Tables.Item(1).Cell(1,1).Range.text:=
    '������������� � ������� ��������, ��������� � ��������� �����������, �����������: '+ #13 +
    combobox1.Text+ #13 +
    labelededit6.Text + #13 +
    '(�������, ���, �������� (2) ������������ ����������� ���� ������������� ���������)'+ #13+
    '��� (6): '+ inn.Text + #13 +
    '����� ����������� ��������: '+ labelededit7.text + #13 +
    'E-Mail: '+ labelededit8.Text + #13 +
    '�������: __________ '+ '����: '+ DateToStr(DateTimePicker2.DateTime)+ #13 +
    '������������ ���������, ��������������� ���������� �������������: '+ #13 +
    labelededit9.text;
    Doc.Paragraphs.Item(17).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Paragraphs.Item(17).Range.Font.Bold := 1;
    Doc.Paragraphs.Item(20).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Paragraphs.Item(25).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Tables.Item(1).Cell(1,2).Range.text:=
    '����������� ���������� ���������� ������'+ #13 +
    '�������� � �������������� �����������'+ #13 +
    '������ ����������� ������������� (���): '+ labelededit1.Text + #13 +
    '�� 1 �������� � ����������� ����� ��������� (5)'+ #13 +
    '�� '+inttostr(spinedit2.value)+ ' ������'+ #13 +
    '���� ������������� �����������: '+ DateToStr(DateTimePicker3.DateTime)+ #13+
    '���������������� �� �: '+ labelededit2.Text+ #13 +
    '���: '+ labelededit10.text + ' �������: __________'+ #13;
    Doc.Paragraphs.Item(27).Format.Alignment:=wdAlignParagraphCenter;
    Doc.Paragraphs.Item(27).Range.Font.Bold := 1;
    Doc.Paragraphs.Item(28).Format.Alignment:=wdAlignParagraphCenter;
    wa.Selection.EndKey(wdstory,emptyparam);
  Doc.Paragraphs.item(37).Range.Text:=   #13+
  '1) �������������� � ������ ����������� ������������� ���� ����� ������������������� ������������ � �������������� �������� ������������� ��������, � ��������� ������� ���������� �������� ����.'+ #13+
  '2) �������� ����������� ��� �������.'+ #13+
  '3) ����������� ���������� ������������.'+ #13+
  '4) ����������� �������������� ����������������.'+ #13+
  '5) � ����������� ����������� ����� ���������, ��������������� ���������� �������������.'+ #13+
  '6) ����������� � ��������� ���������� ���, ������� ��������, �������������� ���������� ��� (������������� � ���������� �� ���� � ��������� ������, ������� � �������� ���������� ���������� ���������), � ������������ ��� ������ � ������������� �������.';

  WA.Selection.WholeStory;  //�������� ���
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
