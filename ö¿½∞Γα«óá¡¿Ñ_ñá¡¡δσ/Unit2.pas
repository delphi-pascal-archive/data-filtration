unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, StdCtrls, CheckLst;

type
  TForm2 = class(TForm)
    CheckListBox1: TCheckListBox;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.SpeedButton1Click(Sender: TObject);
var
str:string;
begin
if InputQuery('Добавление исключения','Введите исключение:',str) then begin
 CheckListbox1.Items.Add(str);
 CheckListbox1.Checked[CheckListBox1.Count-1]:=true;
end;
end;

procedure TForm2.SpeedButton2Click(Sender: TObject);
begin
if CheckListbox1.ItemIndex>-1 then CheckListbox1.DeleteSelected;
end;

procedure TForm2.SpeedButton3Click(Sender: TObject);
var
i:integer;
begin
for i:=CheckListbox1.count-1 downto 0 do begin
 if CheckListbox1.Checked[i]=true then CheckListbox1.Items.Delete(i);
end;
end;

procedure TForm2.SpeedButton4Click(Sender: TObject);
begin
CheckListbox1.Clear;
end;

end.
