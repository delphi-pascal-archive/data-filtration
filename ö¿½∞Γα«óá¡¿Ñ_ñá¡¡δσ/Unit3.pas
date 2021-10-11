unit Unit3;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, StdCtrls;

type
  TForm3 = class(TForm)
    Label1: TLabel;
    Edit1: TEdit;
    Label2: TLabel;
    Edit2: TEdit;
    Label3: TLabel;
    Edit3: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

uses Unit1;

{$R *.dfm}

procedure TForm3.SpeedButton1Click(Sender: TObject);
var
val1,val2,val3:integer;
begin
if (TryStrToInt(Form3.Edit1.Text,val1)=true) and
   (TryStrToInt(Form3.Edit2.Text,val2)=true) and
   (TryStrToInt(Form3.Edit3.Text,val3)=true) then begin
 val1:=StrToInt(Form3.Edit1.Text);
 val2:=StrToInt(Form3.Edit2.Text);
 val3:=StrToInt(Form3.Edit3.Text);
 if (val1<=30) and (val1>0)
               and (val2>0)    and (val2<=1000)
               and (val3<=100) and (val3>0) then begin
  Form1.NewDoc;
  Form1.GenTables(val1,val2,val3);
  form3.Close;
  exit;
 end;
end;
showmessage('Неверные данные');
form3.Close;
end;

procedure TForm3.SpeedButton2Click(Sender: TObject);
begin
form3.Close;
end;

end.
