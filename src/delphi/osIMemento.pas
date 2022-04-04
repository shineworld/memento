unit osIMemento;

interface

const
  IID_Memento: TGUID = '{8F50CE52-4311-45A6-A928-CA15E25ED249}';

type
{$IFDEF UNICODE}
  TXMLString = string;
{$ELSE}
  TXMLString = WideString;
{$ENDIF}

type
  IMemento = interface;
  IMementoArray = array of IMemento;

  IMemento = interface
  ['{8F50CE52-4311-45A6-A928-CA15E25ED249}']
    function CreateChild(const Type_: string): IMemento;
    function CreateChildSmart(const Type_: string): IMemento;
    function GetBoolean(const Key: string; var Value: Boolean): Boolean;
    function GetCardinal(const Key: string; var Value: Cardinal): Boolean;
    function GetChild(const Type_: string): IMemento;
    function GetChildFromPath(const Path: string): IMemento;
    function GetChildren(const Type_: string): IMementoArray;
    function GetChildrenFromPath(const Path: string): IMementoArray;
    function GetDateTime(const Key: string; var Value: TDateTime): Boolean;
    function GetDouble(const Key: string; var Value: Double): Boolean;
    function GetGUID(const Key: string; var Value: TGUID): Boolean;
    function GetInteger(const Key: string; var Value: Integer): Boolean;
    function GetName: string;
    function GetParent: IMemento;
    function GetRoot: IMemento;
    function GetString(const Key: string; var Value: TXMLString): Boolean;
    function GetTextData(var Value: TXMLString): Boolean;
    procedure PutBoolean(const Key: string; Value: Boolean);
    procedure PutCardinal(const Key: string; Value: Cardinal);
    procedure PutDateTime(const Key: string; Value: TDateTime);
    procedure PutDouble(const Key: string; Value: Double);
    procedure PutGUID(const Key: string; const Value: TGUID);
    procedure PutInteger(const Key: string; Value: Integer);
    procedure PutString(const Key: string; const Value: TXMLString);
    procedure PutTextData(const Data: TXMLString);
  end;

implementation

end.
