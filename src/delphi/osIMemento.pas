unit osIMemento;

interface

uses
  Classes;

const
  IID_Memento: TGUID = '{8F50CE52-4311-45A6-A928-CA15E25ED249}';

type
  IMemento = interface;
  IMementoArray = array of IMemento;

  IMemento = interface
  ['{8F50CE52-4311-45A6-A928-CA15E25ED249}']
    function CreateChild(const Type_: string): IMemento;
    function CreateChildSmart(const Type_: string): IMemento;
    function GetBoolean(const Key: string; var Value: Boolean): Boolean;
    function GetChild(const Type_: string): IMemento;
    function GetChildren(const Type_: string): IMementoArray;
    function GetChildFromPath(const Path: string): IMemento;
    function GetChildrenFromPath(const Path: string): IMementoArray;
    function GetDouble(const Key: string; var Value: Double): Boolean;
    function GetGUID(const Key: string; var Value: TGUID): Boolean;
    function GetInteger(const Key: string; var Value: Integer): Boolean;
    function GetName: string;
    function GetRoot: IMemento;
    function GetParent: IMemento;
    function GetString(const Key: string; var Value: WideString): Boolean;
    function GetTextData(var Value: WideString): Boolean;
    procedure PutBoolean(const Key: string; Value: Boolean);
    procedure PutDouble(const Key: string; Value: Double);
    procedure PutGUID(const Key: string; Value: TGUID);
    procedure PutInteger(const Key: string; Value: Integer);
    procedure PutString(const Key: string; const Value: WideString);
    procedure PutTextData(const Data: WideString);
  end;

implementation

end.