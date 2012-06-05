unit osIPersistable;

interface

uses
  Classes;

const
  IID_Persistable: TGUID = '{F66DB7BE-FB3A-4BE1-8A4F-5C438D4D14AA}';

type
  TNormalizeMode = (nrmd_None, nrmd_UTF8, nrmd_UTF16);

type
  IPersistable = interface
  ['{F66DB7BE-FB3A-4BE1-8A4F-5C438D4D14AA}']
    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(Stream: TStream);
    procedure LoadFromString(const S: string);
    procedure SaveToFile(const FileName: string; NormalizeMode: TNormalizeMode = nrmd_UTF16);
    procedure SaveToStream(Stream: TStream; NormalizeMode: TNormalizeMode = nrmd_UTF16);
    procedure SaveToString(var S: string; NormalizeMode: TNormalizeMode = nrmd_UTF16);
  end;

implementation

end.