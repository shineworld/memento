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
    function LoadFromFile(const FileName: string): Boolean;
    function LoadFromStream(Stream: TStream): Boolean;
    function LoadFromString(const S: string): Boolean;
    function SaveToFile(const FileName: string; NormalizeMode: TNormalizeMode = nrmd_UTF16; Crypted: Boolean = False): Boolean;
    function SaveToStream(Stream: TStream; NormalizeMode: TNormalizeMode = nrmd_UTF16): Boolean;
    function SaveToString(var S: string; NormalizeMode: TNormalizeMode = nrmd_UTF16): Boolean;
  end;

implementation

end.