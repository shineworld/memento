{**
 *  This version of XML Memento uses the MSXML4 SP2 ActiveX Library from Microsoft.
 *  https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms763742(v=vs.85)
 *
 *  The MSXML should be already installed by default in recents Windows OS versions as W10.
 *  However the control software installation procedure adds/updates and registers, if missing or olds, following DLLs:
 *
 *  msxml4.dll    MSXML 4.0 SP2             4.20.9876.0
 *  msxml4a.dll   MSXML 4.0 SP1 Resources   4.10.9404.0
 *  msxml4r.dll   MSXML 4.0 SP1 Resources   4.10.9404.0
 *
 *  Here the list of the availables MSXML versions:
 *  https://docs.microsoft.com/en-us/previous-versions/troubleshoot/msxml/list-of-xml-parser-versions
 *
 *  Here the MSXML Wikipedia page link with a more simple description and the Windows OSs compatibility list:
 *  https://en.wikipedia.org/wiki/MSXML
 *
 *  IMPORTANT NOTES
 *  ===============
 *  This Class is born with BDS2006 in which string type is an alias for AnsiString, a pure string of ANSI chars.
 *  Encrypt and Decrypt was made to work with strings of ANSI chars (AnsiStrings).
 *  The field 'Key' of methods is designed to works fine only using strings of ANSI chars (AnsiStrings).
 *  The field 'Value' for GetString, PutString is designed to works fine only using strings of ANSI chars (AnsiStrings).
 *  The field 'Value' for GetData, PutData is designed to works fine only using strings of ANSI chars (AnsiStrings).
 *  All fields 'Value' before to be sent to XML component are encoded with EncodeStringToXMLString which converts any
 *  ANSI char out of range of Char($20)..Char($7F) into corresponding &#xHHHH; value when HHHH is a 4 bytes hexadecimal
 *  representation of ANSI char out of ASCII set. This is nessary because XML language char-set is limited and wildcards
 *  need to be used.
 *
 *  SYDNEY changed somethings because the string type is an alias for UnicodeString, a complex Unicode string type.
 *  Where string in BDS2006 is an AnsiString in SYDNEY become a UnicodeString.
 *  To reduce the amount of compiling hints and changes to move from AnsiString to UnicodeString we preferred to
 *  minimize code modifications keeping same rule in 'Key' and 'Value' use in code consumer.
 *  Unfortunately when Encrypt/Decrypt was made we missing to add a clean header byte with a version number to detects
 *  if compressed XML is pure ANSI string or UNICODE string. Yes a bad thing...
 *  At this point to maintain the compatibility with already saved, in BDS2006 version, encripted files also the XML
 *  output file is to intend AS pure ANSI XML file so Encription can continue to read legacy files.
 *
 *  This approach keep same features of old BDS2006 development system loosing new capabilities of full UnicodeSupport
 *  gained with SYDNEY.
 *
 **}
unit osXMLMemento;

interface

uses
  Classes,

  osIMemento,
  osIPersistable;

type
  { MSXML types }
  IXMLDOMDocument = Variant;
  IXMLDOMElement = Variant;
  IXMLDOMNode = Variant;
  IXMLDOMNodeList = Variant;
  IXMLDOMAttribute = Variant;
  IXMLDOMText = Variant;

type
  TXMLMemento = class(TInterfacedObject, IMemento, IPersistable)
  private
    FDocument: IXMLDOMDocument;
    FElement: IXMLDOMElement;
  private
    { IMemento }
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

    { IPersistable }
    function LoadFromFile(const FileName: string): Boolean;
    function LoadFromStream(Stream: TStream): Boolean;
    function LoadFromString(const S: string): Boolean;
    function SaveToFile(const FileName: string; NormalizeMode: TNormalizeMode = nrmd_UTF16; Crypted: Boolean = False): Boolean;
    function SaveToStream(Stream: TStream; NormalizeMode: TNormalizeMode = nrmd_UTF16): Boolean;
    function SaveToString(var S: string; NormalizeMode: TNormalizeMode = nrmd_UTF16): Boolean;

    { TXMLMemento }
    function GetElementNodeText: IXMLDOMText;
    function GetNormalizedDocument(NormalizeMode: TNormalizeMode): IXMLDOMDocument;
    constructor Create(Document: IXMLDOMDocument; Element: IXMLDOMElement);
  public
    class function CreateFromString(const S: string): IMemento; static;
  end;

function CreateReadRoot(const FileName: string): IMemento;
function CreateReadRootSmart(const FileName, RootName: string): IMemento;
function CreateWriteRoot(const Type_: string): IMemento;
function EncodeStringToXMLString(const S: TXMLString): TXMLString;
function DecodeXMLStringToString(const S: TXMLString): TXMLString;

implementation

uses
  Math,
  ComObj,
  Windows,
  SysUtils,
  Variants,

  osSysUtils,
  osExceptionUtils;

const
  TRUE_FALSE: array[Boolean] of string = ('false', 'true');

  NODE_ELEMENT                  = 1;
  NODE_ATTRIBUTE                = 2;
  NODE_TEXT                     = 3;
  NODE_CDATA_SECTION            = 4;
  NODE_ENTITY_REFERENCE         = 5;
  NODE_ENTITY                   = 6;
  NODE_PROCESSING_INSTRUCTION   = 7;
  NODE_COMMENT                  = 8;
  NODE_DOCUMENT                 = 9;
  NODE_DOCUMENT_TYPE            = 10;
  NODE_DOCUMENT_FRAGMENT        = 11;
  NODE_NOTATION                 = 12;

{ standalone simple encrypt/decrypt functions }

const
  CRYPT_C1 = 52845;
  CRYPT_C2 = 22719;
  CRYPT_KEY = 9570;

type
  TBytes = array of Byte;

{**
 *  TAKE CARE
 *  =========
 *  The MSXML DOM ProgID change from 4.0 for X86 to 6.0 for X64 bits.
 *
 **}
const
{$IFDEF CPUX86}
  MSXML_DOM_PROG_ID = 'Msxml2.DOMDocument.4.0';
{$ENDIF}
{$IFDEF CPUX64}
  MSXML_DOM_PROG_ID = 'Msxml2.DOMDocument.6.0';
{$ENDIF}

{**
 *  TAKE CARE
 *  =========
 *  For compatibility with BDS2006 version in which native data are AnsiString we have to force the Encrypt to translate
 *  source data S from string (Unicode) to X (AnsiString) because also Encrypt uses this way to read old BDS2006 created
 *  files.
 *
 *  This require that XML engine, which uses Unicode NEVER use pure Unicode chars but always convert them with the
 *  methods EncodeStringToXMLString and DecodeXMLStringToString. This functions conversts Unicode chars in XML plain
 *  ANSI escape strings.
 *
 **}
function Encrypt(const S: string; Key: Word): string;
var
  I: Integer;
  A: AnsiString;
  X: AnsiString;
  Size: Integer;
begin
  X := AnsiString(S);
  Size := Length(X);
  SetLength(A, Size);
  for I := 1 to Size do
  begin
    A[I] := AnsiChar(Byte(X[I]) xor (Key shr 8));
    Key := (Byte(A[I]) + Key) * CRYPT_C1 + CRYPT_C2;
  end;
  Result := string(A);
end;

function Decrypt(const B: TBytes; Key: Word): string;
var
  I: Integer;
  A: AnsiString;
  Size: Integer;
begin
  Size := Length(B);
  SetLength(A, Size);
  for I := 1 to Size do
  begin
    A[I] := AnsiChar(Byte(B[I - 1]) xor (Key shr 8));
    Key := (Byte(B[I - 1]) + Key) * CRYPT_C1 + CRYPT_C2;
  end;
  Result := string(A);
end;

function IsFileCrypted(const Filename: string): Boolean;
const
  XML_HEADER = '<?xml';
var
  A: AnsiString;
  HeaderSize: Integer;
  Stream: TMemoryStream;
begin
  try
    Stream := TMemoryStream.Create;
    try
      Stream.LoadFromFile(FileName);
      HeaderSize := Length(XML_HEADER);
      if Stream.Size < HeaderSize then AbortFast;
      SetLength(A, HeaderSize);
      if Stream.Read(A[1], HeaderSize) <> HeaderSize then AbortFast;
      Result := A <> XML_HEADER;
    finally
      Stream.Free;
    end;
  except
    Result := False;
  end;
end;

{ generic functions from osSysUtils }

function BytesFromFile(const FileName: string): TBytes;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    Stream.LoadFromFile(FileName);
    SetLength(Result, Stream.Size);
    Stream.Read(Pointer(Result)^, Stream.Size);
  finally
    Stream.Free;
  end;
end;

function StringFromFile(const FileName: string): string;
var
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    Stream.LoadFromFile(FileName);
    SetString(Result, PChar(Stream.Memory), Stream.Size);
  finally
    Stream.Free;
  end;
end;

procedure StringToFile(const S: string; const FileName: string);
var
  A: AnsiString;
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    A := AnsiString(S);
    Stream.Write(A[1], Length(A));
    Stream.SaveToFile(FileName);
  finally
    Stream.Free;
  end;
end;

{ TXMLMemento }

function CreateReadRoot(const FileName: string): IMemento;
var
  S: string;
  B: TBytes;
  Document: IXMLDOMDocument;
begin
  try
    if not FileExists(FileName) then Exit;
    if not IsFileCrypted(FileName) then
    begin
      Document := CreateOleObject(MSXML_DOM_PROG_ID);
      Document.async := False;
      Document.validateOnParse := True;
      Document.resolveExternals := True;
      Document.preserveWhiteSpace := True;
      Document.load(FileName);
      if Document.parseError.errorCode <> 0 then AbortFast;
      Result := TXMLMemento.Create(Document, Document.documentElement);
    end
    else
    begin
      B := BytesFromFile(FileName);
      S := Decrypt(B, CRYPT_KEY);
      Result := TXMLMemento.CreateFromString(S);
    end;
  except
    Result := nil;
  end;
end;

function CreateReadRootSmart(const FileName, RootName: string): IMemento;
begin
  try
    Result := CreateReadRoot(FileName);
    if (Result = nil) or (Result.GetName <> RootName) then
      Result := CreateWriteRoot(RootName);
  except
    Result := nil;
  end;
end;

function CreateWriteRoot(const Type_: string): IMemento;
var
  Node: IXMLDOMNode;
  Element: IXMLDOMElement;
  Document: IXMLDOMDocument;
begin
  try
    Document := CreateOleObject(MSXML_DOM_PROG_ID);
    Document.async := False;
    Document.validateOnParse := False;
    Document.resolveExternals := False;
    Document.preserveWhiteSpace := False;
    Node := Document.createProcessingInstruction('xml', 'version="1.0" encoding="utf-8"');
    Document.appendChild(Node);
    Node := Document.createElement(Type_);
    Document.appendChild(Node);
    Element := Document.documentElement;
    Result := TXMLMemento.Create(Document, Element);
  except
    Result := nil;
  end;
end;

function EncodeStringToXMLString(const S: TXMLString): TXMLString;
var
  I: Integer;
  C: WideChar;
begin
  Result := '';
  for I := 1 to Length(S) do
  begin
    C := S[I];
    case C of
      Char($20)..Char($7F): Result := Result + C;
    else
      Result := Result + '&#x' + IntToHex(Integer(C), 4) + ';';
    end;
  end;
end;

function DecodeXMLStringToString(const S: TXMLString): TXMLString;
var
  I: Integer;
  L: Integer;
  C: WideChar;
  NCR: string;
  Value: Integer;
  NCRTail: Integer;
begin
  if Pos('&#', S) = 0 then
    Result := S
  else
  begin
    I := 1;
    Result := '';
    L := Length(S);
    while I <= L do
    begin
      C := S[I];
      if C = '&' then
      begin
        NCR := Copy(S, I, L);
        NCRTail := Pos(';', NCR);
        if NCRTail <> 0 then
        begin
          if Length(NCR) >= 3 then
          begin
            if NCR[2] = '#' then
            begin
              if (NCR[3] <> 'x') and (NCR[3] <> 'X') then
              begin
                NCR := Copy(NCR, 3, NCRTail - 3);
                if TryStrToInt(NCR, Value) then
                begin
                  Result := Result + WideChar(Value);
                  I := I + NCRTail;
                  Continue;
                end;
              end
              else
              begin
                NCR := Copy(NCR, 4, NCRTail - 4);
                if TryStrToInt('$' + NCR, Value) then
                begin
                  Result := Result + WideChar(Value);
                  I := I + NCRTail;
                  Continue;
                end;
              end;
            end;
          end;
        end;
      end;
      Result := Result + C;
      Inc(I);
    end;
  end;
end;

{ IMemento }

function TXMLMemento.CreateChild(const Type_: string): IMemento;
var
  Node: IXMLDOMNode;
  Element: IXMLDOMElement;
begin
  try
    Node := FDocument.createNode(NODE_ELEMENT, Type_, '');
    Element := FElement.appendChild(Node);
    Result := TXMLMemento.Create(FDocument, Element);
  except
    Result := nil;
  end;
end;

function TXMLMemento.CreateChildSmart(const Type_: string): IMemento;
begin
  Result := GetChild(Type_);
  if Result = nil then
    Result := CreateChild(Type_);
end;

function TXMLMemento.GetBoolean(const Key: string; var Value: Boolean): Boolean;
var
  S: TXMLString;
begin
  Result := GetString(Key, S);
  if not Result then Exit;
  try
    if SameText(S, TRUE_FALSE[False]) then
      Value := False
    else if SameText(S, TRUE_FALSE[True]) then
      Value := True
    else
      Result := False;
  except
    Result := False;
  end;
end;

function TXMLMemento.GetCardinal(const Key: string; var Value: Cardinal): Boolean;
var
  S: TXMLString;
begin
  Result := GetString(Key, S);
  if not Result then Exit;
  try
    Value := StrToCardinal(S);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetChild(const Type_: string): IMemento;
var
  I: Integer;
begin
  Result := nil;
  try
    if FElement.childNodes.length = 0 then Exit;
    for I := 0 to FElement.childNodes.length - 1 do
    begin
      if FElement.childNodes.item(I).nodeName = Type_ then
      begin
        Result := TXMLMemento.Create(FDocument, FElement.childNodes.item(I));
        Exit;
      end;
    end;
  except
  end;
end;

function TXMLMemento.GetChildren(const Type_: string): IMementoArray;
var
  I: Integer;
begin
  Result := nil;
  try
    if FElement.childNodes.length = 0 then Exit;
    for I := 0 to FElement.childNodes.length - 1 do
    begin
      if (Type_ = '') or (FElement.childNodes.item(I).nodeName = Type_) then
      begin
        SetLength(Result, Length(Result) + 1);
        Result[Length(Result) - 1] := TXMLMemento.Create(FDocument, FElement.childNodes.item(I));
      end;
    end;
  except
    SetLength(Result, 0);
  end;
end;

function TXMLMemento.GetChildFromPath(const Path: string): IMemento;
var
  P: Pointer;
  Node: IXMLDOMNode;
begin
  Result := nil;
  try
    Node := FElement.selectSingleNode(Path);
    P := TVarData(Node).VDispatch;
    if P <> nil then
      Result := TXMLMemento.Create(FDocument, Node);
  except
  end;
end;

function TXMLMemento.GetChildrenFromPath(const Path: string): IMementoArray;
var
  I: Integer;
  NodeList: IXMLDOMNodeList;
begin
  Result := nil;
  try
    NodeList := FElement.selectNodes(Path);
    if NodeList.length = 0 then Exit;
    SetLength(Result, Integer(NodeList.length));
    for I := 0 to NodeList.length - 1 do
      Result[I] := TXMLMemento.Create(FDocument, NodeList.item(I));
  except
    SetLength(Result, 0);
  end;
end;

function TXMLMemento.GetDateTime(const Key: string; var Value: TDateTime): Boolean;
var
  D: Double;
begin
  Result := GetDouble(Key, D);
  if not Result then Exit;
  Value := D;
end;

function TXMLMemento.GetDouble(const Key: string; var Value: Double): Boolean;
var
  S: TXMLString;
begin
  Result := GetString(Key, S);
  if not Result then Exit;
  try
    if UpperCase(S) = 'NAN' then
      Value := Math.NaN
    else
      Value := StrToFloat(S);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetGUID(const Key: string; var Value: TGUID): Boolean;
var
  S: TXMLString;
begin
  Result := GetString(Key, S);
  if not Result then Exit;
  try
    Value := StringToGUID(S);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetInteger(const Key: string; var Value: Integer): Boolean;
var
  S: TXMLString;
begin
  Result := GetString(Key, S);
  if not Result then Exit;
  try
    Value := StrToInt(S);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetName: string;
begin
  try
    Result := FElement.nodeName;
  except
    Result := '';
  end;
end;

function TXMLMemento.GetRoot: IMemento;
begin
  if VarIsEmpty(FDocument) or VarIsEmpty(FElement) then
  begin
    Result := nil;
    Exit;
  end;
  try
    Result := TXMLMemento.Create(FDocument, FElement);
  except
    Result := nil;
  end;
end;

function TXMLMemento.GetParent: IMemento;
begin
  try
    Result := TXMLMemento.Create(FDocument, FElement.parentNode);
  except
    Result := nil;
  end;
end;

function TXMLMemento.GetString(const Key: string; var Value: TXMLString): Boolean;
var
  Attribute: IXMLDOMAttribute;
begin
  try
    Attribute := FElement.getAttributeNode(Key);
    Result := not (VarIsClear(Attribute) or VarIsEmpty(Attribute));
    if Result then Value := DecodeXMLStringToString(Attribute.Value);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetTextData(var Value: TXMLString): Boolean;
var
  Text: IXMLDOMText;
begin
  try
    Text := GetElementNodeText;
    Result := not VarIsNull(Text);
    if not Result then Exit;
    Value := Text.nodeValue;
  except
    Result := False;
  end;
end;

procedure TXMLMemento.PutBoolean(const Key: string; Value: Boolean);
begin
  PutString(Key, TRUE_FALSE[Value]);
end;

procedure TXMLMemento.PutCardinal(const Key: string; Value: Cardinal);
begin
  PutString(Key, IntToStr(Int64(Value)));
end;

procedure TXMLMemento.PutDateTime(const Key: string; Value: TDateTime);
begin
  PutDouble(Key, Value);
end;

procedure TXMLMemento.PutDouble(const Key: string; Value: Double);
begin
  PutString(Key, FloatToStr(Value));
end;

procedure TXMLMemento.PutGUID(const Key: string; const Value: TGUID);
begin
  PutString(Key, GUIDToString(Value));
end;

procedure TXMLMemento.PutInteger(const Key: string; Value: Integer);
begin
  PutString(Key, IntToStr(Value));
end;

procedure TXMLMemento.PutString(const Key: string; const Value: TXMLString);
var
  Attribute: IXMLDOMAttribute;
begin
  Attribute := FDocument.createAttribute(Key);
  Attribute.Value := EncodeStringToXMLString(Value);
  FElement.setAttributeNode(Attribute);
end;

// TODO: Forse bisogna convertire in EncodeStringToXMLString anche Data prima di buttarla nel node TEXT e forse bisogna
//       fare un parsing ancora più spinto per eleminare eventuali < o > nel testo stesso...

procedure TXMLMemento.PutTextData(const Data: TXMLString);
var
  Text: IXMLDOMText;
begin
  Text := GetElementNodeText;
  if not VarIsNull(Text) then
    Text.nodeValue := Data
  else
  begin
    Text := FDocument.createTextNode(Data);
    FElement.appendChild(Text);
  end;
end;

{ IPersistable }

function TXMLMemento.LoadFromFile(const FileName: string): Boolean;
begin
  try
    FDocument.load(FileName);
    // TODO: in realtà qui dovrei gestire un codice errore da rendere visibile
    if FDocument.parseError.errorCode <> 0 then AbortFast;
    FElement := FDocument.documentElement;
    Result := True;
  except
    FElement := UnAssigned;
    Result := False;
  end;
end;

function TXMLMemento.LoadFromStream(Stream: TStream): Boolean;
var
  C: AnsiChar;
  XML: AnsiString;
begin
  try
    // TODO: qui ho dato per assodato che il contenuto NON sia in (Unicode)string. Bisogna gestire la cosa
    C := ' ';
    XML := StringOfChar(C, Stream.Size);  // force to use AnsiChar version
    Stream.Position := 0;
    Stream.ReadBuffer(XML[1], Stream.Size);
    FDocument.loadXML(XML);
    // TODO: in realtà qui dovrei gestire un codice errore da rendere visibile
    if FDocument.parseError.errorCode <> 0 then AbortFast;
    FElement := FDocument.documentElement;
    Result := True;
  except
    FElement := UnAssigned;
    Result := False;
  end;
end;

function TXMLMemento.LoadFromString(const S: string): Boolean;
begin
  try
    FDocument.loadXML(S);
    // TODO: in realtà qui dovrei gestire un codice errore da rendere visibile
    if FDocument.parseError.errorCode <> 0 then AbortFast;
    FElement := FDocument.documentElement;
    Result := True;
  except
    FElement := UnAssigned;
    Result := False;
  end;
end;

function TXMLMemento.SaveToFile(const FileName: string; NormalizeMode: TNormalizeMode; Crypted: Boolean): Boolean;
var
  S: string;
  NormalizedDocument: IXMLDOMDocument;
begin
  try
    if not Crypted then
    begin
      NormalizedDocument := GetNormalizedDocument(NormalizeMode);
      NormalizedDocument.save(FileName);
    end
    else
    begin
      SaveToString(S, NormalizeMode);
      S := Encrypt(S, CRYPT_KEY);
      StringToFile(S, FileName);
    end;
    Result := True;
  except
    Result := False;
  end;
end;

function TXMLMemento.SaveToStream(Stream: TStream; NormalizeMode: TNormalizeMode): Boolean;
var
  S: string;
  NormalizedDocument: IXMLDOMDocument;
begin
  try
    // TODO: qui ho dato per assodato che il contenuto NON sia in (Unicode)String. Bisogna gestire la cosa
    NormalizedDocument := GetNormalizedDocument(NormalizeMode);
    S := NormalizedDocument.xml;
    Stream.Write(S[1], Length(S));
    Result := True;
  except
    Result := False;
  end;
end;

function TXMLMemento.SaveToString(var S: string; NormalizeMode: TNormalizeMode): Boolean;
var
  Stream: TMemoryStream;
begin
  try
    S := '';
    Stream := TMemoryStream.Create;
    try
      SaveToStream(Stream, NormalizeMode);
      SetLength(S, Stream.Size);
      Stream.Position := 0;
      Stream.Read(S[1], Stream.Size);
    finally
      Stream.Free;
    end;
    Result := True;
  except
    Result := False;
  end;
end;

{ EIS }

function TXMLMemento.GetElementNodeText: IXMLDOMText;
var
  I: Integer;
begin
  Result := Null;
  try
    if not FElement.hasChildNodes then Exit;
    for I := 0 to FElement.childNodes.length - 1 do
    begin
      if FElement.childNodes.item(I).nodeType = NODE_TEXT then
      begin
        Result := FElement.childNodes.item(I);
        Break;
      end;
    end;
  except
  end;
end;

function TXMLMemento.GetNormalizedDocument(NormalizeMode: TNormalizeMode): IXMLDOMDocument;
var
  X: string;
  W: string;
  DOM: IXMLDOMDocument;
  XSL: IXMLDOMDocument;
begin
  DOM := CreateOleObject(MSXML_DOM_PROG_ID);
  DOM.async := False;
  DOM.validateOnParse := False;
  DOM.resolveExternals := False;
  DOM.preserveWhiteSpace := False;
  DOM.documentElement := FDocument.documentElement.cloneNode(True);
  if NormalizeMode <> nrmd_None then
  begin
    if NormalizeMode = nrmd_UTF8 then
      X := 'utf-8'
    else
      X := 'utf-16';
    try
      XSL := CreateOleObject(MSXML_DOM_PROG_ID);
      XSL.async := False;
      XSL.validateOnParse := True;
      XSL.resolveExternals := True;
      XSL.preserveWhiteSpace := True;
      XSL.loadXML
      (
        '<?xml version="1.0" encoding="utf-8"?>' + #13#10 +
        '<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">' + #13#10 +
        Format('<xsl:output method="xml" indent="yes" encoding="%s"/>', [X]) + #13#10 +
        '<xsl:template match="@* | text() | node()">' + #13#10 +
        '  <xsl:copy>' + #13#10 +
        '    <xsl:apply-templates select="@* | text() | node()"/>' + #13#10 +
        '  </xsl:copy>' + #13#10 +
        '</xsl:template>' + #13#10 +
        '</xsl:stylesheet>'
      );
      DOM.transformNodeToObject(XSL, DOM);
    except
      W := Format('<?xml version="1.0" encoding="%s"?>', [X]) + DOM.xml;
      DOM.loadXML(W);
    end;
  end;
  Result := DOM;
end;

constructor TXMLMemento.Create(Document: IXMLDOMDocument; Element: IXMLDOMElement);
begin
  inherited Create;
  FDocument := Document;
  FElement := Element;
end;

class function TXMLMemento.CreateFromString(const S: string): IMemento;
var
  Document: IXMLDOMDocument;
begin
  try
    Document := CreateOleObject(MSXML_DOM_PROG_ID);
    Document.async := False;
    Document.validateOnParse := True;
    Document.resolveExternals := True;
    Document.preserveWhiteSpace := True;
    Document.loadXML(S);
    if Document.parseError.errorCode <> S_OK then AbortFast;
    Result := TXMLMemento.Create(Document, Document.documentElement);
  except
    Result := nil;
  end;
end;

end.
