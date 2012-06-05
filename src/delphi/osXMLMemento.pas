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
    procedure PutGUID(const Key: string; Value: TGUID);
    procedure PutDouble(const Key: string; Value: Double);
    procedure PutInteger(const Key: string; Value: Integer);
    procedure PutString(const Key: string; const Value: WideString);
    procedure PutTextData(const Data: WideString);

    { IPersistable }
    procedure LoadFromFile(const FileName: string);
    procedure LoadFromStream(Stream: TStream);
    procedure LoadFromString(const S: string);
    procedure SaveToFile(const FileName: string; NormalizeMode: TNormalizeMode = nrmd_UTF16);
    procedure SaveToStream(Stream: TStream; NormalizeMode: TNormalizeMode = nrmd_UTF16);
    procedure SaveToString(var S: string; NormalizeMode: TNormalizeMode = nrmd_UTF16);

    { TXMLMemento }
    function GetElementNodeText: IXMLDOMText;
    function GetNormalizedDocument(NormalizeMode: TNormalizeMode): IXMLDOMDocument;
    constructor Create(Document: IXMLDOMDocument; Element: IXMLDOMElement);
  end;

function CreateReadRoot(const FileName: string): IMemento;
function CreateReadRootSmart(const FileName, RootName: string): IMemento;
function CreateWriteRoot(const Type_: string): IMemento;
function WideStringToXMLString(const S: WideString): WideString;
function XMLStringToWideString(const S: WideString): WideString;

implementation

uses
  ComObj,
  Windows,
  SysUtils,
  Variants;

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

function CreateReadRoot(const FileName: string): IMemento;
var
  Document: IXMLDOMDocument;
begin
  try
    Document := CreateOleObject('Msxml2.DOMDocument.4.0');
    Document.async := False;
    Document.validateOnParse := True;
    Document.resolveExternals := True;
    Document.preserveWhiteSpace := True;
    Document.load(FileName);
    if Document.parseError.errorCode <> 0 then Abort;
    Result := TXMLMemento.Create(Document, Document.documentElement);
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
    Document := CreateOleObject('Msxml2.DOMDocument.4.0');
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

function WideStringToXMLString(const S: WideString): WideString;
var
  I: Integer;
  W: WideChar;
begin
  Result := '';
  for I := 1 to Length(S) do
  begin
    W := S[I];
    case W of
      WideChar($20)..WideChar($7F): Result := Result + W;
    else
      Result := Result + '&#x' + IntToHex(Integer(W), 4) + ';';
    end;
  end;
end;

function XMLStringToWideString(const S: WideString): WideString;
var
  I: Integer;
  L: Integer;
  WC: WideChar;
  Value: Integer;
  NCR: WideString;
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
      WC := S[I];
      if WC = '&' then
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
      Result := Result + WC;
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
  W: WideString;
begin
  Result := GetString(Key, W);
  if not Result then Exit;
  try
    if WideSameText(W, TRUE_FALSE[False]) then
      Value := False
    else if WideSameText(W, TRUE_FALSE[True]) then
      Value := True
    else
      Result := False;
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
      if FElement.childNodes.item(I).nodeName = Type_ then
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

function TXMLMemento.GetDouble(const Key: string; var Value: Double): Boolean;
var
  W: WideString;
begin
  Result := GetString(Key, W);
  if not Result then Exit;
  try
    Value := StrToFloat(W);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetGUID(const Key: string; var Value: TGUID): Boolean;
var
  W: WideString;
begin
  Result := GetString(Key, W);
  if not Result then Exit;
  try
    Value := StringToGUID(W);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetInteger(const Key: string; var Value: Integer): Boolean;
var
  W: WideString;
begin
  Result := GetString(Key, W);
  if not Result then Exit;
  try
    Value := StrToInt(W);
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

function TXMLMemento.GetString(const Key: string; var Value: WideString): Boolean;
var
  Attribute: IXMLDOMAttribute;
begin
  try
    Attribute := FElement.getAttributeNode(Key);
    Result := not (VarIsClear(Attribute) or VarIsEmpty(Attribute));
    if Result then Value := XMLStringToWideString(Attribute.Value);
  except
    Result := False;
  end;
end;

function TXMLMemento.GetTextData(var Value: WideString): Boolean;
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

procedure TXMLMemento.PutDouble(const Key: string; Value: Double);
begin
  PutString(Key, FloatToStr(Value));
end;

procedure TXMLMemento.PutGUID(const Key: string; Value: TGUID);
begin
  PutString(Key, GUIDToString(Value));
end;

procedure TXMLMemento.PutInteger(const Key: string; Value: Integer);
begin
  PutString(Key, IntToStr(Value));
end;

procedure TXMLMemento.PutString(const Key: string; const Value: WideString);
var
  Attribute: IXMLDOMAttribute;
begin
  Attribute := FDocument.createAttribute(Key);
  Attribute.Value := WideStringToXMLString(Value);
  FElement.setAttributeNode(Attribute);
end;

// todo: Forse bisogna convertire in WideStringToXMLString anche Data prima di buttarla nel node TEXT e forse bisogna
//       fare un parsing ancora più spinto per eleminare eventuali < o > nel testo stesso... 

procedure TXMLMemento.PutTextData(const Data: WideString);
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

procedure TXMLMemento.LoadFromFile(const FileName: string);
begin
  FDocument.load(FileName);
  // todo: in realtà qui dovrei gestire un codice errore da rendere visibile
  if FDocument.parseError.errorCode = 0 then
    FElement := FDocument.documentElement
  else
    FElement := UnAssigned;
end;

procedure TXMLMemento.LoadFromStream(Stream: TStream);
var
  XML: string;
begin
  // todo: qui ho dato per assodato che il contenuto NON sia in WideString. Bisogna gestire la cosa
  XML := StringOfChar(' ', Stream.Size);
  Stream.Position := 0;
  Stream.ReadBuffer(XML[1], Stream.Size);
  FDocument.loadXML(XML);
  // todo: in realtà qui dovrei gestire un codice errore da rendere visibile
  if FDocument.parseError.errorCode = 0 then
    FElement := FDocument.documentElement
  else
    FElement := UnAssigned;
end;

procedure TXMLMemento.LoadFromString(const S: string);
begin
  FDocument.loadXML(S);
  // todo: in realtà qui dovrei gestire un codice errore da rendere visibile
  if FDocument.parseError.errorCode = 0 then
    FElement := FDocument.documentElement
  else
    FElement := UnAssigned;
end;

procedure TXMLMemento.SaveToFile(const FileName: string; NormalizeMode: TNormalizeMode);
var
  NormalizedDocument: IXMLDOMDocument;
begin
  NormalizedDocument := GetNormalizedDocument(NormalizeMode);
  NormalizedDocument.save(FileName);
end;

procedure TXMLMemento.SaveToStream(Stream: TStream; NormalizeMode: TNormalizeMode);
var
  S: string;
  NormalizedDocument: IXMLDOMDocument;
begin
  // todo: qui ho dato per assodato che il contenuto NON sia in WideString. Bisogna gestire la cosa
  NormalizedDocument := GetNormalizedDocument(NormalizeMode);
  S := NormalizedDocument.xml;
  Stream.Write(S[1], Length(S));
end;

procedure TXMLMemento.SaveToString(var S: string; NormalizeMode: TNormalizeMode);
var
  Stream: TMemoryStream;
begin
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
  X: WideString;
  W: WideString;
  DOM: IXMLDOMDocument;
  XSL: IXMLDOMDocument;
begin
  DOM := CreateOleObject('Msxml2.DOMDocument.4.0');
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
      XSL := CreateOleObject('Msxml2.DOMDocument.4.0');
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
      W := WideString(Format('<?xml version="1.0" encoding="%s"?>', [X])) + DOM.xml;
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

end.
