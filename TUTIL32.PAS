unit TUtil32;

interface

uses Bde, DbiTypes;

{ TUtility (TUVerifyTable) Session Options }
const
  TU_Append_Errors = 1;
  TU_No_Secondary  = 2;
  TU_No_Warnings   = 4;
  TU_Header_Only   = 8;
  TU_Dialog_Hide   = 16;
  TU_No_Lock       = 32;

{ TUtility type definitions }
type
  hTUses  = Word;
  phTUses = ^hTUses;
  { Verify Callback processes }
  TUVerifyProcess = (TUVerifyHeader, TUVerifyIndex, TUVerifyData, TUVerifySXHeader,
                     TUVerifySXIndex, TUVerifySXData, TUVerifySXIntegrity,
                     TUVerifyTableName, TURebuild);
  { Call back info for Verify Callback function }
  TUVerifyCallBack = record
    PercentDone: word;
    TableName: DBIPath;
    Process: TUVerifyProcess;
    CurrentIndex: word;
    TotalIndex: word;
  end;

{ TUtility functions }
function TUInit(var hTUSession: hTUses): DBIResult;  stdcall;

function TUVerifyTable(hTUSession: hTUses;
                       pszTableName,
                       pszDriverType,
                       pszErrTableName,
                       pszPassword: PChar;
                       iOptions: integer;
                       var piErrorLevel: Integer): DBIResult; stdcall;

function TURebuildTable(hTUSession: hTUses;
                        pszTableName,
                        pszDriverType,
                        pszBackupTableName,
                        pszKeyviolName,
                        pszProblemTableName: PChar;
                        pCrDesc: pCRTblDesc): DBIResult; stdcall;

function TUGetCRTblDescCount(hTUSession: hTUses;
                             pszTableName: PChar;
                             var iFldCount,
                             iIdxCount,
                             iSecRecCount,
                             iValChkCount,
                             iRintCount,
                             iOptParams,
                             iOptDataLen: word): DBIResult; stdcall;

function TUFillCRTblDesc(hTUSession: hTUses;
                         pCrDesc: pCRTblDesc;
                         pszTableName,
                         pszPassword: PChar): DBIResult; stdcall;

function TUFillCURProps(hTUSession: hTUses;
                        pszTableName: PChar;
                        var tblProps: CURProps): DBIResult; stdcall;

function TUGetExtTblProps(hTUSession: hTUses;
                          pszTableName: PChar;
                          var pTS: TimeStamp;
                          var pbReadOnly: Boolean): DBIResult; stdcall;

function TUExit(hTUSession: hTUses): DBIResult; stdcall;

function TUGetErrorString(iErrorcode: DBIResult;
                          pszError: PChar): DBIResult; stdcall;

implementation

const
  TU32 = 'TUTIL32.DLL';

function TUInit;              external TU32 name 'TUInit';
function TUVerifyTable;       external TU32 name 'TUVerifyTable';
function TURebuildTable;      external TU32 name 'TURebuildTable';
function TUGetCRTblDescCount; external TU32 name 'TUGetCRTblDescCount';
function TUFillCRTblDesc;     external TU32 name 'TUFillCRTblDesc';
function TUFillCURProps;      external TU32 name 'TUFillCURProps';
function TUGetExtTblProps;    external TU32 name 'TUGetExtTblProps';
function TUExit;              external TU32 name 'TUExit';
function TUGetErrorString;    external TU32 name 'TUGetErrorString';

end.
