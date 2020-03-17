unit vgr_AXScript;

{$I vtk.inc}

interface

uses
  {$IFDEF VTK_NOLIC}vgr_ReportAdd, {$ENDIF} 
  Windows, ActiveX;


type
  tagSCRIPTSTATE = TOleEnum;
  tagDOCUMENTNAMETYPE = TOleEnum;

  TCatId = TGUID;

type
  tagBREAKRESUME_ACTION = TOleEnum;
const
  BREAKRESUMEACTION_ABORT = $00000000;
  BREAKRESUMEACTION_CONTINUE = $00000001;
  BREAKRESUMEACTION_STEP_INTO = $00000002;
  BREAKRESUMEACTION_STEP_OVER = $00000003;
  BREAKRESUMEACTION_STEP_OUT = $00000004;

type
  tagERRORRESUMEACTION = TOleEnum;
const
  ERRORRESUMEACTION_ReexecuteErrorStatement = $00000000;
  ERRORRESUMEACTION_AbortCallAndReturnErrorToCaller = $00000001;
  ERRORRESUMEACTION_SkipErrorStatement = $00000002;

type
  tagBREAKREASON = TOleEnum;
const
  BREAKREASON_STEP = $00000000;
  BREAKREASON_BREAKPOINT = $00000001;
  BREAKREASON_DEBUGGER_BLOCK = $00000002;
  BREAKREASON_HOST_INITIATED = $00000003;
  BREAKREASON_LANGUAGE_INITIATED = $00000004;
  BREAKREASON_DEBUGGER_HALT = $00000005;
  BREAKREASON_ERROR = $00000006;

type
  tagBREAKPOINT_STATE = TOleEnum;
const
  BREAKPOINT_DELETED = $00000000;
  BREAKPOINT_DISABLED = $00000001;
  BREAKPOINT_ENABLED = $00000002;

const
  CLSID_VBScript : TGUID = '{B54F3741-5B07-11cf-A4B0-00AA004A55E8}';
  CLSID_JScript : TGUID = '{f414c260-6ac0-11cf-b6d1-00aa00bbbb58}';

  CLSID_StdComponentCategoriesMgr : TGUID = '{0002E005-0000-0000-C000-000000000046}';
  IID_ICatInformation : TGUID = '{0002E013-0000-0000-C000-000000000046}';

  //Category IDs
  CATID_ActiveScript:TGUID=              '{F0B7A1A1-9847-11cf-8F20-00805F2CD064}';
  CATID_ActiveScriptParse:TGUID=         '{F0B7A1A2-9847-11cf-8F20-00805F2CD064}';

  //Interface IDs
  IID_IActiveScriptSite:TGUID=           '{DB01A1E3-A42B-11cf-8F20-00805F2CD064}';
  IID_IActiveScriptSiteWindow:TGUID=     '{D10F6761-83E9-11cf-8F20-00805F2CD064}';
  IID_IActiveScript:TGUID=               '{BB1A2AE1-A4F9-11cf-8F20-00805F2CD064}';
  IID_IActiveScriptParse:TGUID=          '{BB1A2AE2-A4F9-11cf-8F20-00805F2CD064}';
  IID_IActiveScriptParseProcedure:TGUID= '{1CFF0050-6FDD-11d0-9328-00A0C90DCAA9}';
  IID_IActiveScriptError:TGUID=          '{EAE1BA61-A4ED-11cf-8F20-00805F2CD064}';
  IID_IActiveScriptDebug:TGUID=          '{51973C10-CB0C-11d0-B5C9-00A0244A0E7A}';

  IID_IActiveScriptSiteDebug: TGUID =    '{51973C11-CB0C-11D0-B5C9-00A0244A0E7A}';



  SOURCETEXT_ATTR_KEYWORD	        = $00000001;
  SOURCETEXT_ATTR_COMMENT	        = $00000002;
  SOURCETEXT_ATTR_NONSOURCE	      = $00000004;
  SOURCETEXT_ATTR_OPERATOR	      = $00000008;
  SOURCETEXT_ATTR_NUMBER	        = $00000010;
  SOURCETEXT_ATTR_STRING	        = $00000020;
  SOURCETEXT_ATTR_FUNCTION_START	= $00000040;
  SOURCETEXT_ATTR_IDENTIFIER      = $00000100;


  // Constants used by ActiveX Scripting:
  SCRIPTITEM_ISVISIBLE     = $00000002;
  SCRIPTITEM_ISSOURCE      = $00000004;
  SCRIPTITEM_GLOBALMEMBERS = $00000008;
  SCRIPTITEM_ISPERSISTENT  = $00000040;
  SCRIPTITEM_CODEONLY      = $00000200;
  SCRIPTITEM_NOCODE        = $00000400;
  SCRIPTITEM_ALL_FLAGS     = (SCRIPTITEM_ISSOURCE or
                             SCRIPTITEM_ISVISIBLE or
                             SCRIPTITEM_ISPERSISTENT or
                             SCRIPTITEM_GLOBALMEMBERS or
                             SCRIPTITEM_NOCODE or
                             SCRIPTITEM_CODEONLY);

   // IActiveScript::AddTypeLib() input flags

   SCRIPTTYPELIB_ISCONTROL    = $00000010;
   SCRIPTTYPELIB_ISPERSISTENT = $00000040;
   SCRIPTTYPELIB_ALL_FLAGS    = (SCRIPTTYPELIB_ISCONTROL or
                                 SCRIPTTYPELIB_ISPERSISTENT);

// IActiveScriptParse::AddScriptlet() and IActiveScriptParse::ParseScriptText() input flags */

   SCRIPTTEXT_DELAYEXECUTION    = $00000001;
   SCRIPTTEXT_ISVISIBLE         = $00000002;
   SCRIPTTEXT_ISEXPRESSION      = $00000020;
   SCRIPTTEXT_ISPERSISTENT      = $00000040;
   SCRIPTTEXT_HOSTMANAGESSOURCE = $00000080;
   SCRIPTTEXT_ALL_FLAGS         = (SCRIPTTEXT_DELAYEXECUTION or
                                   SCRIPTTEXT_ISVISIBLE or
                                   SCRIPTTEXT_ISEXPRESSION or
                                   SCRIPTTEXT_HOSTMANAGESSOURCE or
                                   SCRIPTTEXT_ISPERSISTENT);


// IActiveScriptParseProcedure::ParseProcedureText() input flags

  SCRIPTPROC_HOSTMANAGESSOURCE  = $00000080;
  SCRIPTPROC_IMPLICIT_THIS      = $00000100;
  SCRIPTPROC_IMPLICIT_PARENTS   = $00000200;
  SCRIPTPROC_ALL_FLAGS          = (SCRIPTPROC_HOSTMANAGESSOURCE or
                                   SCRIPTPROC_IMPLICIT_THIS or
                                   SCRIPTPROC_IMPLICIT_PARENTS);


// IActiveScriptSite::GetItemInfo() input flags */

   SCRIPTINFO_IUNKNOWN  = $00000001;
   SCRIPTINFO_ITYPEINFO = $00000002;
   SCRIPTINFO_ALL_FLAGS = (SCRIPTINFO_IUNKNOWN or
                           SCRIPTINFO_ITYPEINFO);


// IActiveScript::Interrupt() Flags */

   SCRIPTINTERRUPT_DEBUG          = $00000001;
   SCRIPTINTERRUPT_RAISEEXCEPTION = $00000002;
   SCRIPTINTERRUPT_ALL_FLAGS      = (SCRIPTINTERRUPT_DEBUG or
                                     SCRIPTINTERRUPT_RAISEEXCEPTION);



type
  //new IE4 types
  TUserHWND=HWND;
  TUserBSTR=TBStr;
  TUserExcepInfo=TExcepInfo;
  TUserVariant=OleVariant;

  // script state values
  TScriptState = (
    SCRIPTSTATE_UNINITIALIZED,
    SCRIPTSTATE_STARTED,
    SCRIPTSTATE_CONNECTED,
    SCRIPTSTATE_DISCONNECTED,
    SCRIPTSTATE_CLOSED,
    SCRIPTSTATE_INITIALIZED
    );

  // script thread state values */
  TScriptThreadState = (
    SCRIPTTHREADSTATE_NOTINSCRIPT,
    SCRIPTTHREADSTATE_RUNNING
    );


  // Thread IDs */
  TScriptThreadID = DWORD;

  _GUID = packed record
    Data1: LongWord;
    Data2: Word;
    Data3: Word;
    Data4: array[0..7] of Byte;
  end;


const  //Note: these SCRIPTTHREADID constants were originally macros
       //in the first version of this file.  See the note at the top
       //for more information. (Thanks to Gary Warren King.)
  SCRIPTTHREADID_CURRENT        = TScriptThreadId(-1);
  SCRIPTTHREADID_BASE           = TScriptThreadId(-2);
  SCRIPTTHREADID_ALL            = TScriptThreadId(-3);

type
  //Forward declarations
  IActiveScript = interface;
  IActiveScriptParse = interface;
  IActiveScriptParseProcedure = interface;
  IActiveScriptSite = interface;
  IActiveScriptSiteWindow = interface;
  IActiveScriptError = interface;
  IDebugDocumentContext = interface;
  IDebugDocumentHelper = interface;
  IActiveScriptDebug = interface;
  IDebugApplication = interface;
  IDebugApplicationNode = interface;
  IActiveScriptErrorDebug = interface;
  IDebugDocumentInfo = interface;
  IDebugDocument = interface;
  IEnumDebugCodeContexts = interface;
  IDebugDocumentHost = interface;
  IDebugApplicationThread = interface;
  IRemoteDebugApplication = interface;
  IDebugDocumentProvider = interface;
  IDebugStackFrame = interface;
  IDebugCodeContext = interface;
  IRemoteDebugApplicationThread = interface;
  IApplicationDebugger = interface;
  IEnumRemoteDebugApplicationThreads = interface;
  IEnumDebugExpressionContexts = interface;
  IDebugSyncOperation = interface;
  IDebugAsyncOperation = interface;
  IDebugStackFrameSniffer = interface;
  IDebugThreadCall = interface;
  IProvideExpressionContexts = interface;
  IEnumDebugApplicationNodes = interface;
  IEnumDebugStackFrames = interface;
  IDebugExpressionContext = interface;
  IDebugAsyncOperationCallBack = interface;
  IDebugExpression = interface;
  IDebugExpressionCallBack = interface;
  IDebugProperty = interface;

  tagDebugStackFrameDescriptor = packed record
    pdsf: IDebugStackFrame;
    dwMin: LongWord;
    dwLim: LongWord;
    fFinal: Integer;
    punkFinal: IUnknown;
  end;

  IActiveScriptError = interface(IUnknown)
    ['{EAE1BA61-A4ED-11CF-8F20-00805F2CD064}']

    // HRESULT GetExceptionInfo(
    //     [out] EXCEPINFO *pexcepinfo);
    function GetExceptionInfo(out ExcepInfo: TExcepInfo): HRESULT; stdcall;

    // HRESULT GetSourcePosition(
    //     [out] DWORD *pdwSourceCOntext,
    //     [out] ULONG *pulLineNumber,
    //     [out] LONG *plCharacterPosition);
    function GetSourcePosition(out SourceContext: DWORD; out LineNumber: ULONG; out CharacterPosition: LONGINT): HRESULT; stdcall;

    // HRESULT GetSourceLineText(
    //     [out] BSTR *pbstrSourceLine);
    function GetSourceLineText(out SourceLine: LPWSTR): HRESULT; stdcall;
  end; //IActiveScriptError interface


  IActiveScriptSite = Interface(IUnknown)
    ['{DB01A1E3-A42B-11CF-8F20-00805F2CD064}']
    // HRESULT GetLCID(
    //     [out] LCID *plcid);
    // Allows the host application to indicate the local ID for localization
    // of script/user interaction
    function GetLCID(out Lcid: TLCID): HRESULT; stdcall;

    // HRESULT GetItemInfo(
    //     [in] LPCOLESTR pstrName,
    //     [in] DWORD dwReturnMask,
    //     [out] IUnknown **ppiunkItem,
    //     [out] ITypeInfo **ppti);
    // Called by the script engine to look up named items in host application.
    // Used to map unresolved variable names in scripts to automation interface
    // in host application.  The dwReturnMask parameter will indicate whether
    // the actual object (SCRIPTINFO_INKNOWN) or just a coclass type description
    // (SCRIPTINFO_ITYPEINFO)  is desired.
    function GetItemInfo(const pstrName: POleStr; dwReturnMask: DWORD; out ppiunkItem: IUnknown; out Info: ITypeInfo): HRESULT; stdcall;

    // HRESULT GetDocVersionString(
    //     [out] BSTR *pbstrVersion);
    // Called by the script engine to get a text-based version number of the
    // current document.  This string can be used to validate that any cached
    // state that the script engine may have saved is consistent with the
    // current document.
    function GetDocVersionString(out Version: TBSTR): HRESULT; stdcall;

    // HRESULT OnScriptTerminate(
    //     [in] const VARIANT *pvarResult,
    //     [in] const EXCEPINFO *pexcepinfo);
    // Called by the script engine when the script terminates.  In most cases
    // this method is not called, as it is possible that the parsed script may
    // be used to dispatch events from the host application
    function OnScriptTerminate(const pvarResult: OleVariant; const pexcepinfo: TExcepInfo): HRESULT; stdcall;

    // HRESULT OnStateChange(
    //     [in] SCRIPTSTATE ssScriptState);
    // Called by the script engine when state changes either explicitly via
    // SetScriptState or implicitly via other script engine events.
    function OnStateChange(ScriptState: TScriptState): HRESULT; stdcall;

    // HRESULT OnScriptError(
    //     [in] IActiveScriptError *pscripterror);
    // Called when script execution or parsing encounters an error.  The script
    // engine will provide an implementation of IActiveScriptError that
    // describes the runtime error in terms of an EXCEPINFO in addition to
    // indicating the location of the error in the original script text.
    function OnScriptError(const pscripterror: IActiveScriptError): HRESULT; stdcall;

    // HRESULT OnEnterScript(void);
    // Called by the script engine to indicate the beginning of a unit of work.
    function OnEnterScript: HRESULT; stdcall;

    // HRESULT OnLeaveScript(void);
    // Called by the script engine to indicate the completion of a unit of work.
    function OnLeaveScript: HRESULT; stdcall;

  end; //IActiveScriptSite interface


  IActiveScriptSiteWindow = interface(IUnknown)
   ['{D10F6761-83E9-11CF-8F20-00805F2CD064}']
    // HRESULT GetWindow(
    //     [out] HWND *phwnd);
    function GetWindow(out Handle: HWND): HRESULT; stdcall;

    // HRESULT EnableModeless(
    //     [in] BOOL fEnable);
    function EnableModeless(fEnable: BOOL): HRESULT; stdcall;
  end;  //IActiveScriptSiteWindow interface

  IActiveScript = interface(IUnknown)
    ['{BB1A2AE1-A4F9-11CF-8F20-00805F2CD064}']
    // HRESULT SetScriptSite(
    //     [in] IActiveScriptSite *pass);
    // Conects the host's application site object to the engine
    function SetScriptSite(ActiveScriptSite: IActiveScriptSite): HRESULT; stdcall;

    // HRESULT GetScriptSite(
    //     [in] REFIID riid,
    //     [iid_is][out] void **ppvObject);
    // Queries the engine for the connected site
    function GetScriptSite(riid: TGUID; out OleObject: Pointer): HRESULT; stdcall;

    // HRESULT SetScriptState(
    //     [in] SCRIPTSTATE ss);
    // Causes the engine to enter the designate state
    function SetScriptState(State: TScriptState): HRESULT; stdcall;

    // HRESULT GetScriptState(
    //     [out] SCRIPTSTATE *pssState);
    // Queries the engine for its current state
    function GetScriptState(out State: TScriptState): HRESULT; stdcall;

    // HRESULT Close(void);
    // Forces the engine to enter the closed state, resetting any parsed scripts
    // and disconnecting/releasing all of the host's objects.
    function Close: HRESULT; stdcall;

    // HRESULT AddNamedItem(
    //     [in] LPCOLESTR pstrName,
    //     [in] DWORD dwFlags);
    // Adds a variable name to the namespace of the script engine. The engine
    // will call the site's GetItemInfo to resolve the name to an object.
    function AddNamedItem(Name: POleStr; Flags: DWORD): HRESULT; stdcall;

    // HRESULT AddTypeLib(
    //     [in] REFGUID rguidTypeLib,
    //     [in] DWORD dwMajor,
    //     [in] DWORD dwMinor,
    //     [in] DWORD dwFlags);
    // Adds the type and constant defintions contained in the designated type
    // library to the namespace of the scripting engine.
    function AddTypeLib(TypeLib: TGUID; Major: DWORD; Minor: DWORD; Flags: DWORD): HRESULT; stdcall;

    // HRESULT GetScriptDispatch(
    //     [in] LPCOLESTR pstrItemName,
    //     [out] IDispatch **ppdisp);
    // Gets the IDispatch pointer to the scripting engine.
    function GetScriptDispatch(ItemName: POleStr; out Disp: IDispatch): HRESULT; stdcall;

    // HRESULT GetCurrentScriptThreadID(
    //     [out] SCRIPTTHREADID *pstidThread);
    // Gets the script's logical thread ID that corresponds to the current
    // physical thread.  This allows script engines to execute script code on
    // arbitrary threads that may be distinct from the site's thread.
    function GetCurrentScriptThreadID(out Thread: TScriptThreadID): HRESULT; stdcall;

    // HRESULT GetScriptThreadID(
    //     [in] DWORD dwWin32ThreadID,
    //     [out] SCRIPTTHREADID *pstidThread);
    // Gets the logical thread ID that corresponds to the specified physical
    // thread.  This allows script engines to execute script code on arbitrary
    // threads that may be distinct from the sites thread.
    function GetScriptThreadID(Win32ThreadID: DWORD; out Thread: TScriptThreadID): HRESULT; stdcall;

    // HRESULT GetScriptThreadState(
    //     [in] SCRIPTTHREADID stidThread,
    //     [out] SCRIPTTHREADSTATE *pstsState);
    // Gets the logical thread ID running state, which is either
    // SCRIPTTHREADSTATE_NOTINSCRIPT or SCRIPTTHEADSTATE_RUNNING.
    function GetScriptThreadState(Thread: TScriptThreadID; out State: TScriptThreadState): HRESULT; stdcall;

    // HRESULT InterruptScriptThread(
    //     [in] SCRIPTTHREADID stidThread,
    //     [in] const EXCEPINFO *pexcepInfo,
    //     [in] DWORD dwFlags);
    // Similar to Terminatethread, this method stope the execution of a script thread.
    function InterruptScriptThread(Thread: TScriptThreadID; const ExcepInfo: TExcepInfo; Flags: DWORD): HRESULT; stdcall;

    // HRESULT Clone(
    //     [out] IActiveScript **ppscript);
    // Duplicates the current script engine, replicating any parsed script text
    // and named items, but no the actual pointers to the site's objects.
    function Clone(out ActiveScript: IActiveScript): HRESULT; stdcall;
  end;  //IActiveScript interface

  IActiveScriptParse = interface(IUnknown)
    ['{BB1A2AE2-A4F9-11CF-8F20-00805F2CD064}']

    // HRESULT InitNew(void);
    function InitNew: HRESULT; stdcall;

    // HRESULT AddScriptlet(
    //     [in] LPCOLESTR pstrDefaultName,
    //     [in] LPCOLESTR pstrCode,
    //     [in] LPCOLESTR pstrItemName,
    //     [in] LPCOLESTR pstrSubItemName,
    //     [in] LPCOLESTR pstrEventName,
    //     [in] LPCOLESTR pstrDelimiter,
    //     [in] DWORD dwSourceContextCookie,
    //     [in] ULONG ulStartingLineNumber,
    //     [in] DWORD dwFlags,
    //     [out] BSTR *pbstrName,
    //     [out] EXCEPINFO *pexcepinfo);
    function AddScriptlet(
          DefaultName: POleStr;
          Code: POleStr;
          ItemName: POleStr;
          SubItemName: POleStr;
          EventName: POleStr;
          Delimiter: POleStr;
          SourceContextCookie: DWORD;
          StartingLineNnumber: ULONG;
          Flags: DWORD;
      out Name: TBSTR;
      out ExcepInfo: TExcepInfo
    ): HRESULT; stdcall;

    // HRESULT STDMETHODCALLTYPE ParseScriptText(
    //     [in] LPCOLESTR pstrCode,
    //     [in] LPCOLESTR pstrItemName,
    //     [in] IUnknown  *punkContext,
    //     [in] LPCOLESTR pstrDelimiter,
    //     [in] DWORD dwSourceContextCookie,
    //     [in] ULONG ulStartingLineNumber,
    //     [in] DWORD dwFlags,
    //     [out] VARIANT *pvarResult,
    //     [out] EXCEPINFO *pexcepinfo);
    function ParseScriptText(
      const pstrCode: POLESTR;
      const pstrItemName: POLESTR;
      const punkContext: IUnknown;
      const pstrDelimiter: POLESTR;
            dwSourceContextCookie: DWORD;
            ulStartingLineNumber: ULONG;
            dwFlags: DWORD;
      out   pvarResult: OleVariant;
      out   pExcepInfo: TExcepInfo
    ): HRESULT; stdcall;

end;  //IActivScriptParse interface


IActiveScriptParseProcedure=interface(IUnknown)
  ['{1CFF0050-6FDD-11d0-9328-00A0C90DCAA9}']

  function ParseProcedureText(
     const pstrCode: POLESTR;
     const pstrFormalParams: POLESTR;
     const pstrItemName: POLESTR;
           punkContext: IUnknown;
     const pstrDelimiter: POLESTR;
           dwSourceContextCookie: DWord;
           ulStartingLineNumber: ULong;
           dwFlags: DWord;
     out   ppdisp: IDispatch
  ): HResult; stdcall;

end;  //IActivScriptParseProcedure interface

IActiveScriptDebug = interface(IUnknown)
['{51973C10-CB0C-11d0-B5C9-00A0244A0E7A}']

  // Returns the text attributes for an arbitrary block of script text. Smart hosts
  // HRESULT GetScriptTextAttributes(
  // The script block text. This string need not be null terminated.
  // [in, size_is(uNumCodeChars)]LPCOLESTRpstrCode,
  // The number of characters in the script block text.
  // [in]ULONGuNumCodeChars,
  // See IActiveScriptParse::ParseScriptText for a description of this argument.
  // [in]LPCOLESTRpstrDelimiter,
  // See IActiveScriptParse::ParseScriptText for a description of this argument.
  // [in]DWORDdwFlags,
  // Buffer to contain the returned attributes.
  // [in, out, size_is(uNumCodeChars)]SOURCE_TEXT_ATTR *pattr);

  function GetScriptTextAttributes(
     const pstrCode: POLESTR;
     const uNumCodeChars : ULong;
     const pstrDelimiter: POLESTR;
           dwFlags: DWord;
     var pattr
  ): HResult; stdcall;

  end; // IActiveScriptDebug

// *********************************************************************//
// Interface: IActiveScriptSiteDebug
// Flags:     (0)
// GUID:      {51973C11-CB0C-11D0-B5C9-00A0244A0E7A}
// *********************************************************************//
  IActiveScriptSiteDebug = interface(IUnknown)
    ['{51973C11-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetDocumentContextFromPosition {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_19:0}dwSourceContext: LongWord;
                                                                                  {VT_19:0}uCharacterOffset: LongWord;
                                                                                  {VT_19:0}uNumChars: LongWord;
                                                                                  {VT_29:2}out ppsc: IDebugDocumentContext): HResult; stdcall;

    function  GetApplication {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppda: IDebugApplication): HResult; stdcall;
    function  GetRootApplicationNode {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdanRoot: IDebugApplicationNode): HResult; stdcall;
    function  OnScriptErrorDebug {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_29:1}const pErrorDebug: IActiveScriptErrorDebug;
                                                                      {VT_3:1}out pfEnterDebugger: Integer;
                                                                      {VT_3:1}out pfCallOnScriptErrorWhenContinuing: Integer): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IDebugDocumentContext
// Flags:     (0)
// GUID:      {51973C28-CB0C-11D0-B5C9-00A0244A0E7A}
// *********************************************************************//
  IDebugDocumentContext = interface(IUnknown)
    ['{51973C28-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetDocument {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppsd: IDebugDocument): HResult; stdcall;
    function  EnumCodeContexts {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppescc: IEnumDebugCodeContexts): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IDebugDocumentHelper
// Flags:     (0)
// GUID:      {51973C26-CB0C-11D0-B5C9-00A0244A0E7A}
// *********************************************************************//
  IDebugDocumentHelper = interface(IUnknown)
    ['{51973C26-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  Init {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_29:1}const pda: IDebugApplication;
                                                        {VT_31:0}pszShortName: PWideChar;
                                                        {VT_31:0}pszLongName: PWideChar;
                                                        {VT_19:0}docAttr: LongWord): HResult; stdcall;
    function  Attach {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pddhParent: IDebugDocumentHelper): HResult; stdcall;
    function  Detach {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  AddUnicodeText {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pszText: PWideChar): HResult; stdcall;
    function  AddDBCSText {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_30:0}pszText: PChar): HResult; stdcall;
    function  SetDebugDocumentHost {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pddh: IDebugDocumentHost): HResult; stdcall;
    function  AddDeferredText {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_19:0}cChars: LongWord;
                                                                   {VT_19:0}dwTextStartCookie: LongWord): HResult; stdcall;
    function  DefineScriptBlock {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_19:0}ulCharOffset: LongWord;
                                                                     {VT_19:0}cChars: LongWord;
                                                                     {VT_29:1}const pas: IActiveScript;
                                                                     {VT_3:0}fScriptlet: Integer;
                                                                     {VT_19:1}out pdwSourceContext: LongWord): HResult; stdcall;
    function  SetDefaultTextAttr {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_18:0}staTextAttr: Word): HResult; stdcall;
    function  SetTextAttributes {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}ulCharOffset: LongWord;
                                                                     {VT_19:0}cChars: LongWord;
                                                                     {VT_18:1}var pstaTextAttr: Word): HResult; stdcall;
    function  SetLongName {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pszLongName: PWideChar): HResult; stdcall;
    function  SetShortName {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pszShortName: PWideChar): HResult; stdcall;
    function  SetDocumentAttr {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}pszAttributes: LongWord): HResult; stdcall;
    function  GetDebugApplicationNode {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdan: IDebugApplicationNode): HResult; stdcall;
    function  GetScriptBlockInfo {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_19:0}dwSourceContext: LongWord;
                                                                      {VT_29:2}out ppasd: IActiveScript;
                                                                      {VT_19:1}out piCharPos: LongWord;
                                                                      {VT_19:1}out pcChars: LongWord): HResult; stdcall;
    function  CreateDebugDocumentContext {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}iCharPos: LongWord;
                                                                              {VT_19:0}cChars: LongWord;
                                                                              {VT_29:2}out ppddc: IDebugDocumentContext): HResult; stdcall;
    function  BringDocumentToTop {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  BringDocumentContextToTop {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pddc: IDebugDocumentContext): HResult; stdcall;
  end;

  IActiveScriptErrorDebug = interface(IActiveScriptError)
    ['{51973C12-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetDocumentContext {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppssc: IDebugDocumentContext): HResult; stdcall;
    function  GetStackFrame {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdsf: IDebugStackFrame): HResult; stdcall;
  end;

  IDebugDocumentInfo = interface(IUnknown)
    ['{51973C1F-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetName {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:0}dnt: tagDOCUMENTNAMETYPE;
                                                           {VT_8:1}out pbstrName: WideString): HResult; stdcall;
    function  GetDocumentClassId {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}out pclsidDocument: _GUID): HResult; stdcall;
  end;

  IDebugDocument = interface(IDebugDocumentInfo)
    ['{51973C21-CB0C-11D0-B5C9-00A0244A0E7A}']
  end;

  IEnumDebugCodeContexts = interface(IUnknown)
    ['{51973C1D-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  RemoteNext {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord;
                                                              {VT_29:2}out pscc: IDebugCodeContext;
                                                              {VT_19:1}out pceltFetched: LongWord): HResult; stdcall;
    function  Skip {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord): HResult; stdcall;
    function  Reset {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  Clone {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppescc: IEnumDebugCodeContexts): HResult; stdcall;
  end;

  IDebugDocumentHost = interface(IUnknown)
    ['{51973C27-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetDeferredText {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_19:0}dwTextStartCookie: LongWord;
                                                                   {VT_2:1}var pcharText: Smallint;
                                                                   {VT_18:1}var pstaTextAttr: Word;
                                                                   {VT_19:1}var pcNumChars: LongWord;
                                                                   {VT_19:0}cMaxChars: LongWord): HResult; stdcall;
    function  GetScriptTextAttributes {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_31:0}pstrCode: PWideChar;
                                                                           {VT_19:0}uNumCodeChars: LongWord;
                                                                           {VT_31:0}pstrDelimiter: PWideChar;
                                                                           {VT_19:0}dwFlags: LongWord;
                                                                           {VT_18:1}var pattr: Word): HResult; stdcall;
    function  OnCreateDocumentContext {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_13:1}out ppunkOuter: IUnknown): HResult; stdcall;
    function  GetPathName {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_8:1}out pbstrLongName: WideString;
                                                               {VT_3:1}out pfIsOriginalFile: Integer): HResult; stdcall;
    function  GetFileName {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_8:1}out pbstrShortName: WideString): HResult; stdcall;
    function  NotifyChanged {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
  end;

  IRemoteDebugApplication = interface(IUnknown)
    ['{51973C30-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  ResumeFromBreakPoint {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_29:1}const prptFocus: IRemoteDebugApplicationThread;
                                                                        {VT_29:0}bra: tagBREAKRESUME_ACTION;
                                                                        {VT_29:0}era: tagERRORRESUMEACTION): HResult; stdcall;
    function  CauseBreak {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  ConnectDebugger {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pad: IApplicationDebugger): HResult; stdcall;
    function  DisconnectDebugger {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  GetDebugger {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pad: IApplicationDebugger): HResult; stdcall;
    function  CreateInstanceAtApplication {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_29:1}var rclsid: _GUID;
                                                                               {VT_13:0}const pUnkOuter: IUnknown;
                                                                               {VT_19:0}dwClsContext: LongWord;
                                                                               {VT_29:1}var riid: _GUID;
                                                                               {VT_13:1}out ppvObject: IUnknown): HResult; stdcall;
    function  QueryAlive {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  EnumThreads {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pperdat: IEnumRemoteDebugApplicationThreads): HResult; stdcall;
    function  GetName {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_8:1}out pbstrName: WideString): HResult; stdcall;
    function  GetRootNode {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdanRoot: IDebugApplicationNode): HResult; stdcall;
    function  EnumGlobalExpressionContexts {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppedec: IEnumDebugExpressionContexts): HResult; stdcall;
  end;


  IDebugDocumentProvider = interface(IDebugDocumentInfo)
    ['{51973C20-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetDocument {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppssd: IDebugDocument): HResult; stdcall;
  end;

  IDebugApplication = interface(IRemoteDebugApplication)
    ['{51973C32-CB0C-11D0-B5C9-00A0244A0E7A}']
   function  SetName {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pstrName: PWideChar): HResult; stdcall;
    function  StepOutComplete {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  DebugOutput {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pstr: PWideChar): HResult; stdcall;
    function  StartDebugSession {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  HandleBreakPoint {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:0}br: tagBREAKREASON;
                                                                    {VT_29:1}out pbra: tagBREAKRESUME_ACTION): HResult; stdcall;
    function  Close {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  GetBreakFlags {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_19:1}out pabf: LongWord;
                                                                 {VT_29:2}out pprdatSteppingThread: IRemoteDebugApplicationThread): HResult; stdcall;
    function  GetCurrentThread {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pat: IDebugApplicationThread): HResult; stdcall;
    function  CreateAsyncDebugOperation {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:1}const psdo: IDebugSyncOperation;
                                                                             {VT_29:2}out ppado: IDebugAsyncOperation): HResult; stdcall;
    function  AddStackFrameSniffer {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:1}const pdsfs: IDebugStackFrameSniffer;
                                                                        {VT_19:1}out pdwCookie: LongWord): HResult; stdcall;
    function  RemoveStackFrameSniffer {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}dwCookie: LongWord): HResult; stdcall;
    function  QueryCurrentThreadIsDebuggerThread {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  SynchronousCallInDebuggerThread {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_29:1}const pptc: IDebugThreadCall;
                                                                                   {VT_19:0}dwParam1: LongWord;
                                                                                   {VT_19:0}dwParam2: LongWord;
                                                                                   {VT_19:0}dwParam3: LongWord): HResult; stdcall;
    function  CreateApplicationNode {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppdanNew: IDebugApplicationNode): HResult; stdcall;
    function  FireDebuggerEvent {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:1}var riid: _GUID;
                                                                     {VT_13:0}const punk: IUnknown): HResult; stdcall;
    function  HandleRuntimeError {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_29:1}const pErrorDebug: IActiveScriptErrorDebug;
                                                                      {VT_29:1}const pScriptSite: IActiveScriptSite;
                                                                      {VT_29:1}out pbra: tagBREAKRESUME_ACTION;
                                                                      {VT_29:1}out perra: tagERRORRESUMEACTION;
                                                                      {VT_3:1}out pfCallOnScriptError: Integer): HResult; stdcall;
    function  FCanJitDebug {Flags(1), (0/0) CC:4, INV:1, DBG:6}: Integer; stdcall;
    function  FIsAutoJitDebugEnabled {Flags(1), (0/0) CC:4, INV:1, DBG:6}: Integer; stdcall;
    function  AddGlobalExpressionContextProvider {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:1}const pdsfs: IProvideExpressionContexts;
                                                                                      {VT_19:1}out pdwCookie: LongWord): HResult; stdcall;
    function  RemoveGlobalExpressionContextProvider {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}dwCookie: LongWord): HResult; stdcall;
  end;


  IDebugApplicationNode = interface(IDebugDocumentProvider)
    ['{51973C34-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  EnumChildren {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pperddp: IEnumDebugApplicationNodes): HResult; stdcall;
    function  GetParent {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pprddp: IDebugApplicationNode): HResult; stdcall;
    function  SetDocumentProvider {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pddp: IDebugDocumentProvider): HResult; stdcall;
    function  Close {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  Attach {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pdanParent: IDebugApplicationNode): HResult; stdcall;
    function  Detach {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
  end;

  IDebugCodeContext = interface(IUnknown)
    ['{51973C13-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetDocumentContext {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppsc: IDebugDocumentContext): HResult; stdcall;
    function  SetBreakPoint {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:0}bps: tagBREAKPOINT_STATE): HResult; stdcall;
  end;

  IRemoteDebugApplicationThread = interface(IUnknown)
    ['{51973C37-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetSystemThreadId {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:1}out dwThreadId: LongWord): HResult; stdcall;
    function  GetApplication {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pprda: IRemoteDebugApplication): HResult; stdcall;
    function  EnumStackFrames {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppedsf: IEnumDebugStackFrames): HResult; stdcall;
    function  GetDescription {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_8:1}out pbstrDescription: WideString;
                                                                  {VT_8:1}out pbstrState: WideString): HResult; stdcall;
    function  SetNextStatement {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:1}const pStackFrame: IDebugStackFrame;
                                                                    {VT_29:1}const pCodeContext: IDebugCodeContext): HResult; stdcall;
    function  GetState {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:1}out pState: LongWord): HResult; stdcall;
    function  Suspend {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:1}out pdwCount: LongWord): HResult; stdcall;
    function  Resume {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:1}out pdwCount: LongWord): HResult; stdcall;
    function  GetSuspendCount {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:1}out pdwCount: LongWord): HResult; stdcall;
  end;

  IApplicationDebugger = interface(IUnknown)
    ['{51973C2A-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  QueryAlive {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  CreateInstanceAtDebugger {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_29:1}var rclsid: _GUID;
                                                                            {VT_13:0}const pUnkOuter: IUnknown;
                                                                            {VT_19:0}dwClsContext: LongWord;
                                                                            {VT_29:1}var riid: _GUID;
                                                                            {VT_13:1}out ppvObject: IUnknown): HResult; stdcall;
    function  onDebugOutput {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pstr: PWideChar): HResult; stdcall;
    function  onHandleBreakPoint {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_29:1}const prpt: IRemoteDebugApplicationThread;
                                                                      {VT_29:0}br: tagBREAKREASON;
                                                                      {VT_29:1}const pError: IActiveScriptErrorDebug): HResult; stdcall;
    function  onClose {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  onDebuggerEvent {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_29:1}var riid: _GUID;
                                                                   {VT_13:0}const punk: IUnknown): HResult; stdcall;
  end;

  IEnumRemoteDebugApplicationThreads = interface(IUnknown)
    ['{51973C3C-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  RemoteNext {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord;
                                                              {VT_29:2}out ppdat: IRemoteDebugApplicationThread;
                                                              {VT_19:1}out pceltFetched: LongWord): HResult; stdcall;
    function  Skip {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord): HResult; stdcall;
    function  Reset {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  Clone {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pperdat: IEnumRemoteDebugApplicationThreads): HResult; stdcall;
  end;

  IEnumDebugExpressionContexts = interface(IUnknown)
    ['{51973C40-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  RemoteNext {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord;
                                                              {VT_29:2}out pprgdec: IDebugExpressionContext;
                                                              {VT_19:1}out pceltFetched: LongWord): HResult; stdcall;
    function  Skip {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord): HResult; stdcall;
    function  Reset {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  Clone {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppedec: IEnumDebugExpressionContexts): HResult; stdcall;
  end;

  IDebugSyncOperation = interface(IUnknown)
    ['{51973C1A-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetTargetThread {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppatTarget: IDebugApplicationThread): HResult; stdcall;
    function  Execute {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_13:1}out ppunkResult: IUnknown): HResult; stdcall;
    function  InProgressAbort {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
  end;

  IDebugAsyncOperation = interface(IUnknown)
    ['{51973C1B-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetSyncDebugOperation {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppsdo: IDebugSyncOperation): HResult; stdcall;
    function  Start {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const padocb: IDebugAsyncOperationCallBack): HResult; stdcall;
    function  Abort {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  QueryIsComplete {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  GetResult {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_25:1}out phrResult: HResult;
                                                             {VT_13:1}out ppunkResult: IUnknown): HResult; stdcall;
  end;

  IDebugStackFrameSniffer = interface(IUnknown)
    ['{51973C18-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  EnumStackFrames {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppedsf: IEnumDebugStackFrames): HResult; stdcall;
  end;


  IDebugThreadCall = interface(IUnknown)
    ['{51973C36-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  ThreadCallHandler {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}dwParam1: LongWord;
                                                                     {VT_19:0}dwParam2: LongWord;
                                                                     {VT_19:0}dwParam3: LongWord): HResult; stdcall;
  end;

  IProvideExpressionContexts = interface(IUnknown)
    ['{51973C41-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  EnumExpressionContexts {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppedec: IEnumDebugExpressionContexts): HResult; stdcall;
  end;

  IEnumDebugApplicationNodes = interface(IUnknown)
    ['{51973C3A-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  RemoteNext {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord;
                                                              {VT_29:2}out pprddp: IDebugApplicationNode;
                                                              {VT_19:1}out pceltFetched: LongWord): HResult; stdcall;
    function  Skip {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord): HResult; stdcall;
    function  Reset {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  Clone {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out pperddp: IEnumDebugApplicationNodes): HResult; stdcall;
  end;

  IEnumDebugStackFrames = interface(IUnknown)
    ['{51973C1E-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  RemoteNext {Flags(1), (3/3) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord;
                                                              {VT_29:1}out prgdsfd: tagDebugStackFrameDescriptor;
                                                              {VT_19:1}out pceltFetched: LongWord): HResult; stdcall;
    function  Skip {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_19:0}celt: LongWord): HResult; stdcall;
    function  Reset {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  Clone {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppedsf: IEnumDebugStackFrames): HResult; stdcall;
  end;

  IDebugExpressionContext = interface(IUnknown)
    ['{51973C15-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  ParseLanguageText {Flags(1), (5/5) CC:4, INV:1, DBG:6}({VT_31:0}pstrCode: PWideChar;
                                                                     {VT_23:0}nRadix: SYSUINT;
                                                                     {VT_31:0}pstrDelimiter: PWideChar;
                                                                     {VT_19:0}dwFlags: LongWord;
                                                                     {VT_29:2}out ppe: IDebugExpression): HResult; stdcall;
    function  GetLanguageInfo {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_8:1}out pbstrLanguageName: WideString;
                                                                   {VT_29:1}out pLanguageID: _GUID): HResult; stdcall;
  end;

  IDebugAsyncOperationCallBack = interface(IUnknown)
    ['{51973C1C-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  onComplete {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
  end;

  IDebugExpression = interface(IUnknown)
    ['{51973C14-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  Start {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:1}const pdecb: IDebugExpressionCallBack): HResult; stdcall;
    function  Abort {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  QueryIsComplete {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  GetResultAsString {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_25:1}out phrResult: HResult;
                                                                     {VT_8:1}out pbstrResult: WideString): HResult; stdcall;
    function  GetResultAsDebugProperty {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_25:1}out phrResult: HResult;
                                                                            {VT_29:2}out ppdp: IDebugProperty): HResult; stdcall;
  end;

  IDebugExpressionCallBack = interface(IUnknown)
    ['{51973C16-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  onComplete {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
  end;

  IDebugProperty = interface(IUnknown)
    ['{51973C31-CB0C-11D0-B5C9-00A0244A0E7A}']
  end;

  IDebugApplicationThread = interface(IRemoteDebugApplicationThread)
    ['{51973C38-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  SynchronousCallIntoThread {Flags(1), (4/4) CC:4, INV:1, DBG:6}({VT_29:1}const pstcb: IDebugThreadCall; 
                                                                             {VT_19:0}dwParam1: LongWord; 
                                                                             {VT_19:0}dwParam2: LongWord; 
                                                                             {VT_19:0}dwParam3: LongWord): HResult; stdcall;
    function  QueryIsCurrentThread {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  QueryIsDebuggerThread {Flags(1), (0/0) CC:4, INV:1, DBG:6}: HResult; stdcall;
    function  SetDescription {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pstrDescription: PWideChar): HResult; stdcall;
    function  SetStateString {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_31:0}pstrState: PWideChar): HResult; stdcall;
  end;

  IDebugStackFrame = interface(IUnknown)
    ['{51973C17-CB0C-11D0-B5C9-00A0244A0E7A}']
    function  GetCodeContext {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppcc: IDebugCodeContext): HResult; stdcall;
    function  GetDescriptionString {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_3:0}fLong: Integer; 
                                                                        {VT_8:1}out pbstrDescription: WideString): HResult; stdcall;
    function  GetLanguageString {Flags(1), (2/2) CC:4, INV:1, DBG:6}({VT_3:0}fLong: Integer; 
                                                                     {VT_8:1}out pbstrLanguage: WideString): HResult; stdcall;
    function  GetThread {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppat: IDebugApplicationThread): HResult; stdcall;
    function  GetDebugProperty {Flags(1), (1/1) CC:4, INV:1, DBG:6}({VT_29:2}out ppDebugProp: IDebugProperty): HResult; stdcall;
  end;

implementation


end.


