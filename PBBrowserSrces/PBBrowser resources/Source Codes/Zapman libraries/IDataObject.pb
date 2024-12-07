; ********************************************************
;
;           Implementation of an 'IDataObject'
;
;               For Windows only - PB 6.11
;                   Sept 2024 - Zapman
;
; ********************************************************
;
; Big thanks to Nico from PureBasic Forum. As far as I know,
; he was the first to try to implemante an IDataObject in PureBasic
; and he made an important part of the job presented here.
;
; Ten years ago, starting from his work, I began to work on that
; project for the need of a Drag'n Drop functionnalities for one
; of my program. This year, finally I decided to push the implementation
; as far as I could, in order to get a fully functionnal IDataObject
; for any futur purposes.
;
; This implementation allows the IDataObject to support the 'SetData'
; 'GetData'and  GetDataHere' methods for most of possible formats.
; The object also responds to requests for EnumFormatEtc, EnumAdvise,
; QueryGetData, QueryInterface, AddRef, DAdvise, DUnAdvise and Release.
;
; This IDataObject is conceived to be as generic as possible, but,
; of course, you can adapt its code for your specific needs.
;
;
; The 'IDataObject' interface in Windows allows manipulating and transferring data 
; between different applications or processes, for example, via the clipboard or 
; drag-and-drop. It allows access to data in multiple formats, grouped in an object 
; that other applications can use. The main methods are GetData (to retrieve data), 
; SetData (to set data), and EnumFormatEtc (to list available data formats). IDataObject 
; is often used in OLE environments to link or embed objects in documents or applications.
;
; In Windows, an object represents an entity that contains data and functions. 
; An object has interfaces, which are ways to access its functionalities, such as 
; reading files or managing data streams. Objects in Windows often use 
; COM (Component Object Model), where a base interface like IUnknown allows managing 
; references and communication between objects. This model helps software components 
; work well together by adhering to specific interface rules.
;
; Object-oriented programming (OOP) organizes code into objects that contain both 
; data (called properties) and functions (called methods) to manipulate them. 
; Each object is like a small autonomous unit, representing something from the real 
; world, like a window or a file. This makes programs more structured and reusable. 
; Objects can interact with each other, making the code more flexible and modular.
;
; In comparison, traditional programming, often called 'procedural programming', 
; follows a linear plan where instructions are executed one after another. Data 
; and functions are separated, and the program is usually organized into large 
; sequences of code. Although simpler at first, this approach can have limitations 
; in managing large projects with many interactions between different parts of 
; the program. These limitations can be perfectly managed with a solid structure 
; and segmenting the code into more or less autonomous procedures. But some prefer 
; a more radical approach where each function is organized as an independent object 
; with a strict list of "methods" that allow interacting with it. This is the choice 
; made for many subsets of Windows functions, and it is now inevitable, when 
; interacting with Windows through its APIs, to understand the principle of this 
; form of programming.
;
; For everything related to data manipulation (reading and writing files, 
; copy-pasting, audio and video streams, drag and drop, and display on the screen), 
; it is almost essential, if you want to surpass certain limitations of PureBasic's 
; original functions, to know how to create and manage an IDataObject. Unfortunately, 
; while Windows often imposes the use of this type of Object, no API provides 
; a simple way to create it, and a quick look at forums of various programming 
; languages shows that this question comes up very often: "how to create an IDataObject?"
;
; This library was designed to answer that question.
;
; As mentioned earlier, an 'Object' contains both data and 'Methods' 
; allowing interaction with this data. 
; What must be strictly respected for our IDataObject to function correctly 
; is the number of its methods, the roles they must play, the number and type of 
; parameters they must accept as input, and the parameters they must provide as output.
;
; However, the way of storing the data is almost entirely free, even though 
; this data must be formatted according to criteria set by Windows.
;
;             ********************************************************
;
;-                                   The Data
;
;             ********************************************************
;
; The data of an IDataObject is always defined by two structures that work 
; in tandem:
; • A FormatEtc structure that indicates what kind of data it is.
;     - FormatEtc tells us, for example, if it is ASCII text (CF_Text), Unicode text 
;       (CF_Unicode), a bitmap image (CF_Bitmap), a bitmap encoded as Dib (CF_Dib), 
;       an image encoded in a metafile (CF_METAFILEPICT), etc.
;     - With its 'TYMED' field, FormatEtc also informs us how the data is stored: 
;       on disk, in memory, in a stream, etc.
; • A StgMedium structure that contains the data or specifies how to access it.
;       Simple data is usually short character strings stored directly in StgMedium. 
;       Heavier data is often stored in external sources to which StgMedium provides 
;       access. And as Windows likes "Objects", it is not uncommon for StgMedium 
;       to simply provide the interface of another object that must be queried to 
;       retrieve the data.
;
; Our IDataObject stores the FormatEtc-StgMedium pairs in a structure called 'IDO_DataSaveStruct'.
; We create a list IDO_DataSave.IDO_DataSaveStruct() in which all the FormatEtc-StgMedium 
; pairs (i.e., all the IDO_DataSaveStruct structures) we may need are stored.
;
Structure IDO_DataSaveStruct
  ; The role of this structure is to combine StgMedium-FormatEtc pairs
  ; so they can be stored in memory.
  m_stgmed.StgMedium ; We'll store here pointers to STGMEDIUM objects
  m_format.FormatEtc ; We'll store here pointers to FORMATETC objects
  *DataOwner         ; Proprietary of this data.
EndStructure
;
NewList IDO_DataSave.IDO_DataSaveStruct()
;
; But a link must be established between the created IDataObjects and the data 
; that belongs to them.
; For this purpose, we create a 'iData' structure that will gather the address 
; of the IDataObject, the addresses of the data, as well as other parameters useful 
; for managing the IDataObject.
;
#DataMax = 100 ; If necessary, you can increase #DataMax so that your IDataObjects can 
               ; manage a larger number of data sets.
;
Structure IDO_DetailsStruct
  ; The role of this structure is to bring together
  ; everything necessary for the functioning of an IDataObject.
  pDataObject.IDataobject ; Pointer to an IDataObject interface
  Refcount.l              ; Number of times the IDataObject has been queried by other process
  m_count.l               ; Number of different formats registered (set) by the IDataObject
  m_stgmed.i[#DataMax]    ; We'll store here pointers to STGMEDIUM objects
  m_format.i[#DataMax]    ; We'll store here pointers to FORMATETC objects
  pDataAdviseHolder.IDataAdviseHolder     ; This will be used for the DAdvise management.
EndStructure
;
NewList IDO_Details.IDO_DetailsStruct()
;
; Since an IDataObject can simultaneously contain several types of data,
; our 'IDO_DetailsStruct' structure includes two pointer arrays: m_stgmed.i[] which points 
; to the IDataObject's 'StgMedium' and m_format[] which points to the corresponding 'FormatEtc'.
; The 'm_count' field of this structure will indicate how many data sets the IDataObject has,
; and the 'Refcount' field will count the number of processes querying our IDataObject.
;
; The IDO_Details.IDO_DetailsStruct() list will keep our information in memory 
; as long as necessary. This way, we can create several IDataObjects that can function 
; independently at the same time. Each element of the list will represent an independent IDataObject.
;
; A pointer to one of our IDataObjects will also point to the IDO_DetailsStruct 
; associated with it and the element of the IDO_Details list that contains this IDO_DetailsStruct.
; All of this resides at the same address.
; This may seem very strange to developers unfamiliar with pointers, but you can write:
;   CurrentRefCount = *UPointer.IDO_DetailsStruct\Refcount
; * MyData.IDataObject = *UPointer
;   ChangeCurrentElement(IDO_Details(), *UPointer)
; with the same value for *UPointer.
;
; This is explained by the fact that a pointer to a list element points to 
; the content of that element, and a pointer to a structure also points to 
; the content of this structure.
; Now, our IDO_DetailsStruct structure has pDataObject as its first element, so a pointer to 
; this structure points to pDataObject.
; Similarly, a pointer to an element of IDO_Details() points to its content, that is, 
; to the IDO_DetailsStruct structure which contains pDataObject (among other things).
;
; Nothing else is needed to store the data.
;
; Please note that the approach we chose to follow is not imposed by anything other 
; than the need to provide the IDataObject with what it needs to function.
; The data could be organized in a completely different way without causing 
; problems for the processes interacting with our object, as it is the object itself 
; that accesses the data and "knows" how to do so.
;
; Now, let's see how to implement the methods of an IDataObject. This time, we 
; have no choice but to strictly follow the standards imposed for constructing 
; an object and those imposed by Windows for everything governing the functioning 
; of an IDataObject.
;
;     ********************************************************
;
;-                        The 'Methods'
;
;     ********************************************************
;
; What is called a 'method' in the context of object-oriented programming (OOP) 
; is nothing more than a procedure in PureBasic. And an 'object' is nothing more than 
; a collection of 'methods' capable of manipulating the object's data. This collection 
; of 'methods' is called an interface.
;
; When we call an object through the 'QueryInterface' method,
; we are asking it to provide us with the list of addresses of its methods/procedures.
;
; Of course, this list must exist! The first thing to do, to build
; an object, is to establish the list of methods that will allow interacting with it.
; This list appears at the end of this file in the DataSection, but I reproduce it here for
; the purpose of this explanation:
;
; DataSection
;   VTable_IDataObject:
;     Data.i @dataobject_QueryInterface(), @dataobject_AddRef(), @dataobject_Release()
;     Data.i @dataobject_GetData(),@dataobject_GetDataHere(),@dataobject_QueryGetData()
;     Data.i @dataobject_GetCanonicalFormatEtc(),@dataobject_SetData(),@dataobject_EnumFormatEtc()
;     Data.i @dataobject_DAdvise(),@dataobject_DUnadvise(),@dataobject_EnumDAdvise()
; EndDataSection
;
; Note that this list contains nothing but pointers to our
; procedures. It does not describe the parameters of these procedures, nor their return values.
; Only the bare essentials to be able to call them.
;
; But then, how do we know what parameters to provide them?
;
; There is no other solution than to read the documentation regarding IDataObject.
; One of the downsides of object-oriented programming is that you need to know in advance
; how to interact with an object, what methods it has, and what parameters each
; of its methods requires.
;
; The structure of an IDataObject has been decided by Windows. Therefore, it is the Windows documentation
; that explains how many methods such an object must have and in what order
; to arrange the method addresses in the table that defines the interface.
; If we were to swap @dataobject_QueryInterface() and @dataobject_AddRef() in the table above,
; our object would no longer be compatible with all Windows processes that wish
; to interact with an IDataObject.
; An IDataObject must have at least twelve methods whose addresses must be
; arranged in the table in the order shown above.
; Of course, it is not enough to give the object a list of methods/procedures for 
; it to start working. You will need to WRITE these procedures and ensure they
; are capable of handling the parameters they will be given.
;
; But, before tackling this tedious task, let's start with the dessert and look at
; the two procedures YOU will need to create an IDataObject and destroy it after use:
;
;      ********************************************************
;
;-           Creation and Destruction of the IDataObject
;
;      ********************************************************
;
;
Procedure CreateIDataObject()
  ;
  ; Create a new IDataObject in memory and return a pointer to it.
  ;
  Shared IDO_Details()
  ;
  AddElement(IDO_Details())
  IDO_Details()\pDataObject = ?VTable_IDataObject
  IDO_Details()\m_count = 0
  IDO_Details()\Refcount = 0
  ;
  ProcedureReturn IDO_Details()
EndProcedure
;
Procedure IsIDataObject(*DataObjectDet)
  ;
  ; Check if the *DataObjectDet given as parameter
  ; is an active IDataObject.
  ;
  Shared IDO_Details()
  ; Check if the object exists or exists anymore in memory:
  ForEach IDO_Details()
    If IDO_Details() = *DataObjectDet
      ProcedureReturn #True
    EndIf
  Next
EndProcedure
;
Procedure DestroyIDataObject(*pDataObject.IDataObject)
  ;
  ; Destroy all data contained by *pDataObject
  ; the object itself and the IDO_Details.IDO_DetailsStruct() element
  ; wich was containing the object.
  ;
  If IsIDataObject(*pDataObject)
    While *pDataObject\Release() : Wend
  Else
    Debug "Wrong parameter for DestroyIDataObject(): this object doesn't exist."
  EndIf
  ;
EndProcedure
;
; Seems incredibly simple, right?
; The first procedure merely creates an element
; in the IDO_Details() table and fills the first field
; of its structure with the method table.
; As we saw above, the pointer to the IDO_Details()
; element returned by this procedure also points to the corresponding
; iData structure, and therefore to the IDataObject that is
; the first field of this structure.
;
; The DestroyDataObject() procedure is even simpler:
; it calls the 'Release' method of the object until it
; destroys itself.
;
; We have two things left to do:
; 1• define a few procedures that will be useful for debugging or 
;    for the functionality of our methods,
; 2• then define the methods themselves.
;
; ********************************************************
;-                 Debugging Procedure
;
; ********************************************************
;
; Our procedures/'methods' are loaded with checks and controls to
; prevent our IDataObject from crashing the machine in any way.
;
; They are also incredibly "talkative" if you let them speak.
; The constant SilentMethods is intended to silence them to avoid
; flooding the debugger window with messages about what's happening
; in the life of our IDataObjects.
; When SilentMethods = 0 (you can set it to this value if you want to see what
; is going on), every interaction between the IDataObject and other processes,
; and each step of these interactions, is documented. Make this change and you will
; notice that an IDataObject communicates a lot with the processes that use it.
;
; If you want, you can modify the IDO_RegisterResults procedure so that it
; logs all the messages to a debug file, rather than displaying them in
; the debugger window.
;
Global SilentMethods = 1
;
Procedure IDO_RegisterResults(Mess$)
  If SilentMethods = 0
    Debug Mess$
  EndIf
EndProcedure
;
;
; ********************************************************
;
;-                 Generic Procedures
;
; ********************************************************
;
; To obtain clear debugging information, it was
; necessary to develop a number of procedures
; to get, for example, the name of certain constants
; or to return explanatory text about errors that
; may occur during operation.
;
; The following four procedures cannot, for now,
; be removed from the code because they are used
; for debugging certain situations inside the IDataObject methods.
; If you examine the code of these methods, you will notice that they
; often compute strings intended for the IDO_RegisterResults() procedure
; which is only used for debugging the program.
; If you want to streamline the code, you can remove all
; calls to IDO_RegisterResults() and all string calculations
; intended for debugging. It will then be possible to remove
; these four procedures.
;
CompilerIf Not Defined(GetWinErrorMessage, #PB_Procedure)
  Procedure.s GetWinErrorMessage(errorCode = 0)
    ;
    ; Retourne un texte qui explicite l'erreur dont le numéro
    ; est passé dans le paramètre 'errorCode'.
    : 
    Protected messageBuffer$ ; Buffer pour le message
    Protected MLength, length, ErrorN, MHandle
    ;
    If errorCode = 0
      errorCode = GetLastError_()
    EndIf
    If errorCode
      MLength = 500
      messageBuffer$ = Space((MLength + 1) * 2)
      length = FormatMessage_(#FORMAT_MESSAGE_FROM_SYSTEM, #Null, errorCode, 0, @messageBuffer$, MLength, #Null)
      If length < 1
        #ERROR_INTERNET_EXTENDED_ERROR = 12003
        If errorCode = #ERROR_INTERNET_EXTENDED_ERROR
          ErrorN = 0
          length = InternetGetLastResponseInfo_(@ErrorN, @messageBuffer$, @MLength)
          If length = 0
            length = FormatMessage_(#FORMAT_MESSAGE_FROM_SYSTEM, #Null, GetLastError_(), 0, @messageBuffer$, MLength, #Null)
          EndIf 
        Else
          If errorCode > 11999 And errorCode < 13000
            #FORMAT_MESSAGE_FROM_HMODULE = $00000800
            MHandle = GetModuleHandle_("wininet.dll")
            length = FormatMessage_(#FORMAT_MESSAGE_FROM_HMODULE, MHandle, errorCode, 0, @messageBuffer$, MLength, #Null)
          EndIf
        EndIf 
      EndIf
      messageBuffer$ = ReplaceString(messageBuffer$, Chr(10), "")
      messageBuffer$ = ReplaceString(messageBuffer$, Chr(13), "")
      messageBuffer$ = Trim(messageBuffer$)
      If messageBuffer$ = ""
        messageBuffer$ = "Unknown error"
      EndIf
      ProcedureReturn messageBuffer$ + "  ($" + Hex(errorCode) + ")"
    EndIf
  EndProcedure
CompilerEndIf
;
Procedure.s GetFormatName(cfFormat.w)
  ;
  ; Retourne un nom de format en fonction du code passé dans cfFormat
  ;
  Protected formatName$ = Space(256)
  Protected Format = cfFormat
  If Format < 0 : Format + 65536 : EndIf
  ;
  Select Format
    ; Formats standards
    Case #CF_TEXT : formatName$ = "CF_TEXT"
    Case #CF_BITMAP : formatName$ = "CF_BITMAP"
    Case #CF_METAFILEPICT : formatName$ = "CF_METAFILEPICT"
    Case #CF_SYLK : formatName$ = "CF_SYLK"
    Case #CF_DIF : formatName$ = "CF_DIF"
    Case #CF_TIFF : formatName$ = "CF_TIFF"
    Case #CF_OEMTEXT : formatName$ = "CF_OEMTEXT"
    Case #CF_DIB : formatName$ = "CF_DIB"
    Case #CF_PALETTE : formatName$ = "CF_PALETTE"
    Case #CF_PENDATA : formatName$ = "CF_PENDATA"
    Case #CF_RIFF : formatName$ = "CF_RIFF"
    Case #CF_WAVE : formatName$ = "CF_WAVE"
    Case #CF_UNICODETEXT : formatName$ = "CF_UNICODETEXT"
    Case #CF_ENHMETAFILE : formatName$ = "CF_ENHMETAFILE"
    Case #CF_HDROP : formatName$ = "CF_HDROP"
    Case #CF_LOCALE : formatName$ = "CF_LOCALE"
    Case #CF_DIBV5 : formatName$ = "CF_DIBV5"  
    ;
    Case #CF_OWNERDISPLAY : formatName$ = "CF_OWNERDISPLAY"
    Case #CF_DSPTEXT : formatName$ = "CF_DSPTEXT"
    Case #CF_DSPBITMAP : formatName$ = "CF_DSPBITMAP"
    Case #CF_DSPMETAFILEPICT : formatName$ = "CF_DSPMETAFILEPICT"
    Case #CF_DSPENHMETAFILE : formatName$ = "CF_DSPENHMETAFILE"
    Case #CF_PRIVATEFIRST : formatName$ = "CF_PRIVATEFIRST"
    Case #CF_PRIVATELAST : formatName$ = "CF_PRIVATELAST"
    Case #CF_GDIOBJFIRST : formatName$ = "CF_GDIOBJFIRST"
    Case #CF_GDIOBJLAST : formatName$ = "CF_GDIOBJLAST"
    Default
      GetClipboardFormatName_(Format, @formatName$, 256)
  EndSelect
  ;
  If formatName$ = ""
    formatName$ = "Unknown format (" + Str(Format) + ")"
  EndIf
  ;
  ProcedureReturn formatName$
EndProcedure
;
Procedure.s GetTymedName(Tymed)
  ;
  ; Retourne le nom de la constante dont la valeur est passée dans le paramètre 'Tymed'.
  ;
  Protected Tymed$ = ""
  ;
  If Tymed = 0
    tymed$ = "TYMED_NULL"
  Else
    If Tymed & 1
      Tymed$ + "TYMED_HGLOBAL + "
    EndIf
    If Tymed & 2
      Tymed$ + "TYMED_FILE + "
    EndIf
    If Tymed & 4
      Tymed$ + "TYMED_ISTREAM + "
    EndIf
    If Tymed & 8
      Tymed$ + "TYMED_ISTORAGE + "
    EndIf
    If Tymed & 16
      Tymed$ + "TYMED_GDI + "
    EndIf
    If Tymed & 32
      Tymed$ + "TYMED_MFPICT + "
    EndIf
    If Tymed & 64
      Tymed$ + "TYMED_ENHMF"
    EndIf
  EndIf
  If Tymed$ = ""
    tymed$ = "Unknown tymed value: " + Str(Tymed)
  EndIf
  If Right(Tymed$, 3) = " + "
    Tymed$ = Left(Tymed$, Len(Tymed$) - 3)
  EndIf
  ProcedureReturn Tymed$
EndProcedure
;
Procedure.s StringFromCLSID(MyCLSID)
  ;
  Protected *CLSIDString, Ret$
  ; Appel de StringFromCLSID_ pour obtenir la chaîne CLSID
  Protected Result = StringFromCLSID_(@MyCLSID, @*CLSIDString)
  ;
  If Result = #S_OK
    ; Afficher la chaîne du CLSID
    Ret$ = PeekS(*CLSIDString, -1, #PB_Unicode)
    ;
    ; Libérer la mémoire allouée par StringFromCLSID
    CoTaskMemFree_(*CLSIDString)
  Else
    Ret$ = "Error while converting CLSID to String."
  EndIf
  ProcedureReturn Ret$
EndProcedure
;
Procedure CloneIDataObject(*IDOSource.IDataObject, *IDODest.IDataObject)
  ;
  Protected *enumFormat.IEnumFORMATETC, SC
  Protected formatEtc.FormatEtc, stgm.StgMedium
  ;
  SC = *IDOSource\EnumFormatEtc(#DATADIR_GET, @*enumFormat)
  If SC = #S_OK And *enumFormat
    ;
    ; Parcourir tous les formats disponibles
    While *enumFormat\Next(1, @formatEtc.FormatEtc, #Null) = #S_OK
      SC = *IDOSource\getData(formatEtc, @stgm.StgMedium)
      If SC = #S_OK
        SC = *IDODest\SetData(formatEtc, stgm, #True)
        If SC <> #S_OK
          Debug "Error while setting data: " + GetWinErrorMessage(SC)
        EndIf
      Else
        Debug "Error while getting data: " + GetWinErrorMessage(SC)
      EndIf
    Wend
    ;
    ; Libérer l'énumérateur
    *enumFormat\Release()
  Else
    Debug "Error : Unable to open format enumerator."
  EndIf
EndProcedure
;
; *******************************************************************
;
;-     Auxiliary procedures, essential for the operation
;                          of the IDataObject
;
; *******************************************************************
;
Procedure.i IsPTRValid(*ptr)
  Protected result.i = #False
  Protected mbi.MEMORY_BASIC_INFORMATION
  
  result = VirtualQuery_(*ptr, @mbi, SizeOf(MEMORY_BASIC_INFORMATION))
  If result And result <> #ERROR_INVALID_PARAMETER
    If mbi\State = #MEM_COMMIT And (mbi\Protect & (#PAGE_NOACCESS | #PAGE_GUARD)) = 0
      result = #True
    EndIf
  Else
    result = #False
  EndIf
  
  ProcedureReturn result
EndProcedure
;
Procedure CloneStgMedium(*DataObjectDet.IDO_DetailsStruct, *StgMed_Source.STGMedium, *StgMed_Dest.STGMedium, *AskedFormatEtc.FormatEtc, Method$ = "")
  ;
  ; This procedure will clone data contained in *StgMed_Source into $StgMef_Dest
  ; It supports TYMED_GDI, TYMED_HGLOBAL, TYMED_ENHMF, TYMED_MFPICT, TYMED_FILE, TYMED_ISTORAGE, And TYMED_ISTREAM formats.
  ;
  ; The 'Method$' parameter will determine how to proceed regarding the context.
  ; Method$ can be set with the "GetData" value, or "SetData", or "GetDataHere" or nothing.
  ;
  ; *AskedFormatEtc will also be usefull to precise the context of the copy. It contains
  ; the format and Tymed type asked by the caller of methods GetData and GetDataHere.
  ; *AskedFormatEtc\tymed can contain a combinaison of Tymed constants, when the caller
  ; means that it is asking and can accept more than one type of 'Tymed'.
  ;
  ; When TYMED_ISTREAM is into *AskedFormatEtc\Tymed, this procedure will generate IStream
  ; data from HGLOBAL, even if we don't get nativelly ISTREAM data.
  ;
  ; In the GetDataHere context, the data must be copied to a yet attributed support and
  ; the caller is responsible to free memory after usage. Because it's not possible to copy data
  ; into an existing Bitmap or an existing hGlobal, GetDataHere is not the good way to get data
  ; of hBitmap, hGlobal, ENHMF or MFPICT type, and our procedure will tell it to the caller by
  ; a DV_E_TYMED response.
  ; In opposite, GetDataHere will copy data to an existing IStorage or IStream without problem.
  ;
  ; When we are in the GetData context, we set pUnkForRelease to 0, wich means that the caller
  ; is also responsible to free memory after usage.
  ;
  Protected RetValue, DefaultRetValue, GenerateIStream
  Protected *StreamSrc.IStream, IStreamTest.IStream, IStorageTest.IStorage, TempHGlobal
  ;
  ; Check parameters:
  If IsIDataObject(*DataObjectDet) = 0
    MessageRequester("Oops!", "Program error: Unexisting *DataObjectDet while calling CloneStgMedium.")
    ProcedureReturn #E_FAIL
  ElseIf *StgMed_Source = 0
    IDO_RegisterResults("CloneStgMedium  --> Error with DataSource: *** NULL ARGUMENT ***")
    ProcedureReturn #E_INVALIDARG
  ElseIf *StgMed_Dest = 0
    IDO_RegisterResults("CloneStgMedium  --> Error with DataDest: *** NULL ARGUMENT ***")
    ProcedureReturn #E_INVALIDARG
  ElseIf *AskedFormatEtc = 0
    IDO_RegisterResults("CloneStgMedium  --> Error with *AskedFormatEtc: *** NULL ARGUMENT ***")
    ProcedureReturn #E_INVALIDARG
  ElseIf Method$ And Method$ <> "GetData" And Method$ <> "SetData" And Method$ <> "GetDataHere"
    IDO_RegisterResults("CloneStgMedium  --> Error with Method$: *** ''" + Method$ + "'' VALUE IS NOT VALID ***")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  GenerateIStream = #False
  If Method$ <> "SetData" And Not (*AskedFormatEtc\tymed & *StgMed_Source\tymed) And *AskedFormatEtc\tymed & #TYMED_ISTREAM
    ; The asked format-tymed doesn't match with the Source-tymed we get. ISTREAM is asked.
    ; Look if it's possible to generate IStream from our data:
    If *StgMed_Source\tymed = #TYMED_HGLOBAL Or *StgMed_Source\tymed = #TYMED_MFPICT
      GenerateIStream = #True
    EndIf
  EndIf
  ;
  If Method$ <> "GetDataHere"
    ; When Method$ is "GetDataHere", the destination medium is yet prepared
    ; by the caller. In other cases, we have to prepare it to receive our data:
    If GenerateIStream
      *StgMed_Dest\tymed          = #TYMED_ISTREAM
    Else
      *StgMed_Dest\tymed          = *StgMed_Source\tymed
    EndIf
    ;
    ; The pUnkForRelease of original data (*StgMed_Source\pUnkForRelease)
    ; may be different from zero. This original pUnkForRelease will be used
    ; when original data will be released. But the data we produce never need
    ; any 'special' method to be released. So, we set our pUnkForRelease to zero,
    ; wich mean that the caller is responsible to free data after usage.
    *StgMed_Dest\pUnkForRelease = 0
    *StgMed_Dest\hGlobal        = *StgMed_Source\hGlobal
  EndIf
  ;
  If *StgMed_Source\tymed = #TYMED_GDI And *AskedFormatEtc\cfFormat <> #CF_BITMAP
    RetValue = #E_NOTIMPL
    ;
  ElseIf *StgMed_Source\tymed = #TYMED_GDI Or *StgMed_Source\tymed = #TYMED_HGLOBAL Or *StgMed_Source\tymed = #TYMED_ENHMF Or *StgMed_Source\tymed = #TYMED_MFPICT
    ; All those media store their data in the same field, even if the 'Union' field of StgMedium can
    ; be named hBitmap, or hGlobal or hEnhMetaFile or hMetaFilePict, it is allways the same.
    ;
    DefaultRetValue = #S_OK
    ;
    If Method$ = "GetDataHere" And GenerateIStream = #False
      ; You cannot copy data into an existing Bitmap or an existing HGLOBAL.
      ; If IStream is asked, we will transfer data into an existing IStream,
      ; else, the GetDataHere method is not suitable for this type of data.
      RetValue = #DV_E_TYMED
    Else
      If *StgMed_Source\hGlobal = 0 
        If *StgMed_Source\tymed = #TYMED_HGLOBAL And GenerateIStream = 0
          ; To be as flexible as possible, we allow a null hGlobal for #TYMED_HGLOBAL data,
          ; because some applications as Photoshop seem to produce some data of that type,
          ; even if it has not any sense.
          RetValue = #S_OK
        Else
          ; If Source\tymed is different from #TYMED_HGLOBAL or if GenerateIStream is true,
          ; we don't accept to copy the data and we return an error value:
          If Method$ = "SetData" Or Method$ = ""
            RetValue = #E_INVALIDARG
          Else
            RetValue = #E_UNEXPECTED
          EndIf
        EndIf
      ElseIf *StgMed_Source\hGlobal <> 0 And (GenerateIStream = #False Or *StgMed_Source\tymed = #TYMED_HGLOBAL Or *StgMed_Source\tymed = #TYMED_MFPICT)
        ;
        ; Duplication of the data.
        ; OleDuplicateData is able to duplicate hBitmap, hGlobal hEnhMetaFile and hMetaFilePict.
        ; But if GenerateIStream is '#True', we remember that we can only generate a stream from hGlobal or MFPict data.
        ;
        TempHGlobal = OleDuplicateData_(*StgMed_Source\hGlobal, *AskedFormatEtc\cfFormat, #Null)
        If TempHGlobal
          RetValue = #S_OK
          If GenerateIStream = 0
            *StgMed_Dest\hGlobal = TempHGlobal
          EndIf
        Else
          RetValue = #E_OUTOFMEMORY
        EndIf
        ;
      Else
        RetValue = #DV_E_TYMED
      EndIf
    EndIf
  Else
    DefaultRetValue = #E_NOTIMPL
  EndIf
  ;
  If RetValue = #S_OK
    If *StgMed_Source\tymed = #TYMED_ISTORAGE
      ; Handle IStorage.
      If *StgMed_Source\pStg = 0 Or IsPTRValid(*StgMed_Source\pStg) = 0
        If Method$ = "SetData" Or Method$ = ""
          RetValue = #E_INVALIDARG
        Else
          RetValue = #E_UNEXPECTED
        EndIf
      ElseIf Method$ = "GetDataHere"
        If *StgMed_Dest\pStg = 0 Or IsPTRValid(*StgMed_Dest\pStg) = 0
          ; When calling the 'GetDataHere', the caller has to prepare the destination medium
          ; with a storage ready to receive our data. If *StgMed_Dest\pStg is null, the
          ; argument given by the caller is not valid:
          RetValue = #E_INVALIDARG
        Else
          ; Now we check that the destination IStorage is a valid pointer to an
          ; IStorage Interface, able to receive our data:
          RetValue = *StgMed_Dest\pStg\QueryInterface(?IID_IStorage, @IStorageTest)
          If RetValue <> #S_OK
            RetValue = #E_INVALIDARG
          Else
            IStorageTest\Release()
            ; Copy the data from our IStorage to the caller's IStorage:
            RetValue = *StgMed_Source\pStg\CopyTo(0, #Null, #Null, *StgMed_Dest\pStg)
          EndIf
        EndIf
      Else
        ; Method$ is different from "GetDataHere".
        ; We simply dupplicate interface
        RetValue = *StgMed_Source\pStg\QueryInterface(?IID_IStorage, @*StgMed_Dest\pStg)
      EndIf
      ;
    ElseIf *StgMed_Source\tymed = #TYMED_ISTREAM Or GenerateIStream
      ; Handle IStream
      If GenerateIStream
        IDO_RegisterResults("dataobject " + Method$ + " for " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed))
        IDO_RegisterResults(" ---> ON THE FLY GENERATION OF AN ISTREAM FROM THE HGLOBAL")
        RetValue = CreateStreamOnHGlobal_(TemphGlobal, #False, @*StreamSrc)
        If RetValue <> #S_OK
          GlobalFree_(TemphGlobal)
          *StreamSrc = 0
        EndIf
        ; If RetValue is OK, the generated Stream will be released when
        ; the destination medium will be released.
      Else
        *StreamSrc = *StgMed_Source\pstm
      EndIf
      ;
      If *StreamSrc = 0 Or IsPTRValid(*StreamSrc) = 0
        If Method$ = "SetData" Or Method$ = ""
          RetValue = #E_INVALIDARG
        Else
          RetValue = #E_UNEXPECTED
        EndIf
      ElseIf Method$ = "GetDataHere"
        If *StgMed_Dest\pstm = 0 Or IsPTRValid(*StgMed_Dest\pstm) = 0
          ; When calling the 'GetDataHere', the caller has to prepare the destination medium
          ; with a stream ready to receive our data. If *StgMed_Dest\pstm is null, the
          ; argument given by the caller is not valid:
          RetValue = #E_INVALIDARG
        Else
          ; Now we check that the destination IStream is a valid pointer to an
          ; IStream Interface, able to receive our data:
          RetValue = *StgMed_Dest\pstm\QueryInterface(?IID_IStream, @IStreamTest)
          If RetValue <> #S_OK
            RetValue = #E_INVALIDARG
          Else
            IStreamTest\Release()
            Protected bytesToCopy
            ; Copy the data from our IStream to the caller's IStream:
            ; Go to end of stream to know the size of data:
            *StreamSrc\Seek(0, #STREAM_SEEK_END, @bytesToCopy)
            ; Return to start to begin the copy:
            *StreamSrc\Seek(0, #STREAM_SEEK_SET, 0)
            RetValue = *StreamSrc\CopyTo(*StgMed_Dest\pstm, bytesToCopy, #Null, #Null)
          EndIf
        EndIf
      Else
        ; Duplicate the IStream interface
        RetValue = *StreamSrc\QueryInterface(?IID_IStream, @*StgMed_Dest\pstm)
        If RetValue = #S_OK
          ; We need to move the stream cursor to the end of the content.
          ; Microsoft documentation states:
          ; "In a GetData call, the Data returned is from stream position zero
          ; through just before the current seek pointer of the stream (that is, the position on exit)."
          *StgMed_Dest\pstm\Seek(0, #STREAM_SEEK_END, 0)
        EndIf
      EndIf
      ;
    ElseIf *StgMed_Source\tymed = #TYMED_FILE
      ; Handle file format
      If *StgMed_Source\lpszFileName = 0
        If Method$ = "SetData" Or Method$ = ""
          RetValue = #E_INVALIDARG
        Else
          RetValue = #E_UNEXPECTED
        EndIf
      Else
        ; Duplicate the file path
        *StgMed_Dest\lpszFileName = SysAllocString_(*StgMed_Source\lpszFileName)
        If *StgMed_Dest\lpszFileName
          RetValue = #S_OK
        Else
          RetValue = #E_OUTOFMEMORY
        EndIf
      EndIf
    ElseIf *StgMed_Source\tymed = #Null
      ; Some processes need to register simple formats without data into IDataObject.
      ; For exemple, "EnterpriseDataProtectionId" can be found into data copied from
      ; 'WordPad', wiith a Null Tymed associed.
      RetValue = #S_OK
    Else
      RetValue = DefaultRetValue
    EndIf
  EndIf
  ;
  ;
  If RetValue = #S_OK
    IDO_RegisterResults("dataobject " + Method$ + " for " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed) + " ---> OK")
  ElseIf RetValue = #E_NOTIMPL
    IDO_RegisterResults("dataobject " + Method$ + " for: " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed) + " *** ERROR *** --->  Format is not supported (" + Str(*AskedFormatEtc\cfFormat) + ").")
  ElseIf RetValue = #E_OUTOFMEMORY
    IDO_RegisterResults("dataobject " + Method$ + " for: " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed) + " *** ERROR *** --->  Out of memory.")
  ElseIf RetValue = #E_INVALIDARG
    IDO_RegisterResults("dataobject " + Method$ + " for: " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed) + " *** ERROR *** --->  One or more argument is not valid.")
  ElseIf RetValue = #E_FAIL
    IDO_RegisterResults("dataobject " + Method$ + " for: " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed) + " *** ERROR *** --->  Unexpected error.")
  ElseIf RetValue = #DV_E_TYMED
    If Method$ = "GetDataHere"
      ; The GetDataHere method is not suitable for this type of data.
      IDO_RegisterResults("dataobject GetDataHere : " + GetFormatName(*AskedFormatEtc\cfFormat) + GetTymedName(*AskedFormatEtc\tymed) + ": *** ERROR *** --->  Not allowed for GetDataHere.")
    Else
      IDO_RegisterResults("dataobject " + Method$ + ": " + GetFormatName(*AskedFormatEtc\cfFormat) + GetTymedName(*AskedFormatEtc\tymed) + ": *** ERROR *** --->  Tymed type does't match.")
    EndIf
  Else
    IDO_RegisterResults("dataobject " + Method$ + " for: " + GetFormatName(*AskedFormatEtc\cfFormat) + "/" + GetTymedName(*AskedFormatEtc\tymed) + " *** ERROR *** --->  " + GetWinErrorMessage(RetValue) + ".")
  EndIf
  ;
  ProcedureReturn RetValue
EndProcedure
;
;
; *******************************************************************
;
;-                 The twelve methods of the IDataObject
;
; *******************************************************************
;
; Here finally is the essential: the twelve methods essential for the operation
; of an IDataObject, whose addresses are listed in this object's interface table.
; For more information about the function of each of these methods, you can consult
; Microsoft's documentation on IDataObjects.
;
;
; The first parameter of each of these methods should be of type 'IDataObject'.
; But as explained earlier, the pointers of our IDataObject objects also point
; to the 'IDO_DetailsStruct' structure that contains them.
; Therefore, although the processes calling our methods provide them
; with pointer of type 'IDataObject', we interpret it as a pointer to
; an 'IDO_DetailsStruct' structure, which allows us to access the management data of
; the object.
;
Procedure dataobject_GetData(*DataObjectDet.IDO_DetailsStruct, *pformatetcIn.FORMATETC, *pmedium.STGMEDIUM)
  ; This procedure implements the 'GetData' method for the IDataObject we create in this program.
  ; It supports TYMED_GDI, TYMED_HGLOBAL, TYMED_ENHMF, TYMED_MFPICT, TYMED_FILE, TYMED_ISTORAGE, And TYMED_ISTREAM formats.
  ;
  ; Its role is to fill the '*pmedium.STGMEDIUM' structure passed as a pointer parameter
  ; with the data stored in our IDO_DataSave() list and pointed by *DataObjectDet\m_stgmed[].
  ;
  ; The hBitmap or hGlobal (or other type) content of *DataObjectDet\m_stgmed[] is duplicated,
  ; and a pointer to this duplicated content is placed into *pmedium.STGMEDIUM.
  ;
  ; GetData will work only if the requested data format (*pformatetcIn\cfFormat)
  ; matches with one we get in our IDO_Details() list *DataObjectDet\m_format[].
  ;
  ; This method will be called in two different contexts:
  ; • By the normal way of a call to IDataObject\GetData()
  ; • Indirectly, by a call to IDataObject\GetDataHere() which will call dataobject_GetData().
  ; Before calling this procedure, dataobject_GetDataHere() will set the shared variable 'GetDataHere' to 1.
  ; So examining this variable, we're able to know if we are in a GetData call context or
  ; in GetDataHere call context.
  ;
  ; Anyway, all of the differences between 'GetData' and 'GetDataHere' are managed by the
  ; procedure CloneStgMedium() wich will be called to duplicate the data. 
  ;
  Shared GetDataHere, IDO_Details(), IDO_DataSave()
  Protected Found = 0
  Protected CheckedData
  Protected Method$, DPos, *MyStgMedium.STGMEDIUM, *MyFormatEtc.FORMATETC
  ;
  If GetDataHere = 0
    Method$ = "GetData"
  Else
    Method$ = "GetDataHere"
  EndIf
  ;
  ; Check parameters :
  If IsIDataObject(*DataObjectDet) = 0
    IDO_RegisterResults("dataobject_" + Method$ + " ERROR: THIS IDATAOBJECT IS NOT VALID.")
    ProcedureReturn #E_POINTER
  EndIf
  If *pformatetcIn = 0 Or *pmedium = 0
    IDO_RegisterResults("dataobject " + Method$ + "ERROR: *pformatetcIn or *pmedium is null!!")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  ; Examine the data in our DataObjectDet to see if we have what is requested by the caller:
  Found = -1
  For DPos = 0 To *DataObjectDet\m_count - 1
    If *DataObjectDet\m_stgmed[DPos] = 0 Or *DataObjectDet\m_format[DPos] = 0
      ; This should not happen.
      MessageRequester("Oops!", "dataobject_" + Method$ + "(): Bad pointer error.")
      ProcedureReturn #E_UNEXPECTED
    Else
      *MyStgMedium = *DataObjectDet\m_stgmed[DPos]
      *MyFormatEtc = *DataObjectDet\m_format[DPos]
      ; Check if we still have data in memory:
      CheckedData = 0
      ForEach IDO_DataSave()
        If IDO_DataSave()\DataOwner = *DataObjectDet And IDO_DataSave()\m_format = *DataObjectDet\m_format[DPos] And IDO_DataSave()\m_stgmed = *DataObjectDet\m_stgmed[DPos]
          CheckedData = 1
          Break
        EndIf
      Next
      If CheckedData = 0
        IDO_RegisterResults("dataobject_" + Method$ + ": " + GetFormatName(*pformatetcIn\cfFormat) + "/" + GetTymedName(*pformatetcIn\tymed) + " ---> ERROR: IDataObject is pointing to non-existing data!")
        ProcedureReturn #E_UNEXPECTED
      EndIf
      If *pformatetcIn\tymed & *MyFormatEtc\tymed Or *pformatetcIn\tymed = 0 Or (*pformatetcIn\tymed & #TYMED_ISTREAM And (*MyStgMedium\tymed = #TYMED_HGLOBAL Or *MyStgMedium\tymed = #TYMED_MFPICT))
        ; The AND operator is used here because the FORMATETC\tymed member is actually
        ; a bit-flag which can contain more than one value. For example, the caller
        ; of GetData could quite legitimetly specify a FORMATETC::tymed value of
        ; (TYMED_HGLOBAL | TYMED_ISTREAM), which basically means 
        ; “Do you support HGlobal Or IStream?”.
        ;
        ; We also accept *pformatetcIn\tymed = 0, interpreting this type of request
        ; as: “Do you have any tymed data with the requested format?”.
        ;
        ; For our part, even if the value of tymed is always strictly unique
        ; in '*MyStgMedium', it is possible that it is made up of a combination
        ; such as 'TYMED_HGLOBAL | TYMED_ISTREAM' in '*MyFormatEtc', indicating
        ; to the caller that we are able to produce a different 'Tymed' than
        ; what is in our Data.
        If *pformatetcIn\cfFormat = *MyFormatEtc\cfFormat
          If *pformatetcIn\dwAspect = *MyFormatEtc\dwAspect Or *pformatetcIn\dwAspect = 0
            If *pformatetcIn\tymed & *MyStgMedium\tymed Or *pformatetcIn\tymed = 0
              ; The format requested by the procedure is indeed the one we have in our data.
              Found = DPos
            ElseIf Found = -1 And *MyStgMedium\tymed = #TYMED_HGLOBAL And *pformatetcIn\tymed & #TYMED_ISTREAM
              ; TYMED_ISTREAM is asked. We can generate it if necessary.
              Found = DPos
            EndIf
            ;
          EndIf
        EndIf
      EndIf
    EndIf
  Next
  ;
  If Found <> -1
    *MyStgMedium = *DataObjectDet\m_stgmed[Found]
    *MyFormatEtc = *DataObjectDet\m_format[Found]
    ; Fill *pmedium with a copy of *MyStgMedium or with an 'on the fly' generated IStream:
    ProcedureReturn CloneStgMedium(*DataObjectDet, *MyStgMedium, *pmedium, *pformatetcIn, Method$)
  Else
    IDO_RegisterResults("dataobject_" + Method$ + ": " + GetFormatName(*pformatetcIn\cfFormat) + "/" + GetTymedName(*pformatetcIn\tymed) + " ---> NOT FOUND")
    ProcedureReturn #DV_E_FORMATETC
  EndIf
  ;
EndProcedure
;
Procedure dataobject_EnumFormatEtc(*DataObjectDet.IDO_DetailsStruct, dwDirection , *ppenumFormatEtc.IEnumFORMATETC)
  ; When a process calls EnumFormatEtc(), it expects to retrieve
  ; a new IEnumFORMATETC object, provided with an interface, whose
  ; functions allow browsing through the various formats of the examined
  ; IDataObject.
  ; Rather than fully implementing the interface for this type of object, we
  ; use the SHCreateStdEnumFmtEtc_ function, which takes care of all the work for us.
  ; To be able to create the interface on our behalf, SHCreateStdEnumFmtEtc_
  ; only requires that we provide the list of our formats in the form
  ; of an array.
  ; We will now create this array and pass it as a parameter to SHCreateStdEnumFmtEtc_
  ;
  Protected tx$, *FormatEtcArray, i, RetValue
  ;
  ; Check parameters :
  If IsIDataObject(*DataObjectDet) = 0
    IDO_RegisterResults("dataobject_EnumFormatEtc ERROR: THIS IDATAOBJECT IS NOT VALID.")
    ProcedureReturn #E_POINTER
  EndIf
  ;
  tx$ = "dataobject_EnumFormatEtc called with "
  If dwDirection = #DATADIR_GET : tx$ + "#DATADIR_GET" : EndIf
  If dwDirection = #DATADIR_SET : tx$ + "#DATADIR_SET" : EndIf
  IDO_RegisterResults(tx$)
  If *DataObjectDet\m_count = 0
    IDO_RegisterResults("dataobject_EnumFormatEtc error: No format in the object.")
    ProcedureReturn #E_NOTIMPL
  EndIf
  If dwDirection = #DATADIR_GET
   ; Will store a tab of our formatetc objects into a *ppenumFormatEtc enumeration
   ; ATTENTION : our m_format array is an array of POINTERS to formatetc objects
   ; What we need now is an array of formatetc objects formatted as C language formats arrays.
   ; So, we'll create this array now:
    *FormatEtcArray = AllocateMemory(SizeOf(FormatEtc) * *DataObjectDet\m_count)
    For i = 0 To *DataObjectDet\m_count - 1
      CopyMemory(*DataObjectDet\m_format[i], *FormatEtcArray + (i * SizeOf(FormatEtc)), SizeOf(FormatEtc))
    Next
    RetValue = SHCreateStdEnumFmtEtc_(*DataObjectDet\m_count, *FormatEtcArray, *ppenumFormatEtc)
    FreeMemory(*FormatEtcArray)
    If RetValue = #S_OK
      IDO_RegisterResults("   -> SHCreateStdEnumFmtEtc_ called with success")
    Else
      IDO_RegisterResults("   -> SHCreateStdEnumFmtEtc_ error in dataobject_EnumFormatEtc: " + GetWinErrorMessage(RetValue))
    EndIf
    ProcedureReturn RetValue
  Else ; EnumFormatEtc is not implemented For #DATADIR_SET
    ProcedureReturn #E_NOTIMPL
  EndIf
EndProcedure 
;
Procedure dataobject_SetData(*DataObjectDet.IDO_DetailsStruct, *pformatetc.FORMATETC, *pmedium.STGMEDIUM, fRelease)
  ; This procedure sets data in our IDataObject.
  ; It supports TYMED_GDI, TYMED_HGLOBAL, TYMED_ENHMF, TYMED_MFPICT, TYMED_FILE, TYMED_ISTORAGE, And TYMED_ISTREAM formats.
  ;
  Shared IDO_DataSave()
  ; DataSave is a shared list of structures for storing FORMATETC and STGMEDIUM data.
  Protected tx$, Found, RetValue
  ;
  ; Check parameters :
  If IsIDataObject(*DataObjectDet) = 0
    IDO_RegisterResults("dataobject_SetData ERROR: THIS IDATAOBJECT IS NOT VALID.")
    ProcedureReturn #E_POINTER
  EndIf
  If *pformatetc = 0 Or *pmedium = 0
    IDO_RegisterResults("dataobject SetData: *pformatetc or *pmedium is null!!")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  tx$ = "dataobject_SetData has received: " + GetFormatName(*pformatetc\cfFormat) + " - " + GetTymedName(*pmedium\tymed) + "."
  ;
  If *pmedium\tymed And *pmedium\tymed <> #TYMED_GDI And *pmedium\tymed <> #TYMED_HGLOBAL And *pmedium\tymed <> #TYMED_ENHMF And *pmedium\tymed <> #TYMED_MFPICT And *pmedium\tymed <> #TYMED_FILE And *pmedium\tymed <> #TYMED_ISTORAGE And *pmedium\tymed <> #TYMED_ISTREAM
    tx$ + " / This Tymed is not implemented! (" + GetTymedName(*pmedium\tymed) + ")"
    IDO_RegisterResults(tx$)
    ProcedureReturn #E_NOTIMPL
  EndIf
  IDO_RegisterResults(tx$)
  ;
  ; Before adding a new entry, check if there is already data
  ; with the same format and tymed types as the data we have received:
  Found = 0
  ForEach IDO_DataSave()
    If CompareMemory(@IDO_DataSave()\m_format, *pformatetc, SizeOf(FormatEtc))
      ; If found, release the current StgMedium associated with this entry
      ; to free memory:
      ReleaseStgMedium_(@IDO_DataSave()\m_stgmed)
      Found = 1
      Break
    EndIf
  Next
  ;
  If *DataObjectDet\m_count < #DataMax Or Found
    If Found = 0
      ; Create a new IDO_DataSave() list element to store
      ; the new FormatEtc And StgMedium:
      AddElement(IDO_DataSave())
    EndIf
    ; Copy the FormatEtc into IDO_DataSave()
    CopyMemory(*pformatetc, @IDO_DataSave()\m_format, SizeOf(FormatEtc))
    ; We save the memory address of the new format in the IData
    *DataObjectDet\m_format[*DataObjectDet\m_count] = @IDO_DataSave()\m_format
    ;
    ; Copy the StgMedium into IDO_DataSave()
    CopyMemory(*pmedium, @IDO_DataSave()\m_stgmed, SizeOf(StgMedium))
    ; We save the memory address of the new StgMedium in the IData
    *DataObjectDet\m_stgmed[*DataObjectDet\m_count] = @IDO_DataSave()\m_stgmed
    ;
    ; Save the memory address of the IDataObject in the IData
    IDO_DataSave()\DataOwner = *DataObjectDet
    ;
    ; Duplicate data
    RetValue = CloneStgMedium(*DataObjectDet, *pmedium, @IDO_DataSave()\m_stgmed, *pformatetc, "SetData")
    ;
    ;
    If fRelease = #True
      ; It is important to understand that when a process calls SetData()
      ; with the value '#True' in the 'fRelease' parameter, it expects SetData()
      ; to copy the data sent into the StgMedium, and to free the memory of the
      ; original data by deleting it.
      ; However, there are three exceptions:
      ;
      ; • Files whose names are passed in *pmedium\lpszFileName (#TYMED_FILE)
      ;   are not copied, nor deleted when 'fRelease' = '#True'. Copying would consume
      ;   too much disk space, and deletion would be dangerous due to improper
      ;   handling of interface interactions. File deletion must be managed by the
      ;   process that created or manages the files. With #TYMED_FILE, the only thing that
      ;   is copied and deleted is the BitStream containing the file name and path.
      ;
      ; • IStreams referenced in data of type #TYMED_ISTREAM are neither copied
      ;   nor deleted when fRelease = #True. However, the IStream interface is notified
      ;   that it has an additional user through a QueryInterface call.
      ;
      ; • IStorages are managed exactly like IStreams (QueryInterface).
      ;
      ; The data provided during the call is released using the ReleaseStgMedium() 
      ; function, which will handle, regardless of the data type in the StgMedium, 
      ; calling DeleteObject_(), GlobalFree_(), or SysFreeString_(), as needed.
      ;     
      ReleaseStgMedium_(*pmedium)
      IDO_RegisterResults("   SetData --->  ORIGINAL DATA HAS BEEN RELEASED")
    EndIf
    ;
    If RetValue = #S_OK
      *DataObjectDet\m_count + 1
      ;
      If Found And *DataObjectDet\pDataAdviseHolder
        ; 
        ; We have just updated data that was already present
        ; in our IDataObject. It is possible that some processes
        ; have subscribed to this data using the DAdvise method in order
        ; to be notified of changes. We will check for this
        ; and notify the subscribers if necessary
        ;
        *DataObjectDet\pDataAdviseHolder\SendOnDataChange(*DataObjectDet, 0 , 0)
      EndIf
    Else
      *DataObjectDet\m_format[*DataObjectDet\m_count] = 0
      *DataObjectDet\m_stgmed[*DataObjectDet\m_count] = 0
      DeleteElement(IDO_DataSave())
    EndIf
  Else
    IDO_RegisterResults("   SetData: error ---> LIMIT OF " + Str(#DataMax) + " DATA REACHED")
    RetValue = #E_OUTOFMEMORY
  EndIf
  ;
  ProcedureReturn RetValue
EndProcedure
;
Procedure dataobject_QueryInterface(*pDataObject.IDataObject, iid, *ppvObject.Integer)
  ;
  ; The process calling QueryInterface expects to receive a pointer to the method table
  ; in *ppvObject.
  ;
  ; The call to QueryInterface is also a way to notify the IDataObject
  ; that a process is about to use it and that it must retain its data until
  ; a call to IDataObject\Release() frees it from its obligations.
  ;
  ;
  If IsIDataObject(*pDataObject) = 0 Or *ppvObject = 0
    IDO_RegisterResults("dataobject_QueryInterface ERROR: BAD PARAMETER -> " + Str(*pDataObject) + " / " + Str(*ppvObject))
    ProcedureReturn #E_POINTER
  ElseIf iid = 0
    IDO_RegisterResults("dataobject_QueryInterface: #E_NOINTERFACE -> " + StringFromCLSID(iid))
    ProcedureReturn #E_NOINTERFACE
  EndIf
  If CompareMemory(iid, ?IID_IUnknown, 16) = 1 Or CompareMemory(iid, ?IID_IDataObject, 16) = 1
    IDO_RegisterResults("dataobject_QueryInterface: OK")
    *ppvObject\i = *pDataObject
    *pDataObject\AddRef() ; AddRef() va mémoriser le fait que notre object est utilisé.
    ProcedureReturn #S_OK
  Else
    IDO_RegisterResults("dataobject_QueryInterface: #E_NOINTERFACE -> " + StringFromCLSID(iid))
    ProcedureReturn #E_NOINTERFACE
  EndIf
EndProcedure
;
Procedure dataobject_AddRef(*DataObjectDet.IDO_DetailsStruct)
  ;
  If IsIDataObject(*DataObjectDet) > 0
    *DataObjectDet\Refcount = *DataObjectDet\Refcount + 1
    ;
    IDO_RegisterResults("dataobject_AddRef: Refcount = " + Str(*DataObjectDet\Refcount))
    ProcedureReturn *DataObjectDet\Refcount
  EndIf
EndProcedure
;
Procedure dataobject_Release(*DataObjectDet.IDO_DetailsStruct)
  ;
  Shared IDO_DataSave(), IDO_Details()
  Protected ct, tx$
  ;
  If IsIDataObject(*DataObjectDet) > 0
    *DataObjectDet\Refcount = *DataObjectDet\Refcount - 1
    If *DataObjectDet\Refcount > 0
      IDO_RegisterResults("dataobject_Release: Some instance is still using the object: RefCount = " + Str(*DataObjectDet\Refcount))
      ProcedureReturn *DataObjectDet\Refcount
    Else
      IDO_RegisterResults("dataobject_Release:")
      If *DataObjectDet\pDataAdviseHolder
        *DataObjectDet\pDataAdviseHolder\SendOnDataChange(*DataObjectDet, 0 , 64) ; #ADVF_DATAONSTOP = 64
        *DataObjectDet\pDataAdviseHolder\Release()
      EndIf
      ct  = 0
      ForEach IDO_DataSave()
        ct + 1
        If IDO_DataSave()\DataOwner = *DataObjectDet
          ReleaseStgMedium_(IDO_DataSave()\m_stgmed)
          DeleteElement(IDO_DataSave())
        EndIf
      Next
      tx$ + " --> " + Str(ct) + " IDO_DataSave() element(s) have been released."
      IDO_RegisterResults(" •  --> " + Str(ct) + " IDO_DataSave() element(s) have been released.")
      ForEach IDO_Details()
        If IDO_Details() = *DataObjectDet
          DeleteElement(IDO_Details())
          IDO_RegisterResults(" •  --> The IDO_Details() element corresponding to the IDataObject has been released.")
          Break
        EndIf
      Next
    EndIf
  EndIf
  ProcedureReturn #S_OK
EndProcedure
;
Procedure dataobject_QueryGetData(*DataObjectDet.IDO_DetailsStruct, *pformatetcIn.FormatEtc)
  ;
  ; The process calling QueryGetData wants to know if our object is able
  ; to deliver data in the format defined by *pformatetcIn.
  ;
  ; We will compare *pformatetcIn with the various FormatEtc we have in memory
  ; and respond accordingly.
  ;
  Protected DPos, *MyFormatEtc.FORMATETC, *MyStgMedium.StgMedium
  Protected TymedFound, FormatFound, AspectFound
  ;
  If IsIDataObject(*DataObjectDet) = 0 Or *pformatetcIn = 0
    IDO_RegisterResults("dataobject_QueryGetData error: E_INVALIDARG.")
    ProcedureReturn #E_INVALIDARG
  ElseIf *pformatetcIn\cfFormat = 0
    IDO_RegisterResults("dataobject_QueryGetData error: E_INVALIDARG.")
    ProcedureReturn #E_INVALIDARG
  Else
    For DPos = 0 To *DataObjectDet\m_count - 1
      *MyStgMedium = *DataObjectDet\m_stgmed[DPos]
      *MyFormatEtc = *DataObjectDet\m_format[DPos]
      If *pformatetcIn\tymed = 0 Or (*pformatetcIn\tymed & *MyFormatEtc\tymed) Or (*pformatetcIn\tymed & #TYMED_ISTREAM And (*MyStgMedium\tymed = #TYMED_HGLOBAL Or *MyStgMedium\tymed = #TYMED_MFPICT))
        ; The AND (&) operator is used for (*pformatetcIn\tymed & *MyFormatEtc\tymed)
        ; because the FORMATETC\tymed member is actually
        ; a bit-flag which can contain more than one value.
        ; For example, the caller of QueryGetData could quite
        ; legitimetly specify a FORMATETC\tymed value of
        ; (TYMED_HGLOBAL | TYMED_ISTREAM), which basically means 
        ; “Do you support HGLOBAL Or IStream?”.
        ;
        ; We also accept *pformatetcIn\tymed = 0, interpreting this type of request
        ; as: “Do you have any tymed data with the requested format?”.
        ;
        ; If *pformatetcIn\tymed is containing #TYMED_ISTREAM and we get #TYMED_HGLOBAL or #TYMED_MFPICT,
        ; we reply positively, because we are able to generate on the fly IStream from those tymed types.
        TymedFound = 1
        If *pformatetcIn\cfFormat = *MyFormatEtc\cfFormat
          FormatFound = 1
          If *pformatetcIn\dwAspect = *MyFormatEtc\dwAspect
            AspectFound = 1
          EndIf
        EndIf
      EndIf
    Next
    ;
    If TymedFound = 0
      IDO_RegisterResults("dataobject_QueryGetData: " + GetTymedName(*pformatetcIn\tymed) + "/" + GetFormatName(*pformatetcIn\cfFormat) + ". --> Not FOUND")
      ProcedureReturn #DV_E_TYMED
    ElseIf FormatFound = 0
      IDO_RegisterResults("dataobject_QueryGetData: " + GetFormatName(*pformatetcIn\cfFormat) + ". --> NOT FOUND")
      ProcedureReturn #DV_E_FORMATETC
    ElseIf AspectFound = 0
      IDO_RegisterResults("dataobject_QueryGetData: DVASPECT Value #" + Str(*pformatetcIn\dwAspect) + ". --> NOT FOUND")
      ProcedureReturn #DV_E_DVASPECT
    EndIf
  EndIf
  IDO_RegisterResults("dataobject_QueryGetData with " + GetFormatName(*pformatetcIn\cfFormat) + "/" + GetTymedName(*pformatetcIn\tymed) + " --> OK/Found.")
  ProcedureReturn #S_OK
EndProcedure
;
Procedure dataobject_GetCanonicalFormatEtc(*DataObjectDet.IDO_DetailsStruct, *pformatetcIn.FormatEtc, *pformatetcOut.FormatEtc)
  ;
  ; The process calling GetCanonicalFormatEtc wants to know if our object can
  ; simplify the format described by *pformatetcIn to a canonical version.
  ;
  ; We will search through the list of formats and return a matching format if found.
  ;
  Protected DPos, *MyFormatEtc.FORMATETC
  ;
  IDO_RegisterResults("dataobject_GetCanonicalFormatEtc with " + GetFormatName(*pformatetcIn\cfFormat))
  
  If IsIDataObject(*DataObjectDet) = 0 Or *pformatetcIn = 0 Or *pformatetcOut = 0
    IDO_RegisterResults("dataobject_GetCanonicalFormatEtc error: E_INVALIDARG.")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  ; Initialize output to NULL
  *pformatetcOut\cfFormat = 0
  *pformatetcOut\tymed = 0
  *pformatetcOut\dwAspect = 0
  *pformatetcOut\lindex = -1
  
  ; Search through the list of existing formats
  For DPos = 0 To *DataObjectDet\m_count - 1
    *MyFormatEtc = *DataObjectDet\m_format[DPos]
    If *pformatetcIn\tymed & *MyFormatEtc\tymed Or *pformatetcIn\tymed = 0
      If *pformatetcIn\cfFormat = *MyFormatEtc\cfFormat
        ; Matching format found, copy it to *pformatetcOut
        CopyMemory(@*MyFormatEtc, *pformatetcOut, SizeOf(FormatEtc))
        IDO_RegisterResults("dataobject_GetCanonicalFormatEtc result: Format found and copied.")
        ProcedureReturn #S_OK
      EndIf
    EndIf
  Next
  ;
  ; No matching format found
  IDO_RegisterResults("dataobject_GetCanonicalFormatEtc: No matching format found.")
  ProcedureReturn #DATA_S_SAMEFORMATETC
EndProcedure
;
Procedure dataobject_GetDataHere(*DataObjectDet.IDO_DetailsStruct, *pformatetcIn.FORMATETC, *pmedium.STGMEDIUM)
  Shared GetDataHere
  GetDataHere = 1
  Protected RetValue = dataobject_GetData(*DataObjectDet, *pformatetcIn, *pmedium)
  GetDataHere = 0
  ProcedureReturn RetValue
EndProcedure
;
Procedure dataobject_DAdvise(*DataObjectDet.IDO_DetailsStruct, *pformatetc.FORMATETC, advf, *padvSink.IAdviseSink, *pdwConnection)
  ;
  ; In theory, managing DAdvises requires a new implementation:
  ; that of the IEnumSTATDATA interface which allows us to respond to processes
  ; when they call the 'EnumDAdvise()' method of our IDataObject.
  ; Indeed, each DAdvise received by our object must be stored so
  ; that we can manage it by notifying the calling process of any
  ; potential changes to specific data (see the end of 'SetData()' method code
  ; for a better understanding of this management) or consult Microsoft's
  ; documentation regarding the 'DAdvise()' method.
  ;
  ; To avoid having to fully construct an IEnumSTATDATA object,
  ; we use CreateDataAdviseHolder_() which will handle the task for
  ; us. This ensures that the management of DAdvises is done
  ; correctly.
  ;
  Protected RetValue
  ;
  ; Check parameters:
  If IsIDataObject(*DataObjectDet) = 0 Or *pformatetc = 0 Or *padvSink = 0 Or *pdwConnection = 0
    IDO_RegisterResults("dataobject_DAdvise error: E_INVALIDARG.")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  ; If the IDataAdviseHolder is not yet created, do so now
  If *DataObjectDet\pDataAdviseHolder = #Null
    If CreateDataAdviseHolder_(@*DataObjectDet\pDataAdviseHolder) <> #S_OK
      IDO_RegisterResults("dataobject_DAdvise error: Failed to create IDataAdviseHolder.")
      ProcedureReturn #E_FAIL
    EndIf
  EndIf
  ;
  ; Check if we get data with the specified format:
  If dataobject_QueryGetData(*DataObjectDet, *pformatetc) = #S_OK
    ; Delegate the DAdvise call to IDataAdviseHolder
    RetValue = *DataObjectDet\pDataAdviseHolder\Advise(*DataObjectDet, *pformatetc, advf, *padvSink, *pdwConnection)
    ;
    If RetValue = #S_OK
      IDO_RegisterResults("dataobject_DAdvise: Connection ID " + Str(PeekW(*pdwConnection)) + " successfully registered for " + GetTymedName(*pformatetc\tymed) + "/" + GetFormatName(*pformatetc\cfFormat) + ".")
    Else
      IDO_RegisterResults("dataobject_DAdvise error: Unable to register the advise sink: " + GetWinErrorMessage(RetValue) + ".")
    EndIf
  Else
    IDO_RegisterResults("dataobject_DAdvise error: Requested format (" + GetTymedName(*pformatetc\tymed) + "/" + GetFormatName(*pformatetc\cfFormat) + ") is not in our data.")
    RetValue = #OLE_E_ADVISENOTSUPPORTED
  EndIf
  ;
  ProcedureReturn RetValue
EndProcedure
;
Procedure dataobject_DUnAdvise(*DataObjectDet.IDO_DetailsStruct, dwConnection.l)
  ;
  Protected RetValue
  ;
  If IsIDataObject(*DataObjectDet) = 0 Or dwConnection = 0
    IDO_RegisterResults("dataobject_DUnAdvise error: E_INVALIDARG.")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  ; If there's no IDataAdviseHolder, there's nothing to unregister
  If *DataObjectDet\pDataAdviseHolder = #Null
    IDO_RegisterResults("dataobject_DUnAdvise error: IDataAdviseHolder not initialized.")
    ProcedureReturn #E_FAIL
  EndIf
  ;
  ; Delegate the DUnAdvise call to IDataAdviseHolder
  RetValue = *DataObjectDet\pDataAdviseHolder\Unadvise(dwConnection)
  ;
  If RetValue = #S_OK
    IDO_RegisterResults("dataobject_DUnAdvise: Connection ID " + Str(dwConnection) + " successfully unregistered.")
  Else
    IDO_RegisterResults("dataobject_DUnAdvise error: Unable to unregister the advise sink: " + GetWinErrorMessage(RetValue) + ".")
  EndIf
  ;
  ProcedureReturn RetValue
EndProcedure
;
Procedure dataobject_EnumDAdvise(*DataObjectDet.IDO_DetailsStruct, *ppenumAdvise.IEnumSTATDATA)
  ;
  ; See 'dataobject_DAdvise()' to get explanations about the DAvise system.
  ;
  Protected SC
  ;
  If IsIDataObject(*DataObjectDet) = 0 Or *ppenumAdvise = 0
    IDO_RegisterResults("dataobject_EnumDAdvise error: E_INVALIDARG.")
    ProcedureReturn #E_INVALIDARG
  EndIf
  ;
  ; If the IDataAdviseHolder is not yet created, do so now
  If *DataObjectDet\pDataAdviseHolder = #Null
    SC = CreateDataAdviseHolder_(@*DataObjectDet\pDataAdviseHolder)
    If SC <> #S_OK
      IDO_RegisterResults("dataobject_DAdvise error: Failed to create IDataAdviseHolder: " + GetWinErrorMessage(SC) + ".")
      ProcedureReturn #E_FAIL
    EndIf
  EndIf
  ;
  ; Use the EnumAdvise method from the IDataAdviseHolder to get the IEnumSTATDATA
  SC = *DataObjectDet\pDataAdviseHolder\EnumAdvise(*ppenumAdvise)
  If SC <> #S_OK
    IDO_RegisterResults("dataobject_EnumDAdvise error: EnumAdvise failed: " + GetWinErrorMessage(SC) + ".")
    ProcedureReturn #E_FAIL
  EndIf

  IDO_RegisterResults("dataobject_EnumDAdvise: Successfully created IEnumSTATDATA.")
  ProcedureReturn #S_OK
EndProcedure
;
;
DataSection
  VTable_IDataObject: 
    Data.i @dataobject_QueryInterface(), @dataobject_AddRef(), @dataobject_Release()
    Data.i @dataobject_GetData(), @dataobject_GetDataHere(), @dataobject_QueryGetData()
    Data.i @dataobject_GetCanonicalFormatEtc(), @dataobject_SetData(), @dataobject_EnumFormatEtc()
    Data.i @dataobject_DAdvise(), @dataobject_DUnadvise(), @dataobject_EnumDAdvise()
    ;
  IID_IOleObject:   ;"{00000112-0000-0000-C000-000000000046}"
    Data.l $00000112
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
    ;     
  CLSID_DataObject: ;{00000320-0000-0000-C000-000000000046}
    Data.l $00000320
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
    ;
  IID_IDataObject: 
    Data.l $0000010e
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
    ;
  IID_IUnknown:   ;"{00000000-0000-0000-C000-000000000046}"
    Data.l $00000000
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
    ;
  IID_IStream: 
    Data.l $0000000c
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
    ;
  IID_IStorage: 
    Data.l $0000000b
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46
    ;
EndDataSection
; IDE Options = PureBasic 6.10 LTS (Windows - x86)
; CursorPosition = 499
; FirstLine = 343
; Folding = v5PQ-
; EnableXP
; DPIAware