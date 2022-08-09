* NQINCLUDE.PRG
*
* LIBRARY TO INCLUDE VARIOUS PUBLIC LIBRARIES
*
* AUTHOR: VICTOR ESPINA
* VERSION: 1.1
*
* USAGE:
* SET PROC TO NQInclude
*
* bool = NQInclude("resourceId")
*
*
*
* AVAILABLE RESOURCES:
*
* json				JSON parser
* base64Helper		Base64 encoder/decoder
* vfplegacy			Support for legacy versions of VFP
* mail			    Helper library to send emails from VFP
* toolbox           Support library with multiples useful functions
*

#DEFINE NQ_VERSION 	"1.1"


	
PROCEDURE NQInclude(pcResourceID, plAutoLoad)
	LOCAL cUrls,cCRLF
	cUrls = ""
	cCRLF = CHR(13)+CHR(10)
	IF PCOUNT() = 1
		plAutoLoad = .T.
	ENDIF
		
	IF _VFP.StartMode <> 0 
		MESSAGEBOX("This program must be called from the VFP's IDE.",16,"NQINCLUDE.PRG")
		CANCEL
	ENDIF
		
	pcResourceID = LOWER(pcResourceID)	
	DO CASE
	   CASE pcResourceID == "@@version"
	   	    RETURN NQ_VERSION
	   	    
	   CASE pcResourceId == "vfplegacy"
	        cURLs = "https://raw.githubusercontent.com/vespina/vfplegacy/main/vfplegacy.h" + cCRLF + ;
	                "https://raw.githubusercontent.com/vespina/vfplegacy/main/vfplegacy.prg"
	        
	   CASE pcResourceId == "base64helper"
	        cURLs = "https://raw.githubusercontent.com/vespina/base64helper/main/base64helper.PRG"

	   CASE pcResourceId = "json"
	        cURLs = "https://raw.githubusercontent.com/vespina/json/main/json.PRG"
    
	   CASE pcResourceId = "mail"
	        cURLs = "https://raw.githubusercontent.com/vespina/mail/main/mail.PRG"

	   CASE pcResourceId = "toolbox"
	        cURLs = "https://raw.githubusercontent.com/vespina/toolbox/main/toolbox.PRG"
    
       OTHERWISE
			MESSAGEBOX("Invalid resource id '" + pcResourceID + "'", 0 + 48, "NQInclude")
			RETURN .F.
	ENDCASE
	
	LOCAL ARRAY aSources[1]
	LOCAL nCount,i,cLibrary,cFileName,cUrl,cSource,lResult,lDownloaded
	nCount = ALINES(aSources, cURLs)
	lResult = .T.
	lDownloaded = .F.
	FOR i = 1 TO nCount
		cUrl = aSources[i]
		cLibrary = JUSTSTEM(cUrl)
		cFileName = JUSTFNAME(cUrl)
		IF !FILE(cFileName)
			cSource = GetURL(cUrl)
			IF EMPTY(cSource)
				MESSAGEBOX("Resource '" + cUrl + "' is not available at this time", 0 + 48, "NQInclude")
				lResult = .F.
				EXIT
			ENDIF
			STRTOFILE(cSource, cFileName)
			IF LOWER(JUSTEXT(cFileName)) == "prg"
				COMPILE (cFileName)
			ENDIF
		    lDownloaded = .F.
		ENDIF
		
		IF plAutoLoad AND UPPER(JUSTEXT(cFileName)) == "PRG" AND ATC(cLibrary, SET("PROCEDURE")) = 0 
			SET PROCEDURE TO (cLibrary) ADDITIVE
		ENDIF
	ENDFOR
		
	DO CASE
	   CASE pcResourceId == "vfplegacy" AND lDownloaded AND lResult
   	        MESSAGEBOX("The library VFPLEGACY.PRG has been automatically downloaded. Please recompile all your PRG files that depends on this library.",48,"NQINCLUDE.PRG")
	        CANCEL
	ENDCASE
	RETURN lResult



PROCEDURE NQRequire(pcResourceID, plAutoLoad)
	IF PCOUNT() < 2
		plAutoLoad = .T.
	ENDIF
	IF NOT NQInclude(pcResourceId, plAutoLoad)
		CANCEL
	ENDIF
	RETURN



***********************************************
** GETURL.PRG
** Devuelve el contenido de un URL dado.
**
** Versi?1.0
**
** Autor: Victor Espina (vespinas@cantv.net)
**        Walter Valle (wvalle@develcomp.com)
**        (basado en c?o original de Pablo Almunia)
*
** Fecha: 20-Agosto-2003
**
**
** Sint?s:
** cData = GetURL(pcURL[,plVerbose])
**
** Donde:
** cData	 Contenido (texto o binario) del recurso 
**			 indicado en cURL. Si ocurre alg?ror
**           se devolver?a cadena vacia.
** pcURL	 Direcci?RL del recurso o archivo a obtener
** plVerbose Opcional. Si se establece en True, se mostrar?*		     informaci?isual sobre el avance del proceso.
**
** Ejemplo:
** cHTML=GetURL("http://www.portalfox.com")
**
*************************************************
**
** GETURL.PRG
** Returns the contains of any given URL
**
** Version: 1.0
**
** Author: Victor Espina (vespinas@cantv.net)
**         Walter Valle (wvalle@develcomp.com)
**         (based on original source code from Pablo Almunia)
*
** Date: August 20, 2003
**
**
** Syntax:
** cData = GetURL(pcURL[,plVerbose])
**
** Where:
** cData	 Contents (text or binary) of requested URL.
** pcURL	 URL of the requested resource or file. If an
**           error occurs, a empty string will be returned.
** plVerbose Optional. If setted to True, progress info
**			 will be shown.
**
** Example:
** cHTML=GetURL("http://www.portalfox.com")
**
**************************************************
PROCEDURE GetURL
LPARAMETER pcURL,plVerbose
 *
 *-- Se definen las funciones API necesarias
 *
 #DEFINE INTERNET_OPEN_TYPE_PRECONFIG     0
 DECLARE LONG GetLastError IN WIN32API
 DECLARE INTEGER InternetCloseHandle IN "wininet.dll" ;
	LONG hInet
 DECLARE LONG InternetOpen IN "wininet.dll" ;
  STRING   lpszAgent, ;
  LONG     dwAccessType, ;
  STRING   lpszProxyName, ;
  STRING   lpszProxyBypass, ;
  LONG     dwFlags
 DECLARE LONG InternetOpenUrl IN "wininet.dll" ;
    LONG    hInet, ;
 	STRING  lpszUrl, ;
	STRING  lpszHeaders, ;
    LONG    dwHeadersLength, ;
    LONG    dwFlags, ;
    LONG    dwContext
 DECLARE LONG InternetReadFile IN "wininet.dll" ;
	LONG     hFtpSession, ;
	STRING  @lpBuffer, ;
	LONG     dwNumberOfBytesToRead, ;
	LONG    @lpNumberOfBytesRead
	
	
 *-- Se establece la conexi?on internet
 *
 IF plVerbose
  WAIT "Opening Internet connection..." WINDOW NOWAIT
 ENDIF
 
 LOCAL nInetHnd
 nInetHnd = InternetOpen("GETURL",INTERNET_OPEN_TYPE_PRECONFIG,"","",0)
 IF nInetHnd = 0
  WAIT "ERROR: Connection could not be opened" WINDOW NOWAIT
  RETURN ""
 ENDIF
 
 
 *-- Se establece la conexi?on el recurso
 *
 IF plVerbose
  WAIT "Opening connection to URL..." WINDOW NOWAIT
 ENDIF
 
 LOCAL nURLHnd
 nURLHnd = InternetOpenUrl(nInetHnd,pcURL,NULL,0,0,0)
 IF nURLHnd = 0
  InternetCloseHandle( nInetHnd )
  WAIT "ERROR: URL could not be opened" WINDOW NOWAIT
  RETURN ""
 ENDIF


 *-- Se lee el contenido del recurso
 *
 LOCAL cURLData,cBuffer,nBytesReceived,nBufferSize
 cURLData=""
 cBuffer=""
 nBytesReceived=0
 nBufferSize=0

 DO WHILE .T.
  *
  *-- Se inicializa el buffer de lectura (bloques de 2 Kb)
  cBuffer=REPLICATE(CHR(0),2048)
  
  *-- Se lee el siguiente bloque
  InternetReadFile(nURLHnd,@cBuffer,LEN(cBuffer),@nBufferSize)
  IF nBufferSize = 0
   EXIT
  ENDIF
  
  *-- Se acumula el bloque en el buffer de datos
  cURLData=cURLData + SUBSTR(cBuffer,1,nBufferSize)
  nBytesReceived=nBytesReceived + nBufferSize
  
  IF plVerbose
   WAIT WINDOW ALLTRIM(TRANSFORM(INT(nBytesReceived / 1024),"999,999")) + " Kb received..." NOWAIT
  ENDIF
  *
 ENDDO
 IF plVerbose
  WAIT CLEAR
 ENDIF

 
 *-- Se cierra la conexi? Internet
 *
 InternetCloseHandle( nInetHnd )

 *-- Se devuelve el contenido del URL
 *
 RETURN cURLData
 *
ENDPROC	