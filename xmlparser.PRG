* XMLPARSER.PRG
* Parser XML para VFP (6.0 o superior)
*
* Autor: Victor Espina
* Version: 1.3
* Fecha: Jul 2023
*
* USO:
* DO xmlparser
*
* archivo.xml
* <documento>
*    <cliente codigo="001"  nombre="VICTOR ESPINA" />
*    <facturas>
*       <factura>
*          <numero>02002</numero>
*          <fecha>01/01/2019</fecha>
*          <monto>23.25</monto>
*       </factura>
*       <factura>
*          <numero>02003</numero>
*          <fecha>01/02/2019</fecha>
*          <monto>50.00</monto>
*       </factura>
*    </facturas>
* </documento>
*
* oXML = XMLParser.Parse("archivo.xml")
* IF ISNULL(oXML)
*   MESSAGEBOX(oXMLParser.lastError)
*   RETURN
* ENDIF
*
* ?oXML.cliente.codigo --> "001"
* ?oXML.facturas.count --> 2
* oFactura = oXML.Facturas.Items[1]
* ?oFactura.numero  -->  "02002"
*
*
* CHANGE HISTORY
* Jul 8, 2023  VES     Se integro VFPLEGACY.PRG dentro del PRG principal
* Dic 7, 2022  VES     Correccion para caso de nodos con prefijo
* Ago 9, 2022  VES     Mejoras varias
* Jun 2, 2019  VES     Version inicial
*

* Load XMLParser
SET PROC TO xmlParser ADDITIVE

PUBLIC xmlParser
xmlParser = CREATE("xmlParser")


RETURN 

DEFINE CLASS xmlParser AS Custom

     ******************************************
     **          P R O P I E D A D E S       **
     ******************************************
     lastError = ""
   
   
   
     ******************************************
     **             M E T O D O S            **
     ******************************************
   
     PROCEDURE Parse(pcXMLFile)
       LOCAL oXML,oData
       oXML = CREATEOBJECT('MSXML2.DOMdocument')
       oData = NULL
       pcXMLFile = FULLPATH(pcXMLFile)
       IF !FILE(pcXMLFile)
         THIS.lastError = "File [" + LOWER(pcXMLFile) + "] does not exists"
         RETURN NULL
       ENDIF
       oXML.Load( FULLPATH(pcXMLFile) )
       IF oXML.parseError.errorCode <> 0
         THIS.lastError = "Bad formed: " + oXML.parseError.reason
         RETURN NULL
       ENDIF
   
       LOCAL oRootNode,cRootTagName,oMainNode
       oRootNode = oXML.documentElement
       cRootTagName = oRootNode.tagName
       oMainNode = oXML.selectSingleNode("//"+cRootTagName)
  
       oData = THIS._parseNode(oMainNode)
   
       RETURN oData
     ENDPROC


     PROCEDURE parseString(pcString)
     	LOCAL cTempFile,oResult
     	cTempFile = ADDBS(GETENV("TEMP")) + "$" + SYS(2015) + ".XML"  
     	STRTOFILE(pcString, cTempFile)
     	oResult = THIS.Parse(cTempFile)
     	ERASE (cTempFile)
     	RETURN oResult
     	
   
     HIDDEN PROCEDURE _parseNode(poNode)
        LOCAL oData,cNodeName,i,oNode,oAttr,lHasChildren,cAttrName,cNodeName
        cNodeName = LOWER(poNode.nodeName)
        lHasChildren = (TYPE("poNode.childNodes.Length") = "N" AND poNode.childNodes.Length > 0)

   
        * Determinamos si es una coleccion. Es importante hacerlo aqui porque
        * si es una coleccion no se leeran los atributos que pueda tener definidos
        LOCAL lIsCollection,oCollection
        lIsCollection = .F.
   
        IF lHasChildren
          * Asumimos que es una coleccion a menos que los hijos tengan distinto nombre
          * o el nombre del unico hijo no sea el plural del padre
          lIsCollection = .T.
          LOCAL cChildrenName
          cChildrenName = ""    
          FOR i = 1 TO poNode.childNodes.Length
            oNode = poNode.childNodes.Item(i - 1)
            IF oNode.nodeType = 8  && Comentario
              LOOP
            ENDIF
            IF EMPTY(cChildrenName)
            	cChildrenName = LOWER(oNode.nodeName)
            ELSE
            	IF NOT LOWER(oNode.nodeName) == cChildrenName
            		cChildrenName = ""
            		EXIT
            	ENDIF
            ENDIF
          ENDFOR
          IF EMPTY(cChildrenName) OR (poNode.childNodes.Length == 1 AND !(cChildrenName + "s" == cNodeName))
              lIsCollection = .F.
          ENDIF          
        ENDIF
        oData = IIF(lISCollection, CREATEOBJECT("Collection"), THIS._createEmpty())
        IF lIsCollection
          oCollection = oData
        ENDIF
   
   
        * Si no es una coleccion, leemos los atributos que se puedan haber definido
        IF !lISCollection AND TYPE("poNode.Attributes.Length")="N"
          FOR i = 1 TO poNode.Attributes.Length
           oAttr = poNode.Attributes.Item(i - 1)
           cAttrName = LOWER(CHRT(oAttr.nodeName,":","_"))           
           IF !INLIST(cAttrNAme, "#text")
             THIS._ADDPROPERTY(@oData, cAttrName, oAttr.nodeValue)
           ENDIF
          ENDFOR
        ENDIF
   
 		IF cNodeName == "div"
 			*SET STEP ON 
 		ENDIF
 		  
        * Si el nodo tiene hijos, se procesan
        LOCAL lHasAttr
        IF lHasChildren
          FOR i = 1 TO poNode.ChildNodes.Length
            oNode = poNode.childNodes.Item(i - 1)
            IF oNode.nodeType = 8  && Comentario
              LOOP
            ENDIF
            oValue = NULL
            cNodeName = LOWER(oNode.nodeName)
            lHasAttr = (!ISNULL(oNode.Attributes) AND oNode.Attributes.Length > 0)
            DO CASE
               CASE oNode.childNodes.Length = 1 AND oNode.childNodes.Item[0].nodeType = 3 AND !lHasAttr
                    oValue = oNode.childNodes.Item[0].text     
                    cNodeName = "text"               

               CASE oNode.childNodes.Length = 1 AND oNode.childNodes.Item[0].nodeType = 4 AND !lHasAttr
                    oValue = oNode.childNodes.Item[0].text     
                    cNodeName = "cdata" 
                                   
               CASE oNode.childNodes.Length = 0 AND !lHasAttr
                    oValue = oNode.Text
                    cNodeName = "text"
   
   	           CASE !lIsCollection AND TYPE("oDate." + oNode.nodeName)<>"U"
   	                lIsCollection = .T.
         			oCollection = CREATEOBJECT("Collection")
         			oCollection.Add(GETPEM(oData, oNode.nodename))
         			STORE oCollection TO ("oData." + oNode.nodeName)
         			   	                
               OTHERWISE
                    oValue = THIS._parseNode(oNode)
            ENDCASE
            IF lIsCollection
              oCollection.Add(oValue)
            ELSE
              THIS._ADDPROPERTY(@oData, cNodeName, oValue)
            ENDIF
          ENDFOR
        ENDIF
   
        RETURN oData
     ENDPROC
   
   
     HIDDEN PROCEDURE _createEmpty(pcColList)
       #IF VERSION(5) >= 900
       	   LOCAL oTarget,i,cProp
       	   pcColList = EVL(pcColList,"")
       	   oTarget = CREATEOBJECT("Empty")
       	   FOR i = 1 TO GETWORDCOUNT(pcColList,",")
       	   	   cProp = GETWORDNUM(pcColList,i,",")
       	   	   ADDPROPERTY(oTarget,cProp,.F.)
       	   ENDFOR
       	   RETURN oTarget
	   #ELSE
	       LOCAL nWkArea,oEmpty,cSQL
	       nWkArea = SELECT()
	       IF EMPTY(pcColList)
	         pcColList = "name"
	       ENDIF
	       cSQL = "CREATE CURSOR QEMPTY (" + STRT(LOWER(pcColList),[,],[ L,]) + " L)"
	       SELECT 0
	       &cSQL
	       APPEND BLANK
	       SCATTER NAME oEmpty
	       USE
	       SELECT (nWkArea)
	   #ENDIF
       RETURN oEmpty
     ENDPROC
   
   
     HIDDEN PROCEDURE _addProperty(poTarget, pcProperty, puValue)
       LOCAL cType
       pcProperty = CHRTRAN(CHRT(pcProperty,":-","__"), "#", "")
       cType = TYPE("poTarget." + pcProperty)
       IF cType<>"U"
          IF cType = "C"
          	poTarget.Text = poTarget.Text + puValue
          ELSE
       	  	STORE puValue TO ("poTarget." + pcProperty)
       	  ENDIF
       	  RETURN poTarget
       ENDIF
       #IF VERSION(5) >= 800
       	   ADDPROPERTY(poTarget, pcPRoperty, puValue)
       #ELSE
	       LOCAL nWkArea
	       nWkArea = SELECT()
	       LOCAL ARRAY aProps[1]
	       LOCAL nPropCount,cProp,i,cPropList
	       cPropList = ""
	       nPropCount = AMEMBERS(aProps, poTarget)
	       FOR i = 1 TO nPropCount
	        cProp = LOWER(aProps[i])
	        cProp = CHRT(cProp,":","_")
	        IF !INLIST(cProp, "#text") 
	          cPropList = cPropList + IIF(i=1,"",",") + cProp
	        ENDIF
	       ENDFOR
	       cPropList = cPropList + "," + LOWER(pcProperty)
	       LOCAL oClone
	       oClone = THIS._createEmpty(cPropList)
	       FOR i = 1 TO nPropCount
	         cProp = aProps[i]
	         cProp = CHRT(cProp,":","_")
	         STORE GETPEM(poTarget, cProp) TO ("oClone." + cProp)
	       ENDFOR
	       STORE puValue TO ("oClone." + pcProperty)
	       poTarget = oClone
	       SELECT (nWkArea)
	   #ENDIF
       RETURN poTarget
     ENDPROC
ENDDEFINE         

   
   
   
* VFPLEGACY.PRG
*
* ADD SUPPORT FOR FEATURES NOT PRESENT IN OLDER
* VERSIONS OF VFP
*
* AUTHOR: VICTOR ESPINA
*
* FEATURES IMPLEMENTED:
*
* A) EMPTY CLASS
* B) TRY-CATCH
* C) Exception CLASS
* D) Collection CLASS
* E) ADDPROPERTY FUNCTION
* F) EVL FUNCTION
*

******************************************************
**
**               VFP 6 SUPPORT
**
******************************************************
#IF VERSION(5) < 800

	* EMPTY
	* Empty class
	*
	DEFINE CLASS EmptyObject AS Line
	ENDDEFINE


	* TRYCATCH.PRG
	* Funciones para la implementacion de bloques TRY-CATCH en versiones
	* de VFP anteriores a 8.00
	*
	* Autor: Victor Espina
	* Fecha: May 2014
	*
	*
	* Uso:
	*
	* LOCAL ex
	* TRY()
	*   un comando
	*   IF NOEX()
	*    otro comando
	*   ENDIF
	*   IF NOEX()
	*    otro comando
	*   ENDIF
	*
	* IF CATCH(@ex)
	*   manejo de error
	* ENDIF
	*
	* ENDTRY()
	*
	*
	* Ejemplo:
	*
	* lOk = .F.
	* TRY()
	*   Iniciar()
	*   IF NOEX()
	*    Terminar()
	*   ENDIF
	*   lOk = NOEX()
	*
	* IF CATCH(@ex)
	*    MESSAGEBOX(ex.Message)
	* ENDIF
	* ENTRY()
	*
	* IF lok
	*  ...
	* ENDIF
	*

	PROCEDURE TRY
	 IF VARTYPE(gcTRYOnError)="U"
	  PUBLIC gcTRYOnError,goTRYEx,gnTRYNestingLevel
	  gnTRYNestingLevel = 0
	 ENDIF
	 goTRYEx = NULL
	 gnTRYNestingLevel = gnTRYNestingLevel + 1
	 IF gnTRYNestingLevel = 1
	  gcTRYOnError = ON("ERROR")
	  ON ERROR tryCatch(ERROR(), MESSAGE(), MESSAGE(1), PROGRAM(), LINENO())
	 ENDIF
	ENDPROC


	PROCEDURE CATCH(poEx)
	 IF PCOUNT() = 1 AND !ISNULL(goTRYEx)
	  poEx = goTRYEx.Clone()
	 ENDIF
	 LOCAL lEx
	 lEx = !ISNULL(goTRYEx)
	 ENDTRY()
	 RETURN lEx
	ENDPROC

	PROCEDURE ENDTRY
	 gnTRYNestingLevel = gnTRYNestingLevel - 1
	 goTRYEx = NULL
	 IF gnTRYNestingLevel = 0 
	  IF !EMPTY(gcTRYOnError)
	   ON ERROR &gcTRYOnError
	  ELSE
	   ON ERROR
	  ENDIF
	 ENDIF
	ENDPROC

	FUNCTION NOEX()
	 RETURN ISNULL(goTRYEx)
	ENDFUNC

	FUNCTION THROW(pcError)
	 ERROR (pcError)
	ENDFUNC

	PROCEDURE tryCatch(pnErrorNo, pcMessage, pcSource, pcProcedure, pnLineNo)
	 goTRYEx = CREATE("_Exception")
	 WITH goTRYEx
	  .errorNo = pnErrorNo
	  .Message = pcMessage
	  .Source = pcSource
	  .Procedure = pcProcedure
	  .lineNo = pnLineNo
	  .lineContents = pcSource
	 ENDWITH
	ENDPROC

	DEFINE CLASS _Exception AS Custom
	 errorNo = 0
	 Message = ""
	 Source = ""
	 Procedure = ""
	 lineNo = 0 
	 Details = ""
	 userValue = ""
	 stackLevel = 0
	 lineContents = ""
	 

	 PROCEDURE Clone
	  LOCAL oEx 
	  oEx = CREATEOBJECT(THIS.Class)
	  oEx.errorNo = THIS.errorNo
	  oEx.MEssage = THIS.Message
	  oEx.Source = THIS.Source
	  oEx.Procedure = THIS.Procedure
	  oEx.lineNo = THIS.lineNo
	  oEx.Details = THIS.Details
	  oEx.stackLevel = THIS.stackLevel
	  oEx.userValue = THIS.userValue
	  oEx.lineContents = THIS.lineContents
	  RETURN oEx
	 ENDPROC
	ENDDEFINE


	* Collection (Class)
	* Implementacion aproximada de la clase Collection de VFP8+
	*
	* Autor: Victor Espina
	* Fecha: Octubre 2012
	*
	DEFINE CLASS Collection AS Custom

	 DIMEN Keys[1]
	 DIMEN Items[1]
	 DIMEN Item[1]
	 Count = 0
	 
	 PROCEDURE Init(pnCapacity)
	  IF PCOUNT() = 0
	   pnCapacity = 0
	  ENDIF
	  DIMEN THIS.Items[MAX(1,pnCapacity)]
	  DIMEN THIS.Keys[MAX(1,pnCapacity)]
	  THIS.Count = pnCapacity
	 ENDPROC
	  
	 PROCEDURE Items_Access(nIndex1,nIndex2)
	  IF VARTYPE(nIndex1) = "N"
	   RETURN THIS.Items[nIndex1]
	  ENDIF
	  LOCAL i
	  FOR i = 1 TO THIS.Count
	   IF THIS.Keys[i] == nIndex1
	    RETURN THIS.Items[i]
	   ENDIF
	  ENDFOR
	 ENDPROC

	 PROCEDURE Items_Assign(cNewVal,nIndex1,nIndex2)
	  IF VARTYPE(nIndex1) = "N"
	   THIS.Items[nIndex1] = m.cNewVal
	  ELSE
	   LOCAL i
	   FOR i = 1 TO THIS.Count
	    IF THIS.Keys[i] == nIndex1
	     THIS.Items[i] = m.cNewVal
	     EXIT
	    ENDIF
	   ENDFOR
	  ENDIF 
	 ENDPROC
	 
	 PROCEDURE Item_Access(nIndex1, nIndex2)
	  RETURN THIS.Items[nIndex1]
	 ENDPROC
	 
	 PROCEDURE Item_Assign(cNewVal, nIndex1, nIndex2)
	  THIS.Items[nIndex1] = cNewVal
	 ENDPROC


	 PROCEDURE Clear
	  DIMEN THIS.Items[1]
	  DIMEN THIS.Keys[1]
	  THIS.Count = 0
	 ENDPROC
	 
	 PROCEDURE Add(puValue, pcKey)
	  IF !EMPTY(pcKey) AND THIS.getKey(pcKey) > 0
	   RETURN .F.
	  ENDIF
	  THIS.Count = THIS.Count + 1
	  IF ALEN(THIS.Items,1) < THIS.Count
	   DIMEN THIS.Items[THIS.Count]
	   DIMEN THIS.Keys[THIS.Count]
	  ENDIF
	  THIS.Items[THIS.Count] = puValue
	  THIS.Keys[THIS.Count] = IIF(EMPTY(pcKey),"",pcKey)
	 ENDPROC
	 
	 PROCEDURE Remove(puKeyOrIndex)
	  IF VARTYPE(puKeyOrIndex)="C"
	   puKeyOrIndex = THIS.getKey(puKeyOrIndex)
	  ENDIF
	  LOCAL i
	  FOR i = puKeyOrIndex TO THIS.Count - 1
	   THIS.Items[i] = THIS.Items[i + 1]
	   THIS.Keys[i] = THIS.Keys[i + 1]
	  ENDFOR
	  THIS.Items[THIS.Count] = NULL
	  THIS.Keys[THIS.Count] = NULL
	  THIS.Count = THIS.Count - 1
	 ENDPROC

	 PROCEDURE getKey(puKeyOrIndex)
	  LOCAL i,uResult
	  IF VARTYPE(puKeyOrIndex)="N"
	   uResult = THIS.Keys[puKeyOrIndex]
	  ELSE
	   uResult = 0
	   FOR i = 1 TO THIS.Count
	    IF THIS.Keys[i] == puKeyOrIndex
	     uResult = i
	     EXIT
	    ENDIF
	   ENDFOR
	  ENDIF
	  RETURN uResult  
	 ENDPROC

	ENDDEFINE


	* ADDPROPERTY
	* Simula la funcion ADDPROPERTY existente en VFP9
	*
	PROCEDURE AddProperty(poObject, pcProperty, puValue)
	 poObject.addProperty(pcProperty, puValue)
	ENDPROC

	* EVL
	* Simula la funcion EVL de VFP9
	*
	FUNCTION EVL(puValue, puDefault)
	 RETURN IIF(EMPTY(puValue), puDefault, puValue)
	ENDFUNC

	#ENDIF

	#IF VERSION(5) > 600
	FUNCTION NOEX
	 RETURN .T.
	ENDFUNC
#ENDIF
   