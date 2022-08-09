* XMLPARSER.PRG
* Parser XML para VFP (6.0 o superior)
*
* Autor: Victor Espina
* Version: 1.1
* Fecha: Jun 2019
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
* Ago 9, 2022  VES     Mejoras varias
* Jun 2, 2019  VES     Version inicial
*

* Check for NQInclude utility (it will be downloaded if necessary).
checkForNQInclude()

* Check for dependencies
SET PROCEDURE TO NQInclude ADDITIVE
NQRequire("vfplegacy") 

#IF VERSION(5) < 800
	SET PROCEDURE TO VFPLEGACY ADDITIVE
#ENDIF

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
	       pcProperty = CHRTRAN(CHRT(pcProperty,":-","__"), "#", "")
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





	
PROCEDURE checkForNQInclude
   IF NOT FILE("nqinclude.prg")
      STRTOFILE(httpGetFile("https://raw.githubusercontent.com/vespina/nqinclude/main/nqinclude.prg"),"nqinclude.prg")
      IF NOT FILE("nqinclude.prg")
         MESSAGEBOX("This library requires NQINCLUDE.PRG wich could not be downloaded at this time",48)
         CANCEL
      ENDIF
   ENDIF
   RETURN

PROCEDURE httpGetFile(pcUrl)
   pnTimeout = IIF(VARTYPE(pnTimeOut)<>"N",15,pnTimeout) 	
   LOCAL oHTTP
   oHTTP = CREATEOBJECT("MSXML2.XMLHTTP")
   oHTTP.open("GET", pcUrl, .F.)
   oHTTP.Send()
   LOCAL nTimeOut
   nTimeout = SECONDS()
   DO WHILE oHTTP.readyState<>4 OR (SECONDS() - nTimeout) > 15
      DOEVENTS
   ENDDO
   IF oHTTP.readyState <> 4 OR !BETWEEN(oHTTP.status,200,299)
      RETURN ""
   ENDIF 
   RETURN oHTTP.responseText()