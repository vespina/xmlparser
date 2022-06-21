* XMLPARSER.PRG
* Parser XML para VFP (6.0 o superior)
*
* Autor: Victor Espina
* Version: 1.0
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
* Jun 2, 2019  VES     Version inicial
*
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
   
       oData = THIS._parseNode(oRootNode)
   
       RETURN oData
     ENDPROC
   
   
     HIDDEN PROCEDURE _parseNode(poNode)
        LOCAL oData,cNodeName,i,oNode,oAttr,lHasChildren,cAttrName,cNodeName
        cNodeName = LOWER(poNode.nodeName)
        lHasChildren = (TYPE("poNode.childNodes.Length") = "N" AND poNode.childNodes.Length > 0)
   
        * Determinamos si es una coleccion. Es importante hacerlo aqui porque
        * si es una coleccion no se leeran los atributos que pueda tener definidos
        LOCAL lIsCollection,oCollection
        lIsCollection = .F.
   
        IF lHasChildren
          lIsCollection = .T.    && Asumimos que es una coleccion a menos que encontremos un hijo con otro nombre
          FOR i = 1 TO poNode.childNodes.Length
            oNode = poNode.childNodes.Item(i - 1)
            IF oNode.nodeType = 8  && Comentario
              LOOP
            ENDIF
            IF !(LOWER(oNode.nodeName) + "s" == cNodeName)
              lIsCollection = .F.
              EXIT
            ENDIF
          ENDFOR
        ENDIF
        oData = IIF(lISCollection, CREATEOBJECT("XmlParserCollection"), THIS._createEmpty())
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
   
   
        * Si el nodo tiene hijos, se procesan
        IF lHasChildren
          FOR i = 1 TO poNode.ChildNodes.Length
            oNode = poNode.childNodes.Item(i - 1)
            IF oNode.nodeType = 8  && Comentario
              LOOP
            ENDIF
            oValue = NULL
            DO CASE
               CASE oNode.childNodes.Length = 1 AND oNode.childNodes.Item[0].nodeType = 3
                    oValue = oNode.childNodes.Item[0].text                    
               
               CASE oNode.childNodes.Length = 0 AND (ISNULL(oNode.Attributes) OR oNode.Attributes.Length = 0)
                    oValue = oNode.Text
   
   	           CASE !lIsCollection AND PEMSTATUS(oData, oNode.nodeName, 5)
   	                lIsCollection = .T.
         			oCollection = CREATEOBJECT("xmlParserCollection")
         			oCollection.Add(GETPEM(oData, oNode.nodename))
         			STORE oCollection TO ("oData." + oNode.nodeName)
         			   	                
               OTHERWISE
                    oValue = THIS._parseNode(oNode)
            ENDCASE
            IF lIsCollection
              oCollection.Add(oValue)
            ELSE
              cNodeName = LOWER(oNode.nodeName)
              THIS._ADDPROPERTY(@oData, cNodeName, oValue)
            ENDIF
          ENDFOR
        ENDIF
   
        RETURN oData
     ENDPROC
   
   
     HIDDEN PROCEDURE _createEmpty(pcColList)
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
       RETURN oEmpty
     ENDPROC
   
   
     HIDDEN PROCEDURE _addProperty(poTarget, pcProperty, puValue)
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
       pcProperty = CHRT(pcProperty,":","_")
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
       RETURN poTarget
     ENDPROC
ENDDEFINE         

DEFINE CLASS xmlParserCollection AS Custom
               
     HIDDEN ncount
     ncount = 0
   
     HIDDEN leoc
     leoc = .T.
   
     HIDDEN lboc
     lboc = .T.
   
     *-- Indica la posición actual dentro de la colección
     LISTINDEX = 0
   
     *-- Nombre de la clase a instanciar al llamar al método New.
     newitemclass = ""
     NAME = "cbasiccollection"
   
     *-- Nro. de elementos en la colección
     COUNT = .F.
   
     *-- Indica si se ha llegado al final de la colección
     eoc = .F.
   
     *-- Indica si se ha llegado al tope de la colección.
     boc = .F.
   
     *-- Devuelve el valor actual en la colección
     CURRENT = .F.
   
     *-- Lista de elementos en la colección
     DIMENSION items[1,1]
     PROTECTED aitems[1,1]
   
   
     PROCEDURE items_Access
         LPARAMETERS m.nIndex1, m.nIndex2
   
         IF TYPE("m.nIndex1")="C"
             m.nIndex1 = THIS.FindItem(m.nIndex1)
         ENDIF
   
         RETURN THIS.aitems[m.nIndex1]
     ENDPROC
   
   
     PROCEDURE items_Assign
         LPARAMETERS vNewVal, m.nIndex1, m.nIndex2
   
         IF TYPE("m.nIndex1") = "C"
             m.nIndex1 = THIS.FindItem(m.nIndex1)
         ENDIF
   
         IF BETWEEN(m.nIndex1, 1, THIS.COUNT)
             THIS.aitems[m.nIndex1] = m.vNewVal
         ENDIF
     ENDPROC
   
   
     PROCEDURE count_Access
   
         RETURN THIS.ncount
     ENDPROC
   
   
     PROCEDURE count_Assign
         LPARAMETERS vNewVal
     ENDPROC
   
   
     *-- Añade un elemento a la colección
     PROCEDURE ADD
         LPARAMETERS puValue
   
         IF PARAMETERS()=0
             RETURN .F.
         ENDIF
   
         THIS.ncount = THIS.ncount + 1
         DIMEN THIS.aitems[this.nCount]
         THIS.aitems[this.nCount] = puValue
         THIS.leoc = .F.
         THIS.lboc = .F.
   
         IF THIS.LISTINDEX = 0
             THIS.LISTINDEX = 1
         ENDIF
   
         RETURN puValue
     ENDPROC
   
   
     *-- Elimina un elemento de la colección
     PROCEDURE REMOVE
         LPARAMETERS puValue
   
         IF PARAMETERS() = 0
             RETURN .F.
         ENDIF
   
         LOCAL nIndex
         nIndex = THIS.FindItem(puValue)
         IF nIndex > 0
             RETURN THIS.REMOVEITEM(nIndex)
         ELSE
             RETURN .F.
         ENDIF
     ENDPROC
   
   
     *-- Limpia la colección
     PROCEDURE CLEAR
         LOCAL i,uItem
         FOR i = 1 TO THIS.ncount
             uItem = THIS.aitems[i]
             IF TYPE("uItem") = "O"
                 RELEASE uItem
                 THIS.aitems[i] = NULL
             ENDIF
         ENDFOR
   
         DIMEN THIS.aitems[1]
         THIS.aitems[1] = NULL
         THIS.ncount = 0
         THIS.lboc = .T.
         THIS.leoc = .T.
         THIS.LISTINDEX = 0
     ENDPROC
   
   
     *-- Determina si un elemento dado forma parte de la colección.
     PROCEDURE isitem
         LPARAMETERS puValue,pcSearchProp
   
         IF PARAMETERS() = 0
             RETURN .F.
         ENDIF
   
         RETURN (THIS.FindItem(puValue,pcSearchProp) <> 0)
     ENDPROC
   
   
     *-- Devuelve la posición en la colección donde se encuentra el elemento indicado
     PROCEDURE FindItem
         LPARAMETERS puValue, pcSearchProp
   
   
         IF PARAMETERS() = 0 OR THIS.ncount = 0
             RETURN 0
         ENDIF
   
         IF VARTYPE(pcSearchProp) <> "C"
             pcSearchProp = ""
         ENDIF
   
         LOCAL i, uItem, cType1, nPos
         nPos = 0
         cType1 = TYPE("puValue")
         FOR i = 1 TO THIS.ncount
             uItem = THIS.aitems[i]
             IF TYPE("uItem") = "O"
                 IF (cType1 = "O" AND TYPE("uItem.Name") = "C" AND TYPE("puVale.Name") = "C" AND UPPER(uItem.NAME) == UPPER(puValue.NAME)) OR ;
                         (cType1 = "C" AND TYPE("uItem.Name") = "C" AND UPPER(uItem.NAME) == UPPER(puValue)) OR ;
                         (cType1 <> "O" AND NOT EMPTY(pcSearchProp) AND TYPE("uItem." + pcSearchProp) = cType1 AND EVAL("uItem." + pcSearchProp) == puValue)
                     nPos = i
                     EXIT
                 ENDIF
             ELSE
                 IF TYPE("uItem") = cType1 AND ((cType1 <> "C" AND uItem = puValue) OR (cType1 = "C" AND uItem == puValue))
                     nPos = i
                     EXIT
                 ENDIF
             ENDIF
         ENDFOR
   
         RETURN nPos
     ENDPROC
   
   
     *-- Elimina un item por su posición
     PROCEDURE REMOVEITEM
         LPARAMETERS nIndex
   
         IF PARAMETERS() = 0 OR NOT BETWEEN(nIndex, 1, THIS.ncount)
             RETURN .F.
         ENDIF
   
         LOCAL uItem
         uItem = THIS.aitems[nIndex]
   
         IF TYPE("uItem") = "O"
             RELEASE uItem
             THIS.aitems[nIndex] = NULL
         ENDIF
   
         ADEL(THIS.aitems,nIndex)
   
         THIS.ncount = THIS.ncount - 1
         IF THIS.ncount > 0
             DIMEN THIS.aitems[this.nCount]
             IF THIS.ncount > THIS.LISTINDEX
                 THIS.LISTINDEX = THIS.ncount
             ENDIF
         ELSE
             THIS.aitems[1] = NULL
             THIS.leoc = .T.
             THIS.lboc = .T.
             THIS.LISTINDEX = 0
         ENDIF
     ENDPROC
   
   
     PROCEDURE eoc_Access
   
         RETURN THIS.leoc
     ENDPROC
   
   
     PROCEDURE eoc_Assign
         LPARAMETERS vNewVal
     ENDPROC
   
   
     PROCEDURE boc_Access
   
         RETURN THIS.lboc
     ENDPROC
   
   
     PROCEDURE boc_Assign
         LPARAMETERS vNewVal
     ENDPROC
   
   
     *-- Ir al primer elemento en la colección
     PROCEDURE FIRST
         IF THIS.ncount = 0
             RETURN
         ENDIF
   
         THIS.LISTINDEX = 1
         THIS.lboc = .F.
         THIS.leoc = .F.
     ENDPROC
   
   
     *-- Ir al siguiente elemento en la colección
     PROCEDURE NEXT
         IF THIS.ncount = 0
             RETURN
         ENDIF
   
         IF THIS.LISTINDEX < THIS.ncount
             THIS.LISTINDEX = THIS.LISTINDEX + 1
             THIS.lboc = .F.
             THIS.leoc = .F.
         ELSE
             THIS.lboc = (THIS.ncount=1)
             THIS.leoc = .T.
         ENDIF
     ENDPROC
   
   
     *-- Ir al último elemento en la colección
     PROCEDURE LAST
         IF THIS.ncount = 0
             RETURN
         ENDIF
   
         THIS.LISTINDEX = THIS.ncount
         THIS.lboc = .F.
         THIS.leoc = .F.
     ENDPROC
   
   
     *-- Ir al elemento anterior en la colección
     PROCEDURE previous
         IF THIS.ncount = 0
             RETURN
         ENDIF
   
         IF THIS.LISTINDEX > 1
             THIS.LISTINDEX = THIS.LISTINDEX - 1
             THIS.lboc = .F.
             THIS.leoc = .F.
         ELSE
             THIS.lboc = .T.
             THIS.leoc = (THIS.ncount = 1)
         ENDIF
     ENDPROC
   
   
     PROCEDURE listindex_Assign
         LPARAMETERS vNewVal
   
         IF TYPE("m.vNewVal") = "N" AND BETWEEN(m.vNewVal, 1, THIS.ncount)
             THIS.LISTINDEX = m.vNewVal
             THIS.leoc = .F.
             THIS.lboc = .F.
         ENDIF
     ENDPROC
   
   
     PROCEDURE current_Access
         IF THIS.LISTINDEX = 0
             RETURN NULL
         ELSE
             RETURN THIS.aitems[this.ListIndex]
         ENDIF
     ENDPROC
   
   
     PROCEDURE current_Assign
         LPARAMETERS vNewVal
   
         IF THIS.LISTINDEX > 0
             THIS.aitems[this.ListIndex] = m.vNewVal
         ENDIF
     ENDPROC
   
   
     *-- Crea una instancia de la clase indicada en NewItemClass y devuelve una referencia al mismo.
     PROCEDURE new
         IF EMPTY(THIS.newitemclass)
             RETURN NULL
         ENDIF
   
         LOCAL oItem
         oItem = Kernel.CC.new(THIS.newitemclass)
   
         RETURN oItem
     ENDPROC
   
   
     *-- Permite añadir un item a la colección, solo si el mismo no existe.
     PROCEDURE addifnew
         LPARAMETERS puValue
   
         IF PARAMETERS() = 0
             RETURN .F.
         ENDIF
   
         IF NOT THIS.isitem(puValue)
             THIS.ADD(puValue)
         ENDIF
   
         RETURN puValue
     ENDPROC
ENDDEFINE