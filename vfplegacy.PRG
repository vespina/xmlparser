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