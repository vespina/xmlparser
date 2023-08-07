# XMLPARSER.PRG
**XML parser for VFP**

Version: 1.3

Author: V. Espina

----

### BASIC USAGE

#### Load the library
Simply addd this code at the begining of your main program:

    DO xmlparser

This will create a public object called xmlParser. If you prefer to create local instances instead, you can do it by using the xmlParser class:

    SET PROCEDURE TO xmlParser ADDITIVE
    ..
    LOCAL oParser
    oParser = CREATEOBJECT("xmlParser")

 
#### To parse a XML file
    LOCAL oData
    oData = xmlParser.Parse("myfile.xml")
    IF ISNULL(oData)
      ?xmlParser.lastError
      RETURN
    ENDIF
    ?oData.nodeValue
    ?oData.node.attribute
    

#### To parse a XML string
    LOCAL oData
    oData = xmlParser.ParseString("<document><hello>World</hello></document>")
    IF ISNULL(oData)
      ?xmlParser.lastError
      RETURN
    ENDIF
    ?oData.Hello -> "World"


### EXAMPLE #1
    ** TEST1.XML
    <document>
       <customer id="001"  name="VICTOR ESPINA" />
       <invoices>
         <invoice>
           <number>02002</number>
           <date>01/01/2019</date>
           <amount>23.25</amount>
         </invoice>
         <invoice>
           <number>02010</number>
           <date>07/01/2019</date>
           <amount>32.25</amount>
         </invoice>
      </invoices>
    </document>
    
    oData = XmlParser.Parse("test1.xml")
    ?oData.Customer.Id.   --> "001"
    ?oData.Customer.Name  --> "VICTOR ESPINA"
    ?oData.Invoices.Count --> 2
    oInvoice = oData.Invoices.Items[1]
    ?oInvoice.number --> "02002"
    ?oInvoice.amount --> 32.25


### EXAMPLE #2
    TEXT TO cXML NOSHOW
    <?xml version="1.0" encoding="UTF-8"?>
    <ns1:Envelope xmlns:ds="http://www.w3.org/2000/09/xmldsig#" xmlns:ns1="http://www.w3.org/2001/12/soap-envelope" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="http://www.w3.org/2001/12/soap-envelope http://www.codaladi.org/directorio/cod_ver_1.8.0.xsd">
       <ns1:CertOrigin>
          <CODEH id="CODEH">
            <CODExporter>
                <COD id="COD">
                    <CODVer>1.8.0</CODVer>
                    <CODSubmitterType>EXP</CODSubmitterType>
                    <Agreement>
                        <AgreementName>A.C.E. Nro 18</AgreementName>
                        <AgreementAcronym>A18</AgreementAcronym>
                    </Agreement>
                </COD>
            </CODExporter>
          </CODEH>
       </ns1:CertOrigin>
    </ns1:Envelope>
    ENDTEXT
    
    oData = XmlParser.parseString(cXML)
    ?oData.ns1_envelope.xmlns_ds --> "http://www.w3.org/2000/09/xmldsig#"
    ?oNS1_Envelope.ns1_CertOrigin.CODEH.CODExporter.COD.id  --> "CODEH"
    ?oNS1_Envelope.ns1_CertOrigin.CODEH.CODExporter.COD.CODVer  --> "1.8.0"
    
    


### CHANGE HISTORY

|Date         |User|Description|
|-------------|----|-----------|
[Jul, 2023  |VES |Version 1.3. VFPLEGACY.PRG integrated on XMLPARSER.PRG|
[Dic, 2022  |VES |Version 1.2. Fix error with prefixed nodes|
[Ago, 2022  |VES |Version 1.1|
|Ago, 2022  |VES |Improved collection's detection. CDATA support|
|Jul, 2022  |VES |New parseString() method. Fix on collection's detection|
|Jun, 2019  |VES |Version 1.0|



