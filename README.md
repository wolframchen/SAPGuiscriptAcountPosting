# SAPGuiscriptAcountPosting
General Component for Posting Accounting Documents


it is an abstration of SAP Guiscripting Posting Documents by for example FB01 
based on anthoer old posting about excel and guiscript  by https://people.sap.com/st.schnell .
many thanks to this posting we automated lots of accounting operations.

because most of the time we need to post documents, so I create this abstraction layer for general purpose accounting posting, so we don't have to record FB01 again and again for different situations.



please check the file for the actual coding, sorry it is quite messy; I deleted alot of codes and sheets that is irrelevant  so it is not working
you need to learn how to compose an actual running program anyway...






pay attention to the Sub PlayBack_The_Script , this one actually execute the guiscript.

BEFORE that you need to activate the SAP session by
    On Error GoTo ErrorHandler1
    Set session = GetSession("XXP") ' your SAP ID, you can put there the value of an excel cell so that user can change it
    
    Read_Spreadsheet ' read sap data by the index
    ' then you do a lot of preparations later 
    
    'then please fillin the data array for document Header :
    dH(1) = m1(3): dH(2) = today: dH(3) = "SA": dH(4) = "assingment text"
    dH(5) = today: dH(6) = "USD": dH(7) = Left(lastmonth, 6) & scnr & "-" & Str(lDataRow - 2)
    
    'within a loop you can add the data into the array for item 
    For lvi = 1 To 18
      item = item + 1
       dI(item, 1) = "40": dI(item, 2) = ac1(lvi + 2):  dI(item, 4) = amt:
                dI(item, 11) = Left(lastmonth, 6) & d & "some other text"
    next lvi
    'then:
     PlayBack_The_Script
    ' later you will get the errs also in an array    ; you can extract the document number or the error messages :
  If regstr("\d{10} was posted", Er(5)) <> "" Then Sheets("SAP_DATA").Cells(lDataRow, dpos + rpos) = regstr("\d{10}", Er(5)) & ";" & regstr("CN[A-Z0-9]{2}", Er(5)) & ";" & Left(today, 4)  
    
    
    For lvi = 0 To 6
        Sheets("SAP_DATA").Cells(lDataRow, epos + rpos) = Sheets("SAP_DATA").Cells(lDataRow, epos + rpos) & ";" & Er(lvi + 4)
    Next lvi

    ' below is the mapping  (copy to excel Sheet "CUST" to ge it working )
    you will notice there is few fields that can use regular expression, and some fields need to jump to other screen.   play with it you will see.
    
|     Header FieldsName(check with E1!!)(.\*as RegExp!!) | 3               | Additional Menu/btn         | Exit Popup            | Mandatory? | 7 | Items FieldsName(check with E1!!)(.\*as RegExp!!) | 9                 | Additional Menu/btn                  | Exit Popup            | Mandatory? |
| ------------------------------------------------------ | --------------- | --------------------------- | --------------------- | ---------- | - | ------------------------------------------------- | ----------------- | ------------------------------------ | --------------------- | ---------- |
| ctxtBKPF-BUKRS                                         | Company Code    |                             |                       |            |   | ctxtRF05A-NEWBS                                   | PstKy             |                                      |                       |            |
| ctxtBKPF-BLDAT                                         | Document Date   |                             |                       |            |   | ctxtRF05A-NEWKO                                   | Account           |                                      |                       |            |
| ctxtBKPF-BLART                                         | Type            |                             |                       |            |   | ctxtRF05A-NEWUM                                   | special indicator |                                      |                       |            |
| txtBKPF-XBLNR                                          | REFERENCE       |                             |                       |            |   | txtBSEG-WRBTR                                     | Amount            |                                      |                       |            |
| ctxtBKPF-BUDAT                                         | Posting Date    |                             |                       |            |   | ctxtCOBL-RMVCT                                    | Transactn Type    | subBLOCK:SAPLKACB:.\*\\/btnCOBL_MORE | wnd[1]/tbar[0]/btn[0] |            |
| ctxtBKPF-WAERS                                         | Currency/Rate   |                             |                       |            |   | ctxtBSEG-ZFBDT                                    | due on            |                                      |                       |            |
| txtBKPF-BKTXT                                          | Header Text     |                             |                       |            |   | subBLOCK:SAPLKACB:.\*\\/ctxtCOBL-AUFNR            | order             |                                      |                       |            |
| ctxtRF014-VBUND                                        | Trading Partner | wnd[0]/mbar/menu[3]/menu[1] | wnd[1]/tbar[0]/btn[0] |            |   |                                                   | business area     |                                      |                       |            |
|                                                        |                 |                             |                       |            |   | subBLOCK:SAPLKACB:.\*\\/ctxtCOBL-KOSTL            | cost center       |                                      |                       |            |
|                                                        |                 |                             |                       |            |   | subBLOCK:SAPLKACB:.\*\\/ctxtCOBL-PS_POSID         | WBS Element       |                                      |                       |            |
|                                                        |                 |                             |                       |            |   | ctxtBSEG-SGTXT                                    | Text              |                                      |                       |            |
|                                                        |                 |                             |                       |            |   | ctxtBSEG-MWSKZ                                    | Tax code          |                                      |                       |            |
|                                                        |                 |                             |                       |            |   | chkBKPF-XMWST                                     | CalcTax           |                                      |                       | N          |
|                                                        |                 |                             |                       |            |   | txtBSEG-ZUONR                                     | Assignment        |                                      |                       |            |
|                                                        |                 |                             |                       |            |   |                                                   | RF3               |                                      |                       |            |
