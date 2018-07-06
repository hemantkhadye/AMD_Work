rcConnectionError = "-6"
rcManualTermination = "-5"
rcFormatError = "-4"
rcSiteError = "-3"
rcNotFound = "-2"
rcLoginFailure = "-1"
rcGeneralFailure = "0"
rcOK = "1"
rcPartcodeSearch = "p"
rcKeywordSearch = "k"
rcContainSearch = "c"




'Dim oXML As Object
'Dim oDom As Object
'Dim strXML As String
'Dim strResponse As String
'Dim strURL As String

Sub WriteVarToDisk(vartowrite, FiletoWrite)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyFile = fso.CreateTextFile(FiletoWrite, True)
    MyFile.WriteLine (vartowrite)
    MyFile.Close

End Sub
Sub getDetails()

End Sub
Sub DebugStart()
    
    Set DataOut = New HTMLPROCESSORLib.ValuePairCollection
    Set DataIn = New HTMLPROCESSORLib.ValuePairCollection
    Set Parser = New HTMLPROCESSORLib.HtmlParser
    
    DataIn.Add "crawltype", "4"
    DataIn.Add "Itemnumber", "Desktop^https://www.amazon.co.uk/s/ref=sr_pg_1?rh=n%3A340831031%2Cn%3A428651031%2Ck%3Agaming%2Cp_6%3AA3P5ROKL5A1OLE&page=1&bbn=428651031&keywords=gaming&ie=UTF8&qid=1377150458"




    
    
    ''''''''''''''''''''''''Commented Hozefa'''''''''''''''''''''''''''''''''''''

'    DataIn.Add "username", ""
'    DataIn.Add "password", ""
'    DataIn.Add "accno", ""
'    DataIn.Add "searchtype", "c"
'    DataIn.Add "maxresult", 10
'
    DataIn.Add "partcode", "B"
    DataOut.Add "manufacturer", ""
    DataOut.Add "description", ""
    DataOut.Add "price", ""
    DataOut.Add "stock", ""
'    ' DataOut.Add "partcode", "1"
    DataIn.Add "RECORD_DATE", ""
'
'   ' DataIn.Add "field1", "3"
'    DataIn.Add "field3", ""
'    DataIn.Add "primarysupp", "Grainger"
'    DataIn.Add "supplier", "wescodirect"
    DataOut.Add "retcode", "-1"
    DataOut.Add "partcode", ""
'    DataIn.Add "cf", "SHP"
    DataOut.Add "field0", ""
    DataOut.Add "field1", ""   'LOB
    DataOut.Add "field2", ""   ' Country
    DataOut.Add "field3", ""   'Site
    DataOut.Add "field4", ""   'ItemNumber
    DataOut.Add "field5", ""   'MPN
    DataOut.Add "field6", ""   'CategoryURL
    DataOut.Add "field7", ""   'ProductName
    DataOut.Add "field8", ""   'ProductURL
    DataOut.Add "field9", ""   ' ListPrice
    DataOut.Add "field10", ""   'PromoPrice
    DataOut.Add "field11", ""   ' CurrencyType
    DataOut.Add "field12", ""   'Processor
    DataOut.Add "field13", ""   'RetailerID
    DataOut.Add "field14", ""   'Date
    DataOut.Add "field15", ""   'Concat
    DataOut.Add "field16", ""   'Active
    DataOut.Add "field17", ""   'Check
    DataOut.Add "field18", ""   'SrNo
    DataOut.Add "field19", ""   'IntelCheck

''''''''''''''''''''''''Commented Hozefa'''''''''''''''''''''''''''''''''''''

    Parser.Licence "B84H-PK3W-NMW1-KLM6"
End Sub

Sub start()
    
    iTotalResultLimit = 10
    On Error Resume Next
    Parser.CookiesEnabled = True
    Parser.RefererEnabled = True
        
    DataOut.Item("retcode").Value = rcOK
    'strPath = GetPath()    '-----Tejas
    
End Sub

Sub login()
    Parser.CookiesEnabled = True
    Parser.RefererEnabled = True
    DataOut.Item("retcode").Value = rcOK
    
End Sub

Sub query()

    
    Parser.RefererEnabled = True
    strModel_Number = ""
    DataOut.Item("retcode").Value = rcNotFound
   
    strModel_Number = DataIn.Item("Itemnumber").Value
   
    AssignFields
    strTempModel_No = strModel_Number
    bMatched_model = 0
     
    
    
    
'    Set oXML = CreateObject("Microsoft.XMLHTTP")
'    With oXML
'    .Open "GET", strModel_Number, False
'    .send
'    End With
    
'    STRHTml = oXML.responseText
'    STRHTml = RemoveChr(STRHTml)

  LOB = GetText2(strModel_Number, "", "^")

    strModel_Number = Replace(strModel_Number, LOB & "^", "")
    
    '''''''''''''''''''''''''''''''XML Code'''''''''''''''''''''''''''''''''''''''''
' Set xmlhttp = Nothing
'        Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
'''''''''''''''''''''''''''' Proxy Code''''''''''''''''''''''''''''''''''''''''''''

PROXY_IP = "shp-prx109-uk-v00002.tp-ns.com"

        PROXY_PORT = 80

        Kpl_Proxy_UserName = "eclerxamd"
        Kpl_Proxy_PassW = "Rid8B67I2Q"

''''''''''''''''''This piece of proxy code is not required in XML''''''''''''''''''
Parser.SetUseProxy 1
    Parser.SetProxyAddress PROXY_IP
    Parser.SetProxyPort PROXY_PORT

    Parser.SetProxyUsername Kpl_Proxy_UserName
    Parser.SetProxyPassword Kpl_Proxy_PassW
'
'
'''''''''''''''''''This piece of proxy code is not required in XML''''''''''''''''''
'
''''''''''''''''''''''''''''' Proxy Code''''''''''''''''''''''''''''''''''''''''''''
'
''''''''''''''''''''''''''''''''XML Code''''''''''''''''''''''''''''''''''''''''''''
'
'
''xmlhttp.SetProxy 2, PROXY_IP & ":" & PROXY_PORT
'
'xmlhttp.Open "GET", strModel_Number, False
''     Call xmlhttp.setRequestHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
''     Call xmlhttp.setRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36")
''     Call xmlhttp.setRequestHeader("Accept-Language", "en-GB,en-US;q=0.8,en;q=0.6")
''     Call xmlhttp.setRequestHeader("Host", "www.amazon.co.uk")
''     Call xmlhttp.setRequestHeader("Proxy-Connection", "Keep-Alive")
'
'       xmlhttp.send
''''
''''
'STRHTml = xmlhttp.responsetext
'Pageload = Replace(Pageload, "Chr10", "")
'Parser.LoadText Pageload



Set XMLHTTP = Nothing
Set XMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
XMLHTTP.Open "GET", strModel_Number, False
XMLHTTP.setProxy 2, PROXY_IP & ":" & PROXY_PORT
XMLHTTP.setProxyCredentials Kpl_Proxy_UserName, Kpl_Proxy_PassW
Call XMLHTTP.setRequestHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8")
     Call XMLHTTP.setRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.143 Safari/537.36")
     Call XMLHTTP.setRequestHeader("Accept-Language", "en-GB,en-US;q=0.8,en;q=0.6")
     Call XMLHTTP.setRequestHeader("Host", "www.amazon.co.uk")
     Call XMLHTTP.setRequestHeader("Proxy-Connection", "Keep-Alive")
'
XMLHTTP.send
StrHTML = XMLHTTP.responseText



'Call WriteVarToDisk(XMLHTTP.responsetext, "kpl_prod.html")
'Call WriteVarToDisk(XMLHTTP.responsetext, "kpl_prod.txt")

'''''''''''''''''''''''''''''''''''''''''XML Code''''''''''''''''''''''''''''''''''''


  


    'Parser.GetPage strModel_Number

''''''''''''Delay Code''''''''''''''''
          'Parser.SetDelay (500)
'          delayval = ""
'          delayval = Round(Rnd(200) * 1000)
'          Parser.SetDelay (delayval)
          
''''''''''''Delay Code''''''''''''''''
          
    '  'Call WriteVarToDisk(STRHTml, "C:\Hoz\Newegg.txt")
 '      Call WriteVarToDisk(Parser.HTML, "C:\Hoz\Amazon_UK.html")
   'MsgBox (STRHTml)

    
'StrHtml = Parser.HTML
'Call WriteVarToDisk(Parser.HTML, "C:\Documents and Settings\AMD\Desktop\Niteesh\amazone_uk.txt")


Bunch = GetText2(StrHTML, "<li id=""result_", "</div></div></li>")


  '''Call WriteVarToDisk(Bunch, "C:\Newegg_bunch.txt")
  
Do While InStr(1, StrHTML, Bunch) > 0

nextpagenumber = InStr(1, StrHTML, "id=""pagnNextLink""")
 '''Call WriteVarToDisk(Parser.HTML, "C:\Newegg_test12.html")

Bunch = GetText2(StrHTML, "<li id=""result_", "</div></div></li>")

While Bunch <> ""
    'MsgBox (Bunch)
    
   
    
    ProductURL = GetText2(Bunch, "href=""", """>")
    ProductURL = Replace(ProductURL, "&amp;", "&")

    'ProductURL = "http://www.amazon.co.uk" & ProductURL
ProdcutURL = Trim(ProductURL)

    ProductName = GetText2(Bunch, "title=""", """")
ProductName = Trim(ProductName)
 
   Manufacturer = GetText2(ProductName, "", " ")
Manufacturer = Trim(Manufacturer)
    
'ItemNumber = GetText2(ProductURL, "/dp/", "/")
ItemNumber = GetText2(Bunch, "asin=""", """")
ItemNumber = Trim(ItemNumber)
    MPN = ItemNumber

    ListPrice = GetText2(Bunch, "price s-price a-text-bold"">", "<")
    ListPrice = Replace(ListPrice, "Â£", "")
    ListPrice = Replace(ListPrice, "£", "")
    ListPrice = Replace(ListPrice, ",", "")
ListPrice = Trim(ListPrice)
    PromoPrice = ListPrice
    
    
   ''''''''''''''''''''''''Hit the ProductURL of 1st product on catalog page'''''''''''''''''''''''''''
'    Parser.GetPage (ProductURL)
'    ''Call WriteVarToDisk(Parser.HTML, "C:\Newegg_Product.html")
'
'    ProductSTRhtml = Parser.HTML
'
'
'    Manufacturer = GetText2(ProductSTRhtml, "", " ")
    
    ''''''''''''''''''''''''''''''''''ProductPage code ends here''''''''''''''''''''''''''''''''''
    
        
        iResults = 0

                                
          If N > 0 Then Call addMoreResult
                DataOut.Item("manufacturer").Value = strmanufacturer
                DataOut.Item("description").Value = ""
                DataOut.Item("price").Value = strPrice
                DataOut.Item("stock").Value = ""
                DataOut.Item("partcode").Value = strModel_Number
                DataOut.Item("field0").Value = LOB
                DataOut.Item("field1").Value = "UK" 'Country
                DataOut.Item("field2").Value = "amazon-uk"   'Site
                DataOut.Item("field3").Value = ItemNumber
                DataOut.Item("field4").Value = MPN
                DataOut.Item("field5").Value = Manufacturer
                DataOut.Item("field6").Value = ProductName
                DataOut.Item("field7").Value = ProductURL
                DataOut.Item("field8").Value = ListPrice
                DataOut.Item("field9").Value = PromoPrice
                DataOut.Item("field10").Value = "GBP"  'CurrencyType
                DataOut.Item("field11").Value = Processor
                DataOut.Item("field12").Value = "28857" 'RetailerId
                DataOut.Item("field13").Value = Now()
                DataOut.Item("field14").Value = Concat
                DataOut.Item("field15").Value = Active
                DataOut.Item("field16").Value = ""
                DataOut.Item("field17").Value = ""
                DataOut.Item("field18").Value = ""
                DataOut.Item("field19").Value = ""
                N = N + 1
    
    
            'strHTML = Replace(strHTML, "<div class=""itemCell""" & Bunch & "id=""addCartHref", "", , 1)
            StrHTML = Replace(StrHTML, "<li id=""result_" & Bunch & "</div></div></li>", "")
         Bunch = GetText2(StrHTML, "<li id=""result_", "</div></div></li>")


             '''Call WriteVarToDisk(Parser.HTML, "C:\Newegg_test.html")
        Wend
            
'            Set XMLHTTP = Nothing
'        Set XMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.4.0")
'        XMLHTTP.setProxy 2, PROXY_IP & ":" & PROXY_PORT
'        XMLHTTP.setProxyCredentials Kpl_Proxy_UserName, Kpl_Proxy_PassW
        
            '''Page Turn'''
             If nextpagenumber > 0 Then
             Oldpage = GetText2(strModel_Number, "page=", "&")
             Oldpage = CInt(Oldpage)
             NewPage = CInt(Oldpage) + 1
              
            strModel_Number = Replace(strModel_Number, "page=" & Oldpage & "&", "page=" & NewPage & "&")
            
               XMLHTTP.Open "GET", strModel_Number, False
              XMLHTTP.send
'
               StrHTML = XMLHTTP.responseText
               
'               Call WriteVarToDisk(XMLHTTP.responsetext, "kpl_prod.html")
'Call WriteVarToDisk(XMLHTTP.responsetext, "kpl_prod.txt")
            'Parser.GetPage strModel_Number
'            ''Call WriteVarToDisk(Parser.HTML, "C:\hoz\Amazon_UK" & ".html")
'        Parser.SetDelay (1000)
            'STRHTml = Parser.HTML
            
          Else
          Exit Do
          
       End If
        
       Loop
        
        
        DataOut.Index = 0
            If N > 0 Then
                DataOut.Item("retcode").Value = rcOK
            Else
                DataOut.Item("retcode").Value = rcNotFound
                StrText = CleanText(Parser.HTML)
                strGUID_ID = GetGuid()
                WriteResultFile GetPath & strGUID_ID & ".html", StrText
                DataOut.Item("stock").Value = strGUID_ID
            End If
        
            End Sub
Sub addMoreResult()
    DataOut.AddResult
    DataOut.Add "manufacturer", ""
    DataOut.Add "description", ""
    DataOut.Add "price", ""
    DataOut.Add "stock", ""
    DataOut.Add "partcode", ""
    For iCT = 0 To 19
        DataOut.Add "field" & CStr(iCT), ""
    Next
End Sub

Sub AssignFields()
    DataOut.Item("partcode").Value = strModel_Number
    'DataOut.Item("manufacturer").Value = DataIn.Item("manufacturer").Value
    
    ''''' This will out the Category URL for eeach record crawled in Field 1'''''''''''''''''
    DataOut.Item("field6").Value = DataIn.Item("Itemnumber").Value
    ''''' This will out the Category URL for eeach record crawled in Field 1'''''''''''''''''

End Sub

Function replaceUnwanted(StrText)

''''' Commented Hozefa'''''''''''''''''''''
'    strText = Replace(strText, ",", "", 1, -1, 1)
'    strText = Replace(strText, "'", "", 1, -1, 1)
'    strText = Replace(strText, Chr(34), "", 1, -1, 1)
'    strText = Replace(strText, Chr(9), "", 1, -1, 1)
'    strText = Replace(strText, Chr(10), "", 1, -1, 1)
'    strText = Replace(strText, Chr(11), "", 1, -1, 1)
'    strText = Replace(strText, Chr(12), "", 1, -1, 1)
'    strText = Replace(strText, Chr(13), "", 1, -1, 1)
'    replaceUnwanted = strText
End Function


Private Function RemoveChr(strString)
    ''''''''''''''''''''''''' Commented Hozefa '''''''''''''''''''''''
    tmp = Trim(Replace(strString, "&nbsp;", ""))
    tmp = Trim(Replace(tmp, "+", " "))
    tmp = Trim(Replace(tmp, "&#039;", "'"))
    tmp = Replace(tmp, "&amp;", " & ")
    tmp = Replace(tmp, "&ldquo;", "“")
    tmp = Replace(tmp, "&rdquo;", "”")
    tmp = Replace(tmp, "&ndash;", "")
    tmp = Replace(tmp, "&mdash;", "-")
    tmp = Replace(tmp, "&trade;", "")
    tmp = Replace(tmp, "&reg;", "")
    tmp = Trim(Replace(tmp, "&pound;", ""))
    tmp = Trim(Replace(tmp, Chr(13), ""))
    tmp = Trim(Replace(tmp, Chr(10), ""))
    tmp = Trim(Replace(tmp, Chr(9), ""))
    tmp = Trim(Replace(tmp, Chr(34), ""))
    tmp = Trim(Replace(tmp, "&#034;", """"))
    tmp = Trim(Replace(tmp, "&#102;", "f"))
    tmp = Trim(Replace(tmp, "&#174;", "®"))
    tmp = Trim(Replace(tmp, "&#153;", "™"))
    tmp = Trim(Replace(tmp, "&#46;", "."))
    tmp = Trim(Replace(tmp, "&#47;", "/"))
    tmp = Trim(Replace(tmp, "&#176;", "°"))
    tmp = Trim(Replace(tmp, "&#38;", "&"))
    tmp = Trim(Replace(tmp, "&quot;", """"))
    tmp = Trim(Replace(tmp, "&#36;", ""))
    tmp = Trim(Replace(tmp, "(CS)", ""))
    tmp = Trim(Replace(tmp, "(QTL)", ""))
    tmp = Trim(Replace(tmp, "(GAL)", ""))
    tmp = Trim(Replace(tmp, "(IN)", ""))
    tmp = Trim(Replace(tmp, "(M)", ""))
    tmp = Trim(Replace(tmp, "(PKG)", ""))
    tmp = Trim(Replace(tmp, "(PR)", ""))
    tmp = Trim(Replace(tmp, "(ROL)", ""))
    tmp = Trim(Replace(tmp, "(SPL)", ""))
    tmp = Trim(Replace(tmp, "(LBS)", ""))
    
    RemoveChr = tmp
End Function


Function CleanText(StrText)
'''''''''''''''''''Commented Hozefa''''''''''''''''''
'    strResults = ""
'    iLen = Len(strText)
'    iSTart = 1
'    iInitPos = 1
'    While iSTart < iLen And iSTart <> 0
'        iSTart = InStr(iSTart, strText, "<script", 1)
'        If iSTart > 0 Then
'            strResults = strResults & Mid(strText, iInitPos, (iSTart - iInitPos))
'            iEnd = InStr(iSTart, strText, "</script", 1)
'            iInitPos = iEnd + Len("</scirpt>")
'            iSTart = iInitPos
'        End If
'    Wend
'    If iInitPos < iLen Then
'        strResults = strResults & Mid(strText, iInitPos)
'    End If
'
'    iBody = InStr(1, strResults, "<body", 1)
'    If iBody > 0 Then
'        iEndBody = InStr(iBody, strResults, ">", 1)
'        strText = Mid(strResults, iBody, (iEndBody - iBody) + 1)
'        strResults = Replace(strResults, strText, "<body>")
'    End If
'    strResults = Replace(strResults, "onmouseout", "c", 1, -1, 1)
'    strResults = Replace(strResults, "onmouseover", "c", 1, -1, 1)
'    CleanText = strResults
End Function
Function midtext(mainText, sipos, sitext, eitext)
    On Error Resume Next
    AAA = InStr(sipos, mainText, sitext, 1)
    AAb = InStr(AAA + Len(sitext), mainText, eitext, 1)
    midtext = Trim(Mid(mainText, AAA + Len(sitext), AAb - (AAA + Len(sitext))))
End Function
Function GetControlField()
    On Error Resume Next
    strTemp = ""
    For I = 0 To DataIn.Count - 1
        Set oVPair = DataIn.Item(CLng(I))
        If oVPair.Name = "cf" Then
            strTemp = oVPair.Value
            Exit For
        End If
    Next
    GetControlField = strTemp
    
End Function
   Function GetGuid()
    Dim rbHotel
    Set rbHotel = CreateObject("DipBag.Helper")
    GetGuid = rbHotel.getFileName()
End Function
 
 Sub WriteResultFile(strFilename, StrText)

    Dim fso, MyFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyFile = fso.opentextFile(strFilename, 2, True)
    MyFile.Write (StrText)
    MyFile.Close
    Set fso = Nothing
    
End Sub
Function GetText2(ByVal StrText, ByVal strStartTag, ByVal strEndTag)
    Dim intStart, intEnd
    intStart = CLng(InStr(1, StrText, strStartTag, vbTextCompare))

    If intStart Then
        intStart = CLng(intStart + Len(strStartTag))
        intEnd = InStr(intStart + 1, StrText, strEndTag, vbTextCompare)
        If intEnd <> 0 Then
         GetText2 = Mid(StrText, intStart, intEnd - intStart)
       Else
          GetText2 = ""
        End If
    Else
        GetText2 = ""
    End If
End Function
Sub write1()
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set MyFile = fso.opentextFile("Tesco.Csv", 8, True)
    
    AAA = Chr(34) & DataOut.Item("Manufacturer").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("description").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("stock").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Price").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Field0").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Field1").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Field2").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Field3").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Field4").Value & Chr(34) & ","
    'aaa = aaa & Chr(34) & DataOut.Item("Field5").Value & Chr(34) & ","
    'aaa = aaa & Chr(34) & DataOut.Item("Field6").Value & Chr(34) & ","
    AAA = AAA & Chr(34) & DataOut.Item("Field7").Value & Chr(34)
    
    MyFile.WriteLine AAA
    
    MyFile.Close
    Set fso = Nothing
    
End Sub
Sub GetProxy()



On Error Resume Next
   
    Set ofso = CreateObject("Scripting.FileSystemObject")
    If ofso.FileExists("C:\Proxyinfo\ProxyList.txt") Then
        Set proxyfile = ofso.opentextFile("C:\Proxyinfo\ProxyList.txt", 1)
        proxytext = proxyfile.readall
        proxyfile.Close
    End If
    Set ofso = Nothing
    
ProxyCount = CountText
    proxylist = proxytext
    ProxyIp = Split(proxylist, ",")

For Q = 0 To UBound(ProxyIp) - 1
     IPPort = Split(ProxyIp(Q), ":")
           
        Parser.SetUseProxy "1"
        Parser.SetProxyAddress IPPort(0)
        Parser.SetProxyPort IPPort(1)
        Parser.GetPage "http://www1.mscdirect.com/"
        
       WriteVarToDisk Parser.HTML, "ProxyTest.html"
       
       If InStr(1, Parser.HTML, "2000 - 2012 MSC Industrial Direct", 1) > 1 Then
                PROXY_IP = IPPort(0)
                PROXY_PORT = IPPort(1)
                
                'WriteVarToDisk Q, "C:\Proxyinfo\ProxyNumber.txt"
                Exit For
            End If
       
    Next
 Set xmlHttpProxy = Nothing
 Set ofso = Nothing
 'Set ofso2 = Nothing
End Sub









