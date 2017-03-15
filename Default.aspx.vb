Imports System.Xml
Imports System.Xml.Serialization
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Net
Imports System.IO
Imports System.Net.Mail
Imports System.Net.Mime

Partial Class _Default
    Inherits System.Web.UI.Page

    Dim xmlDoc As New XmlDocument
    Dim xmlDocNodeList As XmlNodeList
    Dim xmlCXMLNode As XmlNode
    Dim xmlNodeBillTo As XmlNodeList
    Dim StrResponse As String
    Dim StrError As String
    Dim OrderRequestHeader As XmlNode
    Dim InfoLine As XmlNode
    Dim OrdNbr As String
    Dim TotalAmount As Integer
    Dim LinNum As Integer
    Dim LinQuant As Double
    Dim LinSKU As String
    Dim LinPrice As Double
    Dim EntryXMLTimeStamp As String = ""
    Dim OrderDate As String = ""
    Dim TotalOrder As String = ""
    Dim PNASku, CustNbr As String
    Dim Node_PNASku As XmlNode
    Dim ConB2B As New SqlConnection
    Dim ConB2B2 As New SqlConnection
    Dim cmdInsertStatus As New SqlCommand
    Dim carrierCode, User, Password As String
    Dim SuffixSql As String

    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        User = "ESPINOLA"
        Password = "Imes-12345"
        carrierCode = "GN"

        ConB2B.ConnectionString = ConfigurationManager.ConnectionStrings("ConB2B").ConnectionString
        ConB2B2.ConnectionString = ConfigurationManager.ConnectionStrings("ConB2B").ConnectionString
        Dim ErrorFound As Boolean = False
        Dim ErrorString As String = ""
        'Dim OrderID1 As String = ""
        Dim OrderID As String = ""
        Dim OrderType As String = ""
        Dim TotalMoney As String = ""
        Dim gSave As New Guid
        gSave = Guid.NewGuid()
        Dim payloadID As String = ""
        Response.ContentType = "text/xml"
        If Request.TotalBytes > 0 Then
            StrResponse = "Sin datos"



            If (xmlDoc.Value <> 0) Then
                StrError = xmlDoc.Value
                StrResponse = "Error"
                'Response.Write("Error")
            Else

            End If

            Try

                ' ____ BORRAR LINEA DTD ____ '
                xmlDoc.Load(Request.InputStream)
                Dim XDType As XmlDocumentType = xmlDoc.DocumentType
                xmlDoc.RemoveChild(XDType)
                'xmlDoc.Save("C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml") 'Producció
                xmlDoc.Save("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml") 'Local
                'Dim linesList As New List(Of String)(System.IO.File.ReadAllLines("C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml"))
                Dim linesList As New List(Of String)(System.IO.File.ReadAllLines("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml"))
                linesList.RemoveAt(0)
                'linesList.RemoveAt(1)

                'xmlDoc.Save("C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml") 'Producció
                xmlDoc.Save("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml") 'Local

                'File.WriteAllLines("C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml", linesList.ToArray()) 'Producció
                File.WriteAllLines("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml", linesList.ToArray()) 'Local
                'xmlDoc.Load("C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml") 'Producció
                xmlDoc.Load("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml") 'Local
                'EntryXMLTimeStamp = xmlCXMLNode.Attributes("timestamp").InnerText


            Catch ex As Exception
                ErrorFound = True
                'Response.Write(ex.ToString)
                ErrorString = ex.ToString()
            End Try
            xmlCXMLNode = xmlDoc.SelectSingleNode("cXML")
            Dim Node_OrderID As XmlNode

            '__________ INFO GENERAL __________'
            Try
                EntryXMLTimeStamp = xmlCXMLNode.Attributes("timestamp").InnerText
                OrderDate = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader").Attributes("orderDate").InnerText
                TotalOrder = xmlDoc.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/Total/Money").Value
                payloadID = xmlCXMLNode.Attributes("payloadID").InnerText
                Node_OrderID = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader").Attributes("orderID")
                OrderID = GetInnerText(Node_OrderID)

            Catch ex As Exception
                'Response.Write("Error Info General")
                ErrorFound = True
                ErrorString = ErrorString + "-Error Info General"
            End Try



            '____ ACK TO ARIBA ____ '

            Dim xmlResponderAribaOK As String = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
            xmlResponderAribaOK = xmlResponderAribaOK & "<!DOCTYPE cXML SYSTEM ""http://xml.cxml.org/schemas/cXML/1.2.014/cXML.dtd"">" & vbCrLf
            xmlResponderAribaOK = xmlResponderAribaOK & "<cXML payloadID='" + payloadID + "' xml:lang=""en"" timestamp=""2002-03-12T18:39:09-08:00"">" & vbCrLf
            xmlResponderAribaOK = xmlResponderAribaOK & " <Response>" & vbCrLf
            xmlResponderAribaOK = xmlResponderAribaOK & " <Status code=""200"" text=""OK""/>" & vbCrLf
            xmlResponderAribaOK = xmlResponderAribaOK & "</Response>" & vbCrLf
            xmlResponderAribaOK = xmlResponderAribaOK & "</cXML>"
            Response.Write(xmlResponderAribaOK)
            Response.Flush()

            '____ INSERT ORDER ID @DB ORDER_HEADER ____'
            Dim cmdinsertOrderID As New SqlCommand
            ConB2B.Open()
            Dim insertedID As String = ""
            cmdinsertOrderID.CommandText = "INSERT INTO ARIBA_OrderRequest_Header (OrderID) VALUES ('" + OrderID + "'); Select SCOPE_IDENTITY()"
            cmdinsertOrderID.Connection = ConB2B
            Try
                insertedID = cmdinsertOrderID.ExecuteScalar()

            Catch ex As Exception
                ErrorFound = True
                ErrorString = "- Error insertando OrderID"
            End Try
            ConB2B.Close()


            '___________________ GUARDAR DB __________________________'

            Try
                ConB2B.Open()
            Catch ex As Exception
                ErrorFound = True
                ErrorString = ErrorString + "-Error al conectar a B2B"
            End Try

            ' ____ IDENTIFICADORES ____'
            Dim From_NID, From_SID, To_NID As String
            Dim Node_From_NID, Node_From_SID, Node_To_NID As XmlNode
            Try
                Node_From_NID = xmlCXMLNode.SelectSingleNode("/cXML/Header/From/Credential[@domain='NetworkID']/Identity")
                From_NID = GetInnerText(Node_From_NID)
                Node_From_SID = xmlCXMLNode.SelectSingleNode("/cXML/Header/From/Credential[@domain='SystemID']/Identity")
                From_SID = GetInnerText(Node_From_SID)
                Node_To_NID = xmlCXMLNode.SelectSingleNode("/cXML/Header/To/Credential[@domain='NetworkID']/Identity")
                To_NID = GetInnerText(Node_To_NID)
            Catch ex As Exception
                ErrorFound = True
                ErrorString = ErrorString + "-Error IDs"
            End Try

            ' ____ ORDER INFO ____ '

            Dim Node_OrderType, Node_TotalMoney As XmlNode
            Try

                Node_OrderType = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader").Attributes("orderType")
                OrderType = GetInnerText(Node_OrderType)
                Node_TotalMoney = xmlDoc.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/Total/Money")
                TotalMoney = GetInnerText(Node_TotalMoney)
            Catch ex As Exception
                'Response.Write("Error OrderInfo")
                ErrorFound = True
                ErrorString = ErrorString + "-Error OrderInfo"
            End Try

            ' ____ SHIPTO ____ '
            Dim STName, STAddressID, STPAName, STPADeliverTo, STPAStreet, STPACity, STPAState, STPAPostalCode, STPACountry As String
            Dim Node_STName, Node_STAddressID, Node_STPAName, Node_STPADeliverTo, Node_STPAStreet, Node_STPACity, Node_STPAState, Node_STPAPostalCode, Node_STPACountry As XmlNode

            Try
                Node_STName = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/Name")
                STName = GetInnerText(Node_STName)
                Node_STAddressID = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address").Attributes("addressID")
                STAddressID = GetInnerText(Node_STAddressID)
                Node_STPAName = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress").Attributes("name")
                STPAName = GetInnerXml(Node_STPAName)
                Node_STPADeliverTo = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress/DeliverTo")
                STPADeliverTo = GetInnerText(Node_STPADeliverTo)
                Node_STPAStreet = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress/Street")
                STPAStreet = GetInnerText(Node_STPAStreet)
                Node_STPACity = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress/City")
                STPACity = GetInnerText(Node_STPACity)
                Node_STPAState = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress/State")
                STPAState = GetInnerText(Node_STPAState)
                Node_STPAPostalCode = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress/PostalCode")
                STPAPostalCode = GetInnerText(Node_STPAPostalCode)
                Node_STPACountry = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/ShipTo/Address/PostalAddress/Country").Attributes("isoCountryCode")
                STPACountry = GetInnerText(Node_STPACountry)

            Catch ex As Exception
                ErrorString = ErrorString + "-ErrorShipTo"
                ErrorFound = True
            End Try

            ' ____ EXTRINSICS ____ '
            Dim ExtCompanyCode, ExtBuyerPurchasingCode, ExtVendorIDNbr, ExtBuyerVatID As String
            Dim Node_ExtCompanyCode, Node_ExtBuyerPurhcasingCode, Node_ExtVendorIDNbr, Node_ExtBuyerVatID As XmlNode
            Try
                Node_ExtCompanyCode = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/Extrinsic[@name='CompanyCode']")
                ExtCompanyCode = GetInnerText(Node_ExtCompanyCode)
                Node_ExtBuyerPurhcasingCode = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/Extrinsic[@name='buyerPurchasingCode']")
                ExtBuyerPurchasingCode = GetInnerText(Node_ExtBuyerPurhcasingCode)
                Node_ExtVendorIDNbr = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/Extrinsic[@name='vendorIDNo']")
                ExtVendorIDNbr = GetInnerText(Node_ExtVendorIDNbr)
                Node_ExtBuyerVatID = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader/Extrinsic[@name='buyerVatID']")
                ExtBuyerVatID = GetInnerText(Node_ExtBuyerVatID)
            Catch ex As Exception

            End Try

            ' ____ Find ShipToSql ____ '

            Dim cmdSelectQuerySuffix, cmdSelectQueryKeyWord, cmdSelectQueryPC, cmdSelectQueryCity As New SqlCommand

            Dim KeyWord(3) As String

            cmdSelectQueryKeyWord.CommandText = "SELECT KeyWord1 FROM ARIBA_CustomerInfo WHERE (Ext_BuyerVATID LIKE '" + ExtBuyerVatID + "%' AND ST_PA_City = '" + STPACity + "' AND ST_PA_PostalCode = '" + STPAPostalCode + "')"
            cmdSelectQueryKeyWord.Connection = ConB2B2
            ConB2B2.Open()
            KeyWord(0) = cmdSelectQueryKeyWord.ExecuteScalar()
            cmdSelectQueryKeyWord.CommandText = "SELECT KeyWord2 FROM ARIBA_CustomerInfo WHERE (Ext_BuyerVATID LIKE '" + ExtBuyerVatID + "%' AND ST_PA_City = '" + STPACity + "' AND ST_PA_PostalCode = '" + STPAPostalCode + "')"
            KeyWord(1) = cmdSelectQueryKeyWord.ExecuteScalar()
            cmdSelectQueryKeyWord.CommandText = "SELECT KeyWord3 FROM ARIBA_CustomerInfo WHERE (Ext_BuyerVATID LIKE '" + ExtBuyerVatID + "%' AND ST_PA_City = '" + STPACity + "' AND ST_PA_PostalCode = '" + STPAPostalCode + "')"
            KeyWord(2) = cmdSelectQueryKeyWord.ExecuteScalar()


            Dim FirstKeyWord As Integer = STPAStreet.IndexOf(KeyWord(0))
            Dim SecondKeyWord As Integer = STPAStreet.IndexOf(KeyWord(1))
            Dim ThirdKeyWord As Integer = STPAStreet.IndexOf(KeyWord(2))

            If (FirstKeyWord <> -1 And SecondKeyWord <> -1 And ThirdKeyWord <> -1) Then

                Try
                    cmdSelectQuerySuffix.CommandText = "SELECT ShiptoSuffix FROM ARIBA_CustomerInfo WHERE (Ext_BuyerVATID LIKE '" + ExtBuyerVatID + "%' AND ST_PA_City = '" + STPACity + "' AND ST_PA_PostalCode = '" + STPAPostalCode + "' AND KeyWord1 = '" + KeyWord(0) + "' AND KeyWord2 = '" + KeyWord(1) + " 'AND KeyWord3 = '" + KeyWord(2) + "')"
                    cmdSelectQuerySuffix.Connection = ConB2B2
                    SuffixSql = cmdSelectQuerySuffix.ExecuteScalar()
                Catch ex As Exception
                    ErrorString = ErrorString + "-Error Extincts"
                    ErrorFound = True
                End Try

            Else
                SuffixSql = ""
            End If




            ConB2B2.Close()

            ' ____ TIMESTAMPS ____ '
            Dim MessageTS As String()
            Dim OrderTS As String()
            Dim MessageTSInsert As String
            Dim ORderTSInsert As String
            Dim Node_MessageTS, Node_OrderTS As XmlNode
            Try
                Node_MessageTS = xmlCXMLNode.Attributes("timestamp")
                MessageTS = GetInnerText(Node_MessageTS).Split("+")
                MessageTSInsert = MessageTS(0)
                Node_OrderTS = xmlCXMLNode.SelectSingleNode("/cXML/Request/OrderRequest/OrderRequestHeader").Attributes("orderDate")
                OrderTS = GetInnerText(Node_OrderTS).Split("+")
                ORderTSInsert = OrderTS(0)
                MessageTSInsert = MessageTSInsert.Replace("T", " ")
                ORderTSInsert = ORderTSInsert.Replace("T", " ")
            Catch ex As Exception
                ErrorString = ErrorString + "-Error Timestamps"
                ErrorFound = True
            End Try

            ' ____ ITEMOUT ____ '
            Dim IO_LineNumber, IO_Quantity, IO_IID_SupplierPartID, IO_IID_BuyerPartID, IO_ID_UP_Money, IO_ID_Description As String
            Dim IO_RequestedDeliveryDate As String()
            Dim IO_RequestedDeliveryDateInsert As String
            Dim IO_QuantityInt As Integer
            Dim IO_QuantityArr As String()
            Dim Node_IO_LineNumber, Node_IO_Quantity, Node_IO_RequestedDeliverYDate, Node_IO_IID_SupplierPartID, Node_IO_IID_BuyerPartID, Node_IO_ID_UP_Money, Node_IO_ID_Descrption As XmlNode
            Dim ItemOutNodes = xmlCXMLNode.SelectNodes("/cXML/Request/OrderRequest/ItemOut")
            Dim xmlProductLineText As String = ""

            For Each Item In ItemOutNodes

                Try
                    Node_IO_LineNumber = Item.Attributes("lineNumber")
                    IO_LineNumber = GetInnerText(Node_IO_LineNumber)
                    Node_IO_Quantity = Item.Attributes("quantity")
                    IO_Quantity = GetInnerText(Node_IO_Quantity)
                    Node_IO_RequestedDeliverYDate = Item.Attributes("requestedDeliveryDate")
                    IO_RequestedDeliveryDate = GetInnerText(Node_IO_RequestedDeliverYDate).Split("+")
                    IO_RequestedDeliveryDateInsert = IO_RequestedDeliveryDate(0).Replace("T", " ")
                    Node_IO_IID_SupplierPartID = Item.SelectSingleNode("ItemID/SupplierPartID")
                    IO_IID_SupplierPartID = GetInnerText(Node_IO_IID_SupplierPartID)
                    'IO_IID_SupplierPartID = IO_IID_SupplierPartID.Replace("#", "Ñ")
                    Node_IO_IID_BuyerPartID = Item.SelectSingleNode("ItemID/BuyerPartID")
                    IO_IID_BuyerPartID = GetInnerText(Node_IO_IID_BuyerPartID)
                    Node_IO_ID_UP_Money = Item.SelectSingleNode("ItemDetail/UnitPrice/Money")
                    IO_ID_UP_Money = GetInnerText(Node_IO_ID_UP_Money)
                    Node_IO_ID_Descrption = Item.SelectSingleNode("ItemDetail/Description")
                    IO_ID_Description = GetInnerText(Node_IO_ID_Descrption)
                Catch ex As Exception
                    'Response.Write("Error ItemOut")
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error ItemOut"
                End Try

                IO_QuantityArr = Split(IO_Quantity, ".")
                IO_QuantityInt = Convert.ToDecimal((IO_QuantityArr(0)))

                ' ____ SEND PNA ____ '
                Try
                    ConB2B2.Open()
                Catch ex As Exception
                    'Response.Write("Error conectando B2B")
                End Try
                ' __ Finding SKU __ '
                Dim cmdSelectQueryB2BLine As New SqlCommand
                cmdSelectQueryB2BLine.CommandText = "SELECT * FROM SKU_VPN WHERE VPN LIKE '" + IO_IID_SupplierPartID + "%'"

                cmdSelectQueryB2BLine.Connection = ConB2B2
                Try
                    PNASku = cmdSelectQueryB2BLine.ExecuteScalar()
                Catch ex As Exception
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error instertando LineInfo"
                End Try

                Dim gPNA As Guid
                gPNA = Guid.NewGuid()
                Dim xmlPNAReq As New XmlDocument()
                Dim xmlText As String

                '__Genero el .XML como un string
                xmlText = "<PNARequest>" & vbCrLf
                xmlText = xmlText & "<Version>2.0</Version>" & vbCrLf
                xmlText = xmlText & "<TransactionHeader>" & vbCrLf
                xmlText = xmlText & "<SenderID>IM CUSTOMER</SenderID>" & vbCrLf
                xmlText = xmlText & "<ReceiverID>INGRAM MICRO</ReceiverID>" & vbCrLf
                xmlText = xmlText & "<CountryCode>ES</CountryCode>" & vbCrLf
                xmlText = xmlText & "<LoginID>Carbo12345</LoginID>" & vbCrLf
                xmlText = xmlText & "<Password>29080707xM</Password>" & vbCrLf
                xmlText = xmlText & "<TransactionID>{14B27C08-ADA6-406B-B7E3-9993A5408C2A}</TransactionID>" & vbCrLf
                xmlText = xmlText & "</TransactionHeader>" & vbCrLf
                xmlText = xmlText & "<PNAInformation SKU='" + PNASku + "' Quantity='1' ReservedInventory='N'/>" & vbCrLf
                xmlText = xmlText & "<ShowDetail>2</ShowDetail>" & vbCrLf
                xmlText = xmlText & "</PNARequest>" & vbCrLf

                xmlPNAReq.PreserveWhitespace = True
                xmlPNAReq.LoadXml(xmlText)
                Dim wdata As String = WRequest("https://newport.ingrammicro.com/im-xml", "POST", xmlPNAReq.OuterXml)



                Dim ResponsePNAReq As New XmlDocument()
                Dim IM_Price, IM_AvailableQty, IM_Description As String
                Dim ErrorPNA As Boolean

                ResponsePNAReq.LoadXml(wdata)
                Node_PNASku = ResponsePNAReq.SelectSingleNode("/PNAResponse/PriceAndAvailability/SKU")
                'PNASku = GetInnerText(Node_PNASku)
                If (String.IsNullOrWhiteSpace(PNASku)) Then
                    ErrorPNA = True
                Else
                    ErrorPNA = False
                End If
                IM_Price = GetInnerText(ResponsePNAReq.SelectSingleNode("/PNAResponse/PriceAndAvailability/Price"))
                IM_AvailableQty = GetInnerText(ResponsePNAReq.SelectSingleNode("/PNAResponse/PriceAndAvailability/Quantity"))
                IM_Description = GetInnerText(ResponsePNAReq.SelectSingleNode("/PNAResponse/PriceAndAvailability/Description"))
                IM_Price = IM_Price.Replace(",", ".")



                ' ____ INSERT LINE INFO ____ '
                Dim cmdInsertQueryB2BLine As New SqlCommand
                cmdInsertQueryB2BLine.CommandText = "INSERT INTO ARIBA_OrderRequest_LineInfo (From_NID_Identity, From_SID_Identity, Ext_BuyerVATID, OrderID, IO_LineNUmber, IO_Quantity, IO_RequestDeliveryDate, IO_IID_SupplierPartID, IO_IID_BuyerPartID, IO_ID_UP_Money, IO_ID_Description, IM_SKU, IM_Price, IM_AvailableQty, IM_Description, ErrorPNA) VALUES ('" + From_NID + "','" + From_SID + "','" + ExtBuyerVatID + "','" + OrderID + "','" + IO_LineNumber + "','" + IO_QuantityInt.ToString + "','" + IO_RequestedDeliveryDateInsert + "','" + IO_IID_SupplierPartID + "','" + IO_IID_BuyerPartID + "','" + IO_ID_UP_Money + "','" + IO_ID_Description + "', '" + PNASku + "','" + IM_Price + "','" + IM_AvailableQty + "', '" + IM_Description + "', '" + ErrorPNA.ToString + "')"

                cmdInsertQueryB2BLine.Connection = ConB2B2
                Try
                    cmdInsertQueryB2BLine.ExecuteNonQuery()
                Catch ex As Exception
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error instertando LineInfo"
                End Try


                '____ CONSTRUIR XML ____'
                xmlProductLineText = xmlProductLineText & "<ProductLine>"
                xmlProductLineText = xmlProductLineText & "<SKU>" + PNASku + "</SKU>"
                xmlProductLineText = xmlProductLineText & "<Quantity>" + Convert.ToString(IO_QuantityInt) + "</Quantity>"
                xmlProductLineText = xmlProductLineText & "<CustomerLineNumber>" + IO_LineNumber + "</CustomerLineNumber>"
                xmlProductLineText = xmlProductLineText & "</ProductLine>"


            Next





            ' ____ INSERT ORDER HEADER ____ '
            Dim cmdInsertQueryB2BHeader As New SqlCommand
            cmdInsertQueryB2BHeader.Connection = ConB2B

            cmdInsertQueryB2BHeader.CommandText = "UPDATE ARIBA_OrderRequest_Header SET From_NID_Identity = '" + From_NID + "' , From_SID_Identity = '" + From_SID + "', To_NID_Identity ='" + To_NID + "', OrderID ='" + OrderID + "',OrderType='" + OrderType + "', Total_Money='" + TotalMoney + "', ST_AddressID = '" + STAddressID + "', ST_Name='" + STName + "', ST_PA_Name='" + STPAName + "', ST_PA_DelvierTo='" + STPADeliverTo + "', ST_PA_Street ='" + STPAStreet + "' , ST_PA_City='" + STPACity + "', ST_PA_State='" + STPAState + "', ST_PA_PostalCode = '" + STPAPostalCode + "', ST_PA_Country='" + STPACountry + "', Ext_CompanyCode='" + ExtCompanyCode + "' , Ext_BuyerPurchasingCode='" + ExtBuyerPurchasingCode + "', Ext_VendorIDNbr='" + ExtVendorIDNbr + "', Ext_BuyerVATID='" + ExtBuyerVatID + "', ReceptionDateTime = DEFAULT, MessageTimeStamp = '" + MessageTSInsert + "', OrderTimestamp = '" + ORderTSInsert + "' WHERE OrderID = '" + OrderID + "'"
            Try
                cmdInsertQueryB2BHeader.ExecuteNonQuery()
            Catch ex As Exception
                'Response.Write("Error insertando OrderHeader")
                ErrorFound = True
                ErrorString = ErrorString + "-Error insertando Orderheader"
            End Try
            ConB2B.Close()

            ' ____ INSERT XML DB ____ '
            Dim ReaderXML As XmlTextReader
            Dim ParametroSQLXML As New SqlXml
            Try
                ConB2B.Open()
                Using cmdXML As SqlCommand = ConB2B.CreateCommand()
                    cmdXML.CommandText = "UPDATE ARIBA_OrderRequest_Header SET OriginalOrderXML = @XMLFile WHERE OrderID = '" + OrderID + "'"
                    'Dim xmlPath As String = "C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml" 'Producció
                    Dim xmlPath As String = "N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml" 'Local
                    ReaderXML = New XmlTextReader(xmlPath)
                    ParametroSQLXML = New SqlXml(ReaderXML)
                    cmdXML.Parameters.AddWithValue("@XMLFile", ParametroSQLXML)
                    cmdXML.ExecuteNonQuery()
                    cmdXML.Dispose()
                    ReaderXML.Close()
                End Using
            Catch ex As Exception
                ErrorFound = True
                ErrorString = ErrorString + "-Error insertando SQL"
            End Try

            ConB2B.Close()



            Dim xmlOrderType As New XmlDocument()
            Dim Plantilla As String = ""
            Try
                ConB2B.Open()
            Catch ex As Exception
                'Response.Write("Error conectando DB")
            End Try

            ' ____ AUTENTICACIÓN Y FF CHECK____ '
            Dim Auth As Boolean
            Auth = True
            Dim queryFFCheck As String = "SELECT * FROM ARIBA_CustomerInfo WHERE From_SID_Identity='" + From_SID + "' AND FROM_NID_Identity='" + From_NID + "' AND Ext_BuyerVATID = '" + ExtBuyerVatID + "'"
            Dim cmdQueryFFCheck As SqlCommand = New SqlCommand(queryFFCheck, ConB2B)
            Dim readerQueryFFCheck As SqlDataReader = cmdQueryFFCheck.ExecuteReader()
            While (readerQueryFFCheck.Read())
                If ((String.Compare(STPAStreet, readerQueryFFCheck("ST_PA_Street").ToString.Trim()) = 0) And (String.Compare(STPACity, readerQueryFFCheck("ST_PA_City").ToString.Trim()) = 0) And (String.Compare(STPAPostalCode, readerQueryFFCheck("ST_PA_PostalCode").ToString.Trim())) = 0) Then
                    Plantilla = "NOFF"
                Else
                    Plantilla = "FF"
                End If
                If (readerQueryFFCheck.RecordsAffected < 1) Then
                    'Auth = False
                End If
                If Auth Then

                    ' ____ ELECCIÓN PLANTILLA ____ '
                    'Select Case Plantilla

                    'Case "NOFF" 'Se ha detectado Fullfillment
                    Dim g As Guid
                    g = Guid.NewGuid()
                    ' ____ Header ____ '

                    'xmlOrderType.Load("N:\eCommerce\Aplicaciones WEB\Facturas XML\XMLAribaApp\Plantillas\Order Request Transaction.xml")
                    'xmlOrderType.SelectSingleNode("/OrderRequest/TransactionHeader/SenderID").InnerText = readerQueryFFCheck("IM_SenderID").ToString.Trim()
                    'xmlOrderType.SelectSingleNode("/OrderRequest/TransactionHeader/ReceiverID").InnerText = readerQueryFFCheck("IM_ReceiverID").ToString.Trim()
                    'xmlOrderType.SelectSingleNode("/OrderRequest/TransactionHeader/CountryCode").InnerText = readerQueryFFCheck("IM_CountryCode").ToString.Trim()
                    'xmlOrderType.SelectSingleNode("/OrderRequest/TransactionHeader/LoginID").InnerText = readerQueryFFCheck("IM_Login").ToString.Trim()
                    'xmlOrderType.SelectSingleNode("/OrderRequest/TransactionHeader/Password").InnerText = readerQueryFFCheck("IM_Password").ToString.Trim()
                    '    xmlOrderType.SelectSingleNode("/OrderRequest/TransactionHeader/TransactionID").InnerText = "{" + g.ToString.ToUpper + "}"





                    Try
                        'xmlOrderType.Save("N:\eCommerce\Aplicaciones WEB\Facturas XML\XMLAribaApp\Orders\" + readerQueryFFCheck("BrCustNbr") + "-" + g.ToString.ToUpper + ".xml")
                    Catch ex As Exception
                        'Response.Write("Error guardando el archivo")
                    End Try

                    ' End Select

                Else
                    'Escruire mail  auth = false
                End If

            End While
            readerQueryFFCheck.Close()
            ConB2B.Close()

            ' ____ XML IM ORDER ENTRY 2.0 CONSTRUCTOR ____ '
            Dim xmlOrderEntryText As String
            Dim xmlOrderEntry As New XmlDocument()
            xmlOrderEntryText = "<OrderRequest>"
            xmlOrderEntryText = xmlOrderEntryText & "<Version>2.0</Version>"
            xmlOrderEntryText = xmlOrderEntryText & "<TransactionHeader>"
            xmlOrderEntryText = xmlOrderEntryText & "<SenderID>IM CUSTOMER</SenderID>"
            xmlOrderEntryText = xmlOrderEntryText & "<ReceiverID>INGRAM MICRO</ReceiverID>"
            xmlOrderEntryText = xmlOrderEntryText & "<CountryCode>ES</CountryCode>"
            xmlOrderEntryText = xmlOrderEntryText & "<LoginID>" + User + "</LoginID>"
            xmlOrderEntryText = xmlOrderEntryText & "<Password>" + Password + "</Password>"
            xmlOrderEntryText = xmlOrderEntryText & "<TransactionID>{" + gSave.ToString + "}</TransactionID>"
            xmlOrderEntryText = xmlOrderEntryText & "</TransactionHeader>"
            xmlOrderEntryText = xmlOrderEntryText & "<OrderHeaderInformation>"
            xmlOrderEntryText = xmlOrderEntryText & " <BillToSuffix>000</BillToSuffix>"
            xmlOrderEntryText = xmlOrderEntryText & "<AddressingInformation>"
            xmlOrderEntryText = xmlOrderEntryText & "<CustomerPO>" + OrderID + "</CustomerPO>"
            xmlOrderEntryText = xmlOrderEntryText & "<ShipToAttention>AribaTest</ShipToAttention>"
            xmlOrderEntryText = xmlOrderEntryText & "<ShipTo>"
            If (FirstKeyWord <> -1 Or SecondKeyWord <> -1 Or ThirdKeyWord <> -1) Then
                xmlOrderEntryText = xmlOrderEntryText & "<Address>"
                xmlOrderEntryText = xmlOrderEntryText & "<ShipToAddress1>" + STPAStreet + "</ShipToAddress1>"
                xmlOrderEntryText = xmlOrderEntryText & " <ShipToAddress2></ShipToAddress2>"
                xmlOrderEntryText = xmlOrderEntryText & " <ShipToAddress3></ShipToAddress3>"
                xmlOrderEntryText = xmlOrderEntryText & " <ShipToCity>" + STPACity + "</ShipToCity>"
                xmlOrderEntryText = xmlOrderEntryText & " <ShipToProvince>" + STPAState + "</ShipToProvince>"
                xmlOrderEntryText = xmlOrderEntryText & " <ShipToPostalCode>" + STPAPostalCode + "</ShipToPostalCode>"
                xmlOrderEntryText = xmlOrderEntryText & " </Address>"
            Else
                xmlOrderEntryText = xmlOrderEntryText & "<Suffix><ShipToSuffix>2" + SuffixSql + "</ShipToSuffix> </Suffix>"
            End If
            xmlOrderEntryText = xmlOrderEntryText & " </ShipTo>"
            xmlOrderEntryText = xmlOrderEntryText & "</AddressingInformation>"
            xmlOrderEntryText = xmlOrderEntryText & "<ProcessingOptions>"
            xmlOrderEntryText = xmlOrderEntryText & "<CarrierCode>" + carrierCode + "</CarrierCode>"
            xmlOrderEntryText = xmlOrderEntryText & "<AutoRelease>0</AutoRelease>"
            xmlOrderEntryText = xmlOrderEntryText & "<ShipmentOptions>"
            xmlOrderEntryText = xmlOrderEntryText & "<BackOrderFlag>Y</BackOrderFlag>"
            xmlOrderEntryText = xmlOrderEntryText & "<SplitShipmentFlag>Y</SplitShipmentFlag>"
            xmlOrderEntryText = xmlOrderEntryText & "<SplitLine>N</SplitLine>"
            xmlOrderEntryText = xmlOrderEntryText & "<ShipFromBranches/>"
            xmlOrderEntryText = xmlOrderEntryText & "</ShipmentOptions>"
            xmlOrderEntryText = xmlOrderEntryText & "</ProcessingOptions>"
            xmlOrderEntryText = xmlOrderEntryText & "</OrderHeaderInformation>"
            xmlOrderEntryText = xmlOrderEntryText & "<OrderLineInformation>"
            xmlOrderEntryText = xmlOrderEntryText & xmlProductLineText
            xmlOrderEntryText = xmlOrderEntryText & "<CommentLine>"
            xmlOrderEntryText = xmlOrderEntryText & "<CommentText>///Tests Integracion ARIBA</CommentText>"
            xmlOrderEntryText = xmlOrderEntryText & "</CommentLine>"
            xmlOrderEntryText = xmlOrderEntryText & "</OrderLineInformation>"
            xmlOrderEntryText = xmlOrderEntryText & "<ShowDetail>2</ShowDetail>"
            xmlOrderEntryText = xmlOrderEntryText & "</OrderRequest>"


            xmlOrderEntry.PreserveWhitespace = True
            xmlOrderEntry.LoadXml(xmlOrderEntryText)

            Dim wdataOrderEntry = WRequest("https://newport.ingrammicro.com/im-xml", "POST", xmlOrderEntry.OuterXml)
            Dim LoadXMLErrorOrderheader As New XmlDocument()
            LoadXMLErrorOrderheader.LoadXml(wdataOrderEntry)
            Dim ErrorOrderNodeStatus As XmlNode
            Dim ErrorOrderNodeText As XmlNode
            Dim ErrSt, ErrTx As String
            Try
                ErrorOrderNodeStatus = LoadXMLErrorOrderheader.SelectSingleNode("/OrderResponse/TransactionHeader/ErrorStatus").Attributes("ErrorNumber")
                ErrorOrderNodeText = LoadXMLErrorOrderheader.SelectSingleNode("/OrderResponse/TransactionHeader/ErrorStatus")
                ErrSt = GetInnerXml(ErrorOrderNodeStatus)
                ErrTx = GetInnerText(ErrorOrderNodeText)
                If Not (String.IsNullOrWhiteSpace(ErrSt) And String.IsNullOrWhiteSpace(ErrTx)) Then
                    ErrorString = ErrorString + "- " + ErrSt + " - " + ErrTx
                    ErrorFound = True
                End If

            Catch ex As Exception
                ErrorFound = True
                ErrorString = ErrorString + "- Error POST Order"
            End Try

            ' Response.Write("Transacción realizada correctamente")

            '____ Guardar Error DB ____' 
            Dim ReaderXML2, ReaderXML3 As XmlTextReader
            Dim ParametroSQLXML2, ParametroSQLXML3 As New SqlXml

            If (ErrorFound) Then
                ' ____ INSERT STATUS ____ '

                cmdInsertStatus.Connection = ConB2B

                cmdInsertQueryB2BHeader.CommandText = "UPDATE ARIBA_OrderRequest_Header SET Status = 'Error' WHERE OrderID = '" + OrderID + "'"
                ConB2B.Open()
                Try
                    cmdInsertQueryB2BHeader.ExecuteNonQuery()
                Catch ex As Exception
                    'Response.Write("Error insertando OrderHeader")
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error insertando Orderheader"
                End Try
                ConB2B.Close()
                Try
                    ConB2B.Open()
                    Using cmdXML As SqlCommand = ConB2B.CreateCommand()
                        cmdXML.CommandText = "INSERT INTO ARIBA_cXML_ERRORs (OrderID, OriginalOrderXML_ERROR, ErrorTimeStamp, ErrorString) VALUES('" + OrderID + "', @XMLFILE,DEFAULT,'" + ErrorString + "')"
                        'Dim xmlPath As String = "C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml" 'Producció
                        Dim xmlPath As String = "N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml" 'Local
                        ReaderXML2 = New XmlTextReader(xmlPath)
                        ParametroSQLXML2 = New SqlXml(ReaderXML2)
                        cmdXML.Parameters.AddWithValue("@XMLFile", ParametroSQLXML2)
                        cmdXML.ExecuteNonQuery()
                        cmdXML.Dispose()
                        ReaderXML2.Close()
                    End Using
                Catch ex As Exception
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error insertando StatusError"
                End Try

                ConB2B.Close()

            Else

                ConB2B.Open()
                cmdInsertStatus.Connection = ConB2B

                cmdInsertQueryB2BHeader.CommandText = "UPDATE ARIBA_OrderRequest_Header SET Status = 'Successful' WHERE OrderID = '" + OrderID + "'"
                LoadXMLErrorOrderheader.Save("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\XML-OrdersResponseImpulse\" + gSave.ToString + "-RESPONSE.xml") 'Local

                Try

                    Using cmdXML As SqlCommand = ConB2B.CreateCommand()
                        cmdXML.CommandText = "UPDATE ARIBA_OrderRequest_Header SET OrderResponseImpulse = @XMLFILE2 WHERE OrderID = '" + OrderID + "'"
                        'Dim xmlPath As String = "C:\inetpub\wwwroot\WebApps\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml" 'Producció
                        Dim xmlPath As String = "N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\XML-OrdersResponseImpulse\" + gSave.ToString + "-RESPONSE.xml" 'Local
                        ReaderXML3 = New XmlTextReader(xmlPath)
                        ParametroSQLXML3 = New SqlXml(ReaderXML3)
                        cmdXML.Parameters.AddWithValue("@XMLFILE2", ParametroSQLXML3)
                        cmdXML.ExecuteNonQuery()
                        cmdXML.Dispose()
                        ReaderXML3.Close()
                    End Using
                Catch ex As Exception
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error insertando XML- Response"
                End Try

                Try
                    cmdInsertQueryB2BHeader.ExecuteNonQuery()
                Catch ex As Exception
                    'Response.Write("Error insertando OrderHeader")
                    ErrorFound = True
                    ErrorString = ErrorString + "-Error insertando StatusSuccessful"
                End Try
                ConB2B.Close()
            End If
        Else

            Response.Write("No se ha recibido nada")

        End If

        '____ MAIL ERROR ____'

        'START Config_Mail
        Dim Mail As New System.Net.Mail.MailMessage
        Mail.From = New System.Net.Mail.MailAddress("DoNotReply@ingrammicro.es")
        Mail.To.Add("ernest.espinola@ingrammicro.com")
        'Mail.Bcc.Add("jordi.carbo@ingrammicro.com")
        Mail.Subject = "Errores Ariba" + OrderID
        Mail.IsBodyHtml = True
        Dim smtp As New System.Net.Mail.SmtpClient
        smtp.Host = "172.31.16.50" 'Mail Relay de pruebas para Visual Studio va de PM! local
        'smtp.Port = 25 ' Producción
        'END Config_Mail
        Dim attachment As System.Net.Mail.Attachment


        If (ErrorFound) Then
            Mail.Body = "Se ha detectado un error para el OrderID de Ariba: " + OrderID + ". El motivo de el/los error(es) és el siguiente: " + ErrorString + ". Para cualquier clarificación contactar a <a href=""mailto:SoporteWeb@ingrammicro.com"">SoporteWeb</a>"
            Try
                attachment = New System.Net.Mail.Attachment("N:\eCommerce\Aplicaciones WEB\ARIBA_Integration\cXML-OrdersAriba\" + gSave.ToString + ".xml")
                Mail.Attachments.Add(attachment)
            Catch ex As Exception
                Response.Write("Error adjuntando archivo")
            End Try

            Try
                smtp.Send(Mail)
            Catch ex As Exception
                Response.Write("Error enviando mail")
            End Try
        End If



    End Sub


    Private Function GetInnerText(node As XmlNode) As String
        If node Is Nothing Then Return ""
        Return node.InnerText

    End Function

    Private Function GetInnerInt(node As String) As Integer
        Dim node2 As Integer
        node2 = CInt(node)
        Return (node2)

    End Function

    Private Function GetInnerXml(node As XmlNode) As String
        If node Is Nothing Then Return ""
        Return node.InnerXml
    End Function

    Private Function SetInnerText(node As XmlNode, StrValue As String) As String
        If node Is Nothing Then Return ""
        node.InnerText = StrValue
        Return StrValue
    End Function

    Function WRequest(URL As String, method As String, POSTdata As String) As String
        Dim responseData As String = ""
        Try
            Dim hwrequest As Net.HttpWebRequest = Net.WebRequest.Create(URL)
            hwrequest.Accept = "*/*"
            hwrequest.AllowAutoRedirect = True
            hwrequest.UserAgent = "http_requester/0.1"
            hwrequest.Timeout = 180000
            hwrequest.Method = method
            If hwrequest.Method = "POST" Then
                hwrequest.ContentType = "application/x-www-form-urlencoded"
                hwrequest.ContentLength = POSTdata.Length
                Dim writer As New StreamWriter(hwrequest.GetRequestStream, Encoding.Default, POSTdata.Length)
                writer.Write(POSTdata)
                writer.Close()
            End If
            Dim hwresponse As Net.HttpWebResponse = hwrequest.GetResponse()
            If hwresponse.StatusCode = Net.HttpStatusCode.OK Then
                Dim responseStream As IO.StreamReader = New IO.StreamReader(hwresponse.GetResponseStream())
                responseData = responseStream.ReadToEnd()
            End If
            hwresponse.Close()
        Catch e As Exception
            responseData = "ERROR: " & e.Message
        End Try
        Return responseData
    End Function


End Class
