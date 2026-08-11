Attribute VB_Name = "modKDS"
Public Sub KDS_AnadirNuevaOrden(ByVal kdsRsCabecera As Recordset, ByVal accion As Integer)
On Error GoTo ErrorKDSHandler

    If sOrderInfox = "" Then 'SI SOLO TIENE UN MODELO DE KDS

       Dim Documento As MSXML2.DOMDocument60 ' . DOMDocument60
       Set Documento = New DOMDocument60 'DOMDocument60
       '<Transaction>
       Documento.loadXML ("<Transaction></Transaction>")
       
           '<Order>
           Dim nodoOrder As MSXML2.IXMLDOMNode
           Set nodoOrder = Documento.createElement("Order")
           Documento.documentElement.appendChild nodoOrder
               '<ID>
               Dim nodoID As MSXML2.IXMLDOMNode
               Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
               nodoID.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                   '<PosTerminal>
                   Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                   Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                   nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                   '<TransType>
                   Dim nodoTransType As MSXML2.IXMLDOMNode
                   Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                   nodoTransType.Text = "1"
                   '<OrderStatus>
                   Dim nodoOrderStatus As MSXML2.IXMLDOMNode
                   Set nodoOrderStatus = nodoOrder.appendChild(Documento.createElement("OrderStatus"))
                   nodoOrderStatus.Text = KDS_Obtener_OrderStatus() 'Cambiar
                   '<OrderType>
                   Dim nodoOrderType As MSXML2.IXMLDOMNode
                   Set nodoOrderType = nodoOrder.appendChild(Documento.createElement("OrderType"))
                   nodoOrderType.Text = KDS_Obtener_OrderType() 'Cambiar
                   '<ServerName>
                   Dim nodoServerName As MSXML2.IXMLDOMNode
                   Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                   nodoServerName.Text = kdsRsCabecera!tObservacion
                   '<Destination>
                   Dim nodoDestination As MSXML2.IXMLDOMNode
                   Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                   nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                   '<GuestTable>
                   Dim nodoGuestTable As MSXML2.IXMLDOMNode
                   Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                   nodoGuestTable.Text = KDS_Obtener_Mesa(kdsRsCabecera!tMesa)
                   '<UserInfo>
                   Dim nodoUserInfo As MSXML2.IXMLDOMNode
                   Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
                   
                   Dim kdsRsDetalle As Recordset
                   Set kdsRsDetalle = KDS_ObtenerDetallePedido(kdsRsCabecera!codigo, accion)
                                   
               Do While Not kdsRsDetalle.EOF
                   '<Item>
                   Dim nodoItem As MSXML2.IXMLDOMNode
                   Set nodoItem = nodoOrder.appendChild(Documento.createElement("Item"))
                       '<ID>
                       Dim nodoItemID As MSXML2.IXMLDOMNode
                       Set nodoItemID = nodoItem.appendChild(Documento.createElement("ID"))
                       If (IsNull(kdsRsDetalle!tItemCombo)) Then
                           nodoItemID.Text = val(kdsRsDetalle!tItem)
                       Else
                           nodoItemID.Text = val(kdsRsDetalle!tItem) * 100 + val(kdsRsDetalle!tItemCombo)
                       End If
                       '<TransType>
                       Dim nodoItemTransType As MSXML2.IXMLDOMNode
                       Set nodoItemTransType = nodoItem.appendChild(Documento.createElement("TransType"))
                       nodoItemTransType.Text = "1"
                       '<Name>
                       Dim nodoItemName As MSXML2.IXMLDOMNode
                       Set nodoItemName = nodoItem.appendChild(Documento.createElement("Name"))
                       
                       If Not kdsRsDetalle!lCombinacion Then
                           nodoItemName.Text = kdsRsDetalle!Producto
                       Else
                           Dim NombreTemp As String
    
                           NombreTemp = KDS_Obtener_InicialesDeNombre(kdsRsDetalle!Combo)
                           nodoItemName.Text = NombreTemp + kdsRsDetalle!Producto
                       End If
                       
                       '<Category>
                       Dim nodoItemCategory As MSXML2.IXMLDOMNode
                       Set nodoItemCategory = nodoItem.appendChild(Documento.createElement("Category"))
                       nodoItemCategory.Text = "Monitor1" ' + kdsRsDetalle!tCodigoProducto
                       '<Quantity>
                       Dim nodoItemQuantity As MSXML2.IXMLDOMNode
                       Set nodoItemQuantity = nodoItem.appendChild(Documento.createElement("Quantity"))
                       nodoItemQuantity.Text = kdsRsDetalle!nCantidad
                       '<Color>
                       Dim nodoItemColor As MSXML2.IXMLDOMNode
                       Set nodoItemColor = nodoItem.appendChild(Documento.createElement("Color"))
                     
                       '<KDSStation>
                       Dim nodoItemKDSStation As MSXML2.IXMLDOMNode
                       Set nodoItemKDSStation = nodoItem.appendChild(Documento.createElement("KDSStation"))
                       nodoItemKDSStation.Text = val(KDS_ObtenerAreaImpresionKDS(kdsRsDetalle!tCodigoProducto, IIf(IsNull(kdsRsDetalle!tItemCombo), "", kdsRsDetalle!tItemCombo), kdsRsDetalle!tCodigoPedido, kdsRsDetalle!tItem))
                       
                       Dim kdsRsProductoPropiedad As Recordset
                       Set kdsRsProductoPropiedad = KDS_ObtenerPropiedadesProducto(kdsRsDetalle!tCodigoPedido, kdsRsDetalle!tItem, IIf(IsNull(kdsRsDetalle!tItemCombo), "", kdsRsDetalle!tItemCombo), kdsRsDetalle!tCodigoProducto)
                       
                       Do While Not kdsRsProductoPropiedad.EOF
                           '<Condiment>
                           Dim nodoItemCondiment As MSXML2.IXMLDOMNode
                           Set nodoItemCondiment = nodoItem.appendChild(Documento.createElement("Condiment"))
                               '<ID>
                               Dim nodoItemCondimentID As MSXML2.IXMLDOMNode
                               Set nodoItemCondimentID = nodoItemCondiment.appendChild(Documento.createElement("ID"))
                               nodoItemCondimentID.Text = val(kdsRsProductoPropiedad!tCodigoPropiedad)
                               '<TransType>
                               Dim nodoItemCondimentTransType As MSXML2.IXMLDOMNode
                               Set nodoItemCondimentTransType = nodoItemCondiment.appendChild(Documento.createElement("TransType"))
                               nodoItemCondimentTransType.Text = "1"
                               '<Name>
                               Dim nodoItemCondimentName As MSXML2.IXMLDOMNode
                               Set nodoItemCondimentName = nodoItemCondiment.appendChild(Documento.createElement("Name"))
                               If (kdsRsProductoPropiedad!tCodigoPropiedad <> "9999") Then
                                   nodoItemCondimentName.Text = kdsRsProductoPropiedad!Operador + kdsRsProductoPropiedad!Propiedad
                               Else
                                   nodoItemCondimentName.Text = KDS_ObtenerInfoPropiedadProducto(kdsRsProductoPropiedad!tEnlace, kdsRsProductoPropiedad!tProducto, kdsRsProductoPropiedad!tCodigoPropiedad)
                               End If
                               '<Color>
                               Dim nodoItemCondimentColor As MSXML2.IXMLDOMNode
                               Set nodoItemCondimentColor = nodoItemCondiment.appendChild(Documento.createElement("Color"))
                               '<Action>
                               Dim nodoItemCondimentAction As MSXML2.IXMLDOMNode
                               Set nodoItemCondimentAction = nodoItemCondiment.appendChild(Documento.createElement("Action"))
                               kdsRsProductoPropiedad.MoveNext
                       Loop
                   kdsRsDetalle.MoveNext
               Loop
               
               
       'DONDE GUARDAR EL XML
       Dim direccionArchivo As String
       If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
           sOrderInfo = sOrderInfo + "\"
       End If
       Documento.Save (sOrderInfo & nodoID.Text & ".xml")
       
       
    Else ' SI TIENE MAS DE UN MODELO DE KDS
             Dim kdsSaber As Recordset
             Set kdsSaber = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedidox '" & kdsRsCabecera!codigo & "', '" & 2 & "', '" & xkds & "'", Cn)
              
        Do While Not kdsSaber.EOF

            Dim Documentox As MSXML2.DOMDocument60 ' . DOMDocument60
            Set Documentox = New DOMDocument60 'DOMDocument60
            '<Transaction>
            Documentox.loadXML ("<Transaction></Transaction>")
            
                '<Order>
                Dim nodoOrderx As MSXML2.IXMLDOMNode
                Set nodoOrderx = Documentox.createElement("Order")
                Documentox.documentElement.appendChild nodoOrderx
                    '<ID>
                    Dim nodoIDx As MSXML2.IXMLDOMNode
                    Set nodoIDx = nodoOrderx.appendChild(Documentox.createElement("ID"))
                    nodoIDx.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                        '<PosTerminal>
                        Dim nodoPosTerminalx As MSXML2.IXMLDOMNode
                        Set nodoPosTerminalx = nodoOrderx.appendChild(Documentox.createElement("PosTerminal"))
                        nodoPosTerminalx.Text = KDS_Obtener_PosTerminal()
                        '<TransType>
                        Dim nodoTransTypex As MSXML2.IXMLDOMNode
                        Set nodoTransTypex = nodoOrderx.appendChild(Documentox.createElement("TransType"))
                        nodoTransTypex.Text = "1"
                        '<OrderStatus>
                        Dim nodoOrderStatusx As MSXML2.IXMLDOMNode
                        Set nodoOrderStatusx = nodoOrderx.appendChild(Documentox.createElement("OrderStatus"))
                        nodoOrderStatusx.Text = KDS_Obtener_OrderStatus() 'Cambiar
                        '<OrderType>
                        Dim nodoOrderTypex As MSXML2.IXMLDOMNode
                        Set nodoOrderTypex = nodoOrderx.appendChild(Documentox.createElement("OrderType"))
                        nodoOrderTypex.Text = KDS_Obtener_OrderType() 'Cambiar
                        '<ServerName>
                        Dim nodoServerNamex As MSXML2.IXMLDOMNode
                        Set nodoServerNamex = nodoOrderx.appendChild(Documentox.createElement("ServerName"))
                        nodoServerNamex.Text = kdsRsCabecera!tObservacion
                        '<Destination>
                        Dim nodoDestinationx As MSXML2.IXMLDOMNode
                        Set nodoDestinationx = nodoOrderx.appendChild(Documentox.createElement("Destination"))
                        nodoDestinationx.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                        '<GuestTable>
                        Dim nodoGuestTablex As MSXML2.IXMLDOMNode
                        Set nodoGuestTablex = nodoOrderx.appendChild(Documentox.createElement("GuestTable"))
                        nodoGuestTablex.Text = KDS_Obtener_Mesa(kdsRsCabecera!tMesa)
                        '<UserInfo>
                        Dim nodoUserInfox As MSXML2.IXMLDOMNode
                        Set nodoUserInfox = nodoOrderx.appendChild(Documentox.createElement("UserInfo"))
                        
                        Dim kdsRsDetallex As Recordset
                        'Set kdsRsDetallex = KDS_ObtenerDetallePedido(kdsRsCabecera!codigo, accion)
                        Set kdsRsDetallex = KDS_ObtenerDetallePedidox(kdsRsCabecera!codigo, accion, kdsSaber!KDSX)
                    
                   Do While Not kdsRsDetallex.EOF
                       '<Item>
                       Dim nodoItemx As MSXML2.IXMLDOMNode
                       Set nodoItemx = nodoOrderx.appendChild(Documentox.createElement("Item"))
                           '<ID>
                           Dim nodoItemIDx As MSXML2.IXMLDOMNode
                           Set nodoItemIDx = nodoItemx.appendChild(Documentox.createElement("ID"))
                           If (IsNull(kdsRsDetallex!tItemCombo)) Then
                               nodoItemIDx.Text = val(kdsRsDetallex!tItem)
                           Else
                               nodoItemIDx.Text = val(kdsRsDetallex!tItem) * 100 + val(kdsRsDetallex!tItemCombo)
                           End If
                           '<TransType>
                           Dim nodoItemTransTypex As MSXML2.IXMLDOMNode
                           Set nodoItemTransTypex = nodoItemx.appendChild(Documentox.createElement("TransType"))
                           nodoItemTransTypex.Text = "1"
                           '<Name>
                           Dim nodoItemNamex As MSXML2.IXMLDOMNode
                           Set nodoItemNamex = nodoItemx.appendChild(Documentox.createElement("Name"))
                           
                           If Not kdsRsDetallex!lCombinacion Then
                               nodoItemNamex.Text = kdsRsDetallex!Producto
                           Else
                               Dim NombreTempx As String
        
                               NombreTempx = KDS_Obtener_InicialesDeNombre(kdsRsDetallex!Combo)
                               nodoItemNamex.Text = NombreTempx + kdsRsDetallex!Producto
                           End If
                           
                           '<Category>
                           Dim nodoItemCategoryx As MSXML2.IXMLDOMNode
                           Set nodoItemCategoryx = nodoItemx.appendChild(Documentox.createElement("Category"))
                           nodoItemCategoryx.Text = "Monitor1" ' + kdsRsDetalle!tCodigoProducto
                           '<Quantity>
                           Dim nodoItemQuantityx As MSXML2.IXMLDOMNode
                           Set nodoItemQuantityx = nodoItemx.appendChild(Documentox.createElement("Quantity"))
                           nodoItemQuantityx.Text = kdsRsDetallex!nCantidad
                           '<Color>
                           Dim nodoItemColorx As MSXML2.IXMLDOMNode
                           Set nodoItemColorx = nodoItemx.appendChild(Documentox.createElement("Color"))
                         
                           '<KDSStation>
                           Dim nodoItemKDSStationx As MSXML2.IXMLDOMNode
                           Set nodoItemKDSStationx = nodoItemx.appendChild(Documentox.createElement("KDSStation"))
                           nodoItemKDSStationx.Text = val(KDS_ObtenerAreaImpresionKDS(kdsRsDetallex!tCodigoProducto, IIf(IsNull(kdsRsDetallex!tItemCombo), "", kdsRsDetallex!tItemCombo), kdsRsDetallex!tCodigoPedido, kdsRsDetallex!tItem))
                           
                           Dim kdsRsProductoPropiedadx As Recordset
                           Set kdsRsProductoPropiedadx = KDS_ObtenerPropiedadesProducto(kdsRsDetallex!tCodigoPedido, kdsRsDetallex!tItem, IIf(IsNull(kdsRsDetallex!tItemCombo), "", kdsRsDetallex!tItemCombo), kdsRsDetallex!tCodigoProducto)
                           
                           Do While Not kdsRsProductoPropiedadx.EOF
                               '<Condiment>
                               Dim nodoItemCondimentx As MSXML2.IXMLDOMNode
                               Set nodoItemCondimentx = nodoItemx.appendChild(Documentox.createElement("Condiment"))
                                   '<ID>
                                   Dim nodoItemCondimentIDx As MSXML2.IXMLDOMNode
                                   Set nodoItemCondimentIDx = nodoItemCondimentx.appendChild(Documentox.createElement("ID"))
                                   nodoItemCondimentIDx.Text = val(kdsRsProductoPropiedadx!tCodigoPropiedad)
                                   '<TransType>
                                   Dim nodoItemCondimentTransTypex As MSXML2.IXMLDOMNode
                                   Set nodoItemCondimentTransTypex = nodoItemCondimentx.appendChild(Documentox.createElement("TransType"))
                                   nodoItemCondimentTransTypex.Text = "1"
                                   '<Name>
                                   Dim nodoItemCondimentNamex As MSXML2.IXMLDOMNode
                                   Set nodoItemCondimentNamex = nodoItemCondimentx.appendChild(Documentox.createElement("Name"))
                                   If (kdsRsProductoPropiedadx!tCodigoPropiedad <> "9999") Then
                                       nodoItemCondimentNamex.Text = kdsRsProductoPropiedadx!Operador + kdsRsProductoPropiedadx!Propiedad
                                   Else
                                       nodoItemCondimentNamex.Text = KDS_ObtenerInfoPropiedadProducto(kdsRsProductoPropiedadx!tEnlace, kdsRsProductoPropiedadx!tProducto, kdsRsProductoPropiedadx!tCodigoPropiedad)
                                   End If
                                   '<Color>
                                   Dim nodoItemCondimentColorx As MSXML2.IXMLDOMNode
                                   Set nodoItemCondimentColorx = nodoItemCondimentx.appendChild(Documentox.createElement("Color"))
                                   '<Action>
                                   Dim nodoItemCondimentActionx As MSXML2.IXMLDOMNode
                                   Set nodoItemCondimentActionx = nodoItemCondimentx.appendChild(Documentox.createElement("Action"))
                                   kdsRsProductoPropiedadx.MoveNext
                           Loop
                       kdsRsDetallex.MoveNext
                   Loop
                   
               
            'DONDE GUARDAR EL XML
            Dim direccionArchivox As String
            If kdsSaber!KDSX = 0 Then
                
                If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
                    sOrderInfo = sOrderInfo + "\"
                End If
                Documentox.Save (sOrderInfo & nodoIDx.Text & ".xml")
            Else
                If Mid(sOrderInfox, Len(sOrderInfox) - 1, 1) <> "\" Then
                    sOrderInfox = sOrderInfox + "\"
                End If
                Documentox.Save (sOrderInfox & nodoIDx.Text & ".xml")
            End If
         kdsSaber.MoveNext
        Loop
       
    End If
    
    
    
    Exit Sub
ErrorKDSHandler:
End Sub

Public Sub KDS_EliminarOrden(ByVal kdsRsCabecera As Recordset)
On Error GoTo ErrorKDSHandler
    
     If sOrderInfox = "" Then 'SI SOLO TIENE UN MODELO DE KDS
        Dim Documento As MSXML2.DOMDocument60
        Set Documento = New DOMDocument60
        '- <Transaction>
        Documento.loadXML ("<Transaction></Transaction>")
            '- <Order>
            Dim nodoOrder As MSXML2.IXMLDOMNode
            Set nodoOrder = Documento.createElement("Order")
            Documento.documentElement.appendChild nodoOrder
                '<ID>
                Dim nodoID As MSXML2.IXMLDOMNode
                Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
                nodoID.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                    '<PosTerminal>
                    Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                    Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                    nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                    '<TransType>
                    Dim nodoTransType As MSXML2.IXMLDOMNode
                    Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                    nodoTransType.Text = "2"
                    '<ServerName>
                    Dim nodoServerName As MSXML2.IXMLDOMNode
                    Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                    nodoServerName.Text = kdsRsCabecera!tUsuario
                    '<Destination>
                    Dim nodoDestination As MSXML2.IXMLDOMNode
                    Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                    nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido) 'Cambiar
                    '<GuestTable>
                    Dim nodoGuestTable As MSXML2.IXMLDOMNode
                    Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                    '<UserInfo>
                    Dim nodoUserInfo As MSXML2.IXMLDOMNode
                    Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
            '</Order>
        '</Transaction>
        Dim direccionArchivo As String
        If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
            sOrderInfo = sOrderInfo + "\"
        End If
        Documento.Save (sOrderInfo & nodoID.Text & ".xml")
    Else
        Dim kdsSaber As Recordset
        Set kdsSaber = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedidox '" & kdsRsCabecera!codigo & "', '" & 2 & "', '" & xkds & "'", Cn)
        
        
        Do While Not kdsSaber.EOF
        
        
            Dim Documentox As MSXML2.DOMDocument60
            Set Documentox = New DOMDocument60
            '- <Transaction>
            Documentox.loadXML ("<Transaction></Transaction>")
                '- <Order>
                Dim nodoOrderx As MSXML2.IXMLDOMNode
                Set nodoOrderx = Documentox.createElement("Order")
                Documentox.documentElement.appendChild nodoOrderx
                    '<ID>
                    Dim nodoIDx As MSXML2.IXMLDOMNode
                    Set nodoIDx = nodoOrderx.appendChild(Documentox.createElement("ID"))
                    nodoIDx.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                        '<PosTerminal>
                        Dim nodoPosTerminalx As MSXML2.IXMLDOMNode
                        Set nodoPosTerminalx = nodoOrderx.appendChild(Documentox.createElement("PosTerminal"))
                        nodoPosTerminalx.Text = KDS_Obtener_PosTerminal()
                        '<TransType>
                        Dim nodoTransTypex As MSXML2.IXMLDOMNode
                        Set nodoTransTypex = nodoOrderx.appendChild(Documentox.createElement("TransType"))
                        nodoTransTypex.Text = "2"
                        '<ServerName>
                        Dim nodoServerNamex As MSXML2.IXMLDOMNode
                        Set nodoServerNamex = nodoOrderx.appendChild(Documentox.createElement("ServerName"))
                        nodoServerNamex.Text = kdsRsCabecera!tUsuario
                        '<Destination>
                        Dim nodoDestinationx As MSXML2.IXMLDOMNode
                        Set nodoDestinationx = nodoOrderx.appendChild(Documentox.createElement("Destination"))
                        nodoDestinationx.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido) 'Cambiar
                        '<GuestTable>
                        Dim nodoGuestTablex As MSXML2.IXMLDOMNode
                        Set nodoGuestTablex = nodoOrderx.appendChild(Documentox.createElement("GuestTable"))
                        '<UserInfo>
                        Dim nodoUserInfox As MSXML2.IXMLDOMNode
                        Set nodoUserInfox = nodoOrderx.appendChild(Documentox.createElement("UserInfo"))
                '</Order>
            '</Transaction>
            Dim direccionArchivox As String
            
            If kdsSaber!KDSX = 0 Then
                If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
                    sOrderInfo = sOrderInfo + "\"
                End If
                Documentox.Save (sOrderInfo & nodoIDx.Text & ".xml")
            Else
                If Mid(sOrderInfox, Len(sOrderInfox) - 1, 1) <> "\" Then
                    sOrderInfox = sOrderInfox + "\"
                End If
                Documentox.Save (sOrderInfox & nodoIDx.Text & ".xml")
            End If
            
         kdsSaber.MoveNext
        Loop
        
    End If



    Exit Sub
ErrorKDSHandler:
End Sub

Public Sub KDS_EliminarProducto(ByVal kdsRsCabecera As Recordset, ByVal itItem As String, ByVal accion As Integer)
On Error GoTo ErrorKDSHandler
 If sOrderInfox = "" Then 'SI SOLO TIENE UN MODELO DE KDS
    If (KDS_ObtenerProductoPedidoImpresos(kdsRsCabecera!codigo, itItem, accion).RecordCount = 0) Then
        Exit Sub
    End If

    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<Transaction>
    Documento.loadXML ("<Transaction></Transaction>")
        '<Order>
        Dim nodoOrder As MSXML2.IXMLDOMNode
        Set nodoOrder = Documento.createElement("Order")
        Documento.documentElement.appendChild nodoOrder
            '<ID>
            Dim nodoID As MSXML2.IXMLDOMNode
            Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
            nodoID.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransType As MSXML2.IXMLDOMNode
                Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                nodoTransType.Text = "3"
                '<OrderStatus>
                Dim nodoOrderStatus As MSXML2.IXMLDOMNode
                Set nodoOrderStatus = nodoOrder.appendChild(Documento.createElement("OrderStatus"))
                nodoOrderStatus.Text = KDS_Obtener_OrderStatus() 'Cambiar
                '<OrderType>
                Dim nodoOrderType As MSXML2.IXMLDOMNode
                Set nodoOrderType = nodoOrder.appendChild(Documento.createElement("OrderType"))
                nodoOrderType.Text = KDS_Obtener_OrderType() 'Cambiar
                '<ServerName>
                Dim nodoServerName As MSXML2.IXMLDOMNode
                Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                nodoServerName.Text = kdsRsCabecera!tUsuario
                '<Destination>
                Dim nodoDestination As MSXML2.IXMLDOMNode
                Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                '<GuestTable>
                Dim nodoGuestTable As MSXML2.IXMLDOMNode
                Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                '<UserInfo>
                Dim nodoUserInfo As MSXML2.IXMLDOMNode
                Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
                
                Dim kdsRsDetalleCombo As Recordset
                Set kdsRsDetalleCombo = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "0", accion)
            If (kdsRsDetalleCombo.RecordCount = 0) Then 'si no es combo entonces
                '<Item>
                Dim nodoItem As MSXML2.IXMLDOMNode
                Set nodoItem = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID As MSXML2.IXMLDOMNode
                    Set nodoItemID = nodoItem.appendChild(Documento.createElement("ID"))
                    nodoItemID.Text = val(itItem)
                    '<TransType>
                    Dim nodoItemTransType As MSXML2.IXMLDOMNode
                    Set nodoItemTransType = nodoItem.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType.Text = "2"
            Else 'si no
              Set kdsRsDetalleCombo = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "1", accion)
              Do While Not kdsRsDetalleCombo.EOF
                '<Item>
                Dim nodoItem2 As MSXML2.IXMLDOMNode
                Set nodoItem2 = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID2 As MSXML2.IXMLDOMNode
                    Set nodoItemID2 = nodoItem2.appendChild(Documento.createElement("ID"))
                    nodoItemID2.Text = val(kdsRsDetalleCombo!tItem) * 100 + val(kdsRsDetalleCombo!tItemCombo)
                    '<TransType>
                    Dim nodoItemTransType2 As MSXML2.IXMLDOMNode
                    Set nodoItemTransType2 = nodoItem2.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType2.Text = "2"
                    kdsRsDetalleCombo.MoveNext
              Loop
            End If
    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & nodoID.Text & ".xml")
    
    Dim rsProductopedido As Recordset
    Set rsProductopedido = KDS_ObtenerProductoPedido(kdsRsCabecera!codigo, itItem)
    Dim nombreprod As String
    nombreprod = rsProductopedido!tResumido
    Do While Not rsProductopedido.EOF
        If (rsProductopedido!Area <> "") Then
            Call KDS_EnviaMensaje(rsProductopedido!Area, "(" & val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo))) & ")ELIMINADO:" & nombreprod)
        End If
        rsProductopedido.MoveNext
    Loop
Else

    Dim kdsSaber As Recordset
    Set kdsSaber = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedidoxKDS '" & kdsRsCabecera!codigo & "', '" & itItem & "'", Cn)

Do While Not kdsSaber.EOF
    If (KDS_ObtenerProductoPedidoImpresos(kdsRsCabecera!codigo, itItem, accion).RecordCount = 0) Then
        Exit Sub
    End If

    Dim Documentox As MSXML2.DOMDocument60
    Set Documentox = New DOMDocument60
    '<Transaction>
    Documentox.loadXML ("<Transaction></Transaction>")
        '<Order>
        Dim nodoOrderx As MSXML2.IXMLDOMNode
        Set nodoOrderx = Documentox.createElement("Order")
        Documentox.documentElement.appendChild nodoOrderx
            '<ID>
            Dim nodoIDx As MSXML2.IXMLDOMNode
            Set nodoIDx = nodoOrderx.appendChild(Documentox.createElement("ID"))
            nodoIDx.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminalx As MSXML2.IXMLDOMNode
                Set nodoPosTerminalx = nodoOrderx.appendChild(Documentox.createElement("PosTerminal"))
                nodoPosTerminalx.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransTypex As MSXML2.IXMLDOMNode
                Set nodoTransTypex = nodoOrderx.appendChild(Documentox.createElement("TransType"))
                nodoTransTypex.Text = "3"
                '<OrderStatus>
                Dim nodoOrderxStatusx As MSXML2.IXMLDOMNode
                Set nodoOrderxStatusx = nodoOrderx.appendChild(Documentox.createElement("OrderStatus"))
                nodoOrderxStatusx.Text = KDS_Obtener_OrderStatus() 'Cambiar
                '<OrderType>
                Dim nodoOrderxTypex As MSXML2.IXMLDOMNode
                Set nodoOrderxTypex = nodoOrderx.appendChild(Documentox.createElement("OrderType"))
                nodoOrderxTypex.Text = KDS_Obtener_OrderType() 'Cambiar
                '<ServerName>
                Dim nodoServerNamex As MSXML2.IXMLDOMNode
                Set nodoServerNamex = nodoOrderx.appendChild(Documentox.createElement("ServerName"))
                nodoServerNamex.Text = kdsRsCabecera!tUsuario
                '<Destination>
                Dim nodoDestinationx As MSXML2.IXMLDOMNode
                Set nodoDestinationx = nodoOrderx.appendChild(Documentox.createElement("Destination"))
                nodoDestinationx.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                '<GuestTable>
                Dim nodoGuestTablex As MSXML2.IXMLDOMNode
                Set nodoGuestTablex = nodoOrderx.appendChild(Documentox.createElement("GuestTable"))
                '<UserInfo>
                Dim nodoUserInfox As MSXML2.IXMLDOMNode
                Set nodoUserInfox = nodoOrderx.appendChild(Documentox.createElement("UserInfo"))
                
                Dim kdsRsDetalleCombox As Recordset
                Set kdsRsDetalleCombox = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "0", accion)
            If (kdsRsDetalleCombox.RecordCount = 0) Then 'si no es combo entonces
                '<Item>
                Dim nodoItemx As MSXML2.IXMLDOMNode
                Set nodoItemx = nodoOrderx.appendChild(Documentox.createElement("Item"))
                    '<ID>
                    Dim nodoItemIDx As MSXML2.IXMLDOMNode
                    Set nodoItemIDx = nodoItemx.appendChild(Documentox.createElement("ID"))
                    nodoItemIDx.Text = val(itItem)
                    '<TransType>
                    Dim nodoItemTransTypex As MSXML2.IXMLDOMNode
                    Set nodoItemTransTypex = nodoItemx.appendChild(Documentox.createElement("TransType"))
                    nodoItemTransTypex.Text = "2"
            Else 'si no
              Set kdsRsDetalleCombox = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "1", accion)
              Do While Not kdsRsDetalleCombox.EOF
                '<Item>
                Dim nodoItem2x As MSXML2.IXMLDOMNode
                Set nodoItem2x = nodoOrderx.appendChild(Documentox.createElement("Item"))
                    '<ID>
                    Dim nodoItemID2x As MSXML2.IXMLDOMNode
                    Set nodoItemID2x = nodoItem2x.appendChild(Documentox.createElement("ID"))
                    nodoItemID2x.Text = val(kdsRsDetalleCombox!tItem) * 100 + val(kdsRsDetalleCombox!tItemCombo)
                    '<TransType>
                    Dim nodoItemTransType2x As MSXML2.IXMLDOMNode
                    Set nodoItemTransType2x = nodoItem2x.appendChild(Documentox.createElement("TransType"))
                    nodoItemTransType2x.Text = "2"
                    kdsRsDetalleCombox.MoveNext
              Loop
            End If
            
     If kdsSaber!KDSX = 0 Then
        If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
            sOrderInfo = sOrderInfo + "\"
        End If
        Documentox.Save (sOrderInfo & nodoIDx.Text & ".xml")
    Else
        If Mid(sOrderInfox, Len(sOrderInfox) - 1, 1) <> "\" Then
            sOrderInfox = sOrderInfox + "\"
        End If
        Documentox.Save (sOrderInfox & nodoIDx.Text & ".xml")
    End If
    
    
    '/******************
    Dim rsProductopedidox As Recordset
    Set rsProductopedidox = KDS_ObtenerProductoPedido(kdsRsCabecera!codigo, itItem)
    Dim nombreprodx As String
    nombreprodx = rsProductopedidox!tResumido
    Do While Not rsProductopedidox.EOF
        If (rsProductopedidox!Area <> "") Then
            Call KDS_EnviaMensajex(rsProductopedidox!Area, "(" & val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo))) & ")ELIMINADO:" & nombreprodx, kdsSaber!KDSX)
        End If
        rsProductopedidox.MoveNext
    Loop
 kdsSaber.MoveNext
Loop
    
End If
    Exit Sub
ErrorKDSHandler:
End Sub
Private Sub KDS_EnviaMensajex(ByVal stationID As String, ByVal info As String, ByVal kds As Integer)
On Error GoTo ErrorHandler
    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<StationInfo>
    Documento.loadXML ("<StationInfo></StationInfo>")
        '<StationID>
        Dim nodoStationID As MSXML2.IXMLDOMNode
        Set nodoStationID = Documento.createElement("StationID")
        nodoStationID.Text = val(stationID)
        Documento.documentElement.appendChild nodoStationID
        '<User>
        Dim nodoUser As MSXML2.IXMLDOMNode
        Set nodoUser = Documento.createElement("User")
        nodoUser.Text = "0"
        Documento.documentElement.appendChild nodoUser
        '<Info>
        Dim nodoInfo As MSXML2.IXMLDOMNode
        Set nodoInfo = Documento.createElement("Info")
        nodoInfo.Text = info
        Documento.documentElement.appendChild nodoInfo
    
    Dim direccionArchivo As String
    
    If kds = 0 Then
        If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
            sOrderInfo = sOrderInfo + "\"
        End If
        Documento.Save (sOrderInfo & "message.xml")
    Else
        If Mid(sOrderInfox, Len(sOrderInfox) - 1, 1) <> "\" Then
            sOrderInfox = sOrderInfox + "\"
        End If
        Documento.Save (sOrderInfox & "message.xml")
    End If
    Exit Sub
ErrorHandler:
End Sub
Public Sub KDS_EliminarProductoDeCombo(ByVal kdsRsCabecera As Recordset, ByVal itItem As String, ByVal xItem As String)
On Error GoTo ErrorKDSHandler
If sOrderInfox = "" Then 'SI SOLO TIENE UN MODELO DE KDS
    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<Transaction>
    Documento.loadXML ("<Transaction></Transaction>")
        '<Order>
        Dim nodoOrder As MSXML2.IXMLDOMNode
        Set nodoOrder = Documento.createElement("Order")
        Documento.documentElement.appendChild nodoOrder
            '<ID>
            Dim nodoID As MSXML2.IXMLDOMNode
            Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
            nodoID.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransType As MSXML2.IXMLDOMNode
                Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                nodoTransType.Text = "3"
                '<OrderStatus>
                Dim nodoOrderStatus As MSXML2.IXMLDOMNode
                Set nodoOrderStatus = nodoOrder.appendChild(Documento.createElement("OrderStatus"))
                nodoOrderStatus.Text = KDS_Obtener_OrderStatus() 'Cambiar
                '<OrderType>
                Dim nodoOrderType As MSXML2.IXMLDOMNode
                Set nodoOrderType = nodoOrder.appendChild(Documento.createElement("OrderType"))
                nodoOrderType.Text = KDS_Obtener_OrderType() 'Cambiar
                '<ServerName>
                Dim nodoServerName As MSXML2.IXMLDOMNode
                Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                nodoServerName.Text = kdsRsCabecera!tUsuario
                '<Destination>
                Dim nodoDestination As MSXML2.IXMLDOMNode
                Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                '<GuestTable>
                Dim nodoGuestTable As MSXML2.IXMLDOMNode
                Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                '<UserInfo>
                Dim nodoUserInfo As MSXML2.IXMLDOMNode
                Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
                
                Dim kdsRsDetalleCombo As Recordset
                Set kdsRsDetalleCombo = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "0", 0)
                
                '<Item>
                Dim nodoItem2 As MSXML2.IXMLDOMNode
                Set nodoItem2 = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID2 As MSXML2.IXMLDOMNode
                    Set nodoItemID2 = nodoItem2.appendChild(Documento.createElement("ID"))
                    nodoItemID2.Text = val(itItem) * 100 + val(xItem)
                    '<TransType>
                    Dim nodoItemTransType2 As MSXML2.IXMLDOMNode
                    Set nodoItemTransType2 = nodoItem2.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType2.Text = "2"
                    kdsRsDetalleCombo.MoveNext

    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & nodoID.Text & ".xml")
    
    
    
    
    Dim rsProductopedido As Recordset
    Set rsProductopedido = KDS_ObtenerProductoPedidoDeCombo(kdsRsCabecera!codigo, itItem, xItem)
    Dim nombreprod As String
    nombreprod = rsProductopedido!tResumido
    Dim NombreTemp As String
    NombreTemp = KDS_Obtener_InicialesDeNombre(rsProductopedido!ProtResumido)
    
    If (rsProductopedido!Area <> "") Then
        Call KDS_EnviaMensaje(rsProductopedido!Area, "(" & val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo))) & ")ELIMINADO:" & NombreTemp & "(" & nombreprod & ")")
    End If
    rsProductopedido.MoveNext
Else

    Dim kdsSaber As Recordset
    Set kdsSaber = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedidoxKDS '" & kdsRsCabecera!codigo & "', '" & itItem & "'", Cn)

    Do While Not kdsSaber.EOF

        Dim Documentox As MSXML2.DOMDocument60
        Set Documentox = New DOMDocument60
        '<Transaction>
        Documentox.loadXML ("<Transaction></Transaction>")
            '<Order>
            Dim nodoOrderx As MSXML2.IXMLDOMNode
            Set nodoOrderx = Documentox.createElement("Order")
            Documentox.documentElement.appendChild nodoOrderx
                '<ID>
                Dim nodoIDx As MSXML2.IXMLDOMNode
                Set nodoIDx = nodoOrderx.appendChild(Documentox.createElement("ID"))
                nodoIDx.Text = val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                    '<PosTerminal>
                    Dim nodoPosTerminalx As MSXML2.IXMLDOMNode
                    Set nodoPosTerminalx = nodoOrderx.appendChild(Documentox.createElement("PosTerminal"))
                    nodoPosTerminalx.Text = KDS_Obtener_PosTerminal()
                    '<TransType>
                    Dim nodoTransTypex As MSXML2.IXMLDOMNode
                    Set nodoTransTypex = nodoOrderx.appendChild(Documentox.createElement("TransType"))
                    nodoTransTypex.Text = "3"
                    '<OrderStatus>
                    Dim nodoOrderxStatusx As MSXML2.IXMLDOMNode
                    Set nodoOrderxStatusx = nodoOrderx.appendChild(Documentox.createElement("OrderStatus"))
                    nodoOrderxStatusx.Text = KDS_Obtener_OrderStatus() 'Cambiar
                    '<OrderType>
                    Dim nodoOrderxTypex As MSXML2.IXMLDOMNode
                    Set nodoOrderxTypex = nodoOrderx.appendChild(Documentox.createElement("OrderType"))
                    nodoOrderxTypex.Text = KDS_Obtener_OrderType() 'Cambiar
                    '<ServerName>
                    Dim nodoServerNamex As MSXML2.IXMLDOMNode
                    Set nodoServerNamex = nodoOrderx.appendChild(Documentox.createElement("ServerName"))
                    nodoServerNamex.Text = kdsRsCabecera!tUsuario
                    '<Destination>
                    Dim nodoDestinationx As MSXML2.IXMLDOMNode
                    Set nodoDestinationx = nodoOrderx.appendChild(Documentox.createElement("Destination"))
                    nodoDestinationx.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                    '<GuestTable>
                    Dim nodoGuestTablex As MSXML2.IXMLDOMNode
                    Set nodoGuestTablex = nodoOrderx.appendChild(Documentox.createElement("GuestTable"))
                    '<UserInfo>
                    Dim nodoUserInfox As MSXML2.IXMLDOMNode
                    Set nodoUserInfox = nodoOrderx.appendChild(Documentox.createElement("UserInfo"))
                    
                    Dim kdsRsDetalleCombox As Recordset
                    Set kdsRsDetalleCombox = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "0", 0)
                    
                    '<Item>
                    Dim nodoItem2x As MSXML2.IXMLDOMNode
                    Set nodoItem2x = nodoOrderx.appendChild(Documentox.createElement("Item"))
                        '<ID>
                        Dim nodoItemID2x As MSXML2.IXMLDOMNode
                        Set nodoItemID2x = nodoItem2x.appendChild(Documentox.createElement("ID"))
                        nodoItemID2x.Text = val(itItem) * 100 + val(xItem)
                        '<TransType>
                        Dim nodoItemTransType2x As MSXML2.IXMLDOMNode
                        Set nodoItemTransType2x = nodoItem2x.appendChild(Documentox.createElement("TransType"))
                        nodoItemTransType2x.Text = "2"
                        kdsRsDetalleCombox.MoveNext

        If kdsSaber!KDSX = 0 Then
            If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
                sOrderInfo = sOrderInfo + "\"
            End If
            Documentox.Save (sOrderInfo & nodoIDx.Text & ".xml")
        Else
            If Mid(sOrderInfox, Len(sOrderInfox) - 1, 1) <> "\" Then
                sOrderInfox = sOrderInfox + "\"
            End If
            Documentox.Save (sOrderInfox & nodoIDx.Text & ".xml")
        End If
        
        Dim rsProductopedidox As Recordset
        Set rsProductopedidox = KDS_ObtenerProductoPedidoDeCombo(kdsRsCabecera!codigo, itItem, xItem)
        Dim nombreprodx As String
        nombreprodx = rsProductopedidox!tResumido
        Dim NombreTempx As String
        NombreTempx = KDS_Obtener_InicialesDeNombre(rsProductopedidox!ProtResumido)
        
        If (rsProductopedidox!Area <> "") Then
            Call KDS_EnviaMensajex(rsProductopedidox!Area, "(" & val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo))) & ")ELIMINADO:" & NombreTempx & "(" & nombreprodx & ")", kdsSaber!KDSX)
        End If
        rsProductopedidox.MoveNext

      kdsSaber.MoveNext
    Loop
        

End If
    
    
    
    
    Exit Sub
ErrorKDSHandler:
End Sub

Private Sub KDS_EnviaMensaje(ByVal stationID As String, ByVal info As String)
On Error GoTo ErrorHandler
    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<StationInfo>
    Documento.loadXML ("<StationInfo></StationInfo>")
        '<StationID>
        Dim nodoStationID As MSXML2.IXMLDOMNode
        Set nodoStationID = Documento.createElement("StationID")
        nodoStationID.Text = val(stationID)
        Documento.documentElement.appendChild nodoStationID
        '<User>
        Dim nodoUser As MSXML2.IXMLDOMNode
        Set nodoUser = Documento.createElement("User")
        nodoUser.Text = "0"
        Documento.documentElement.appendChild nodoUser
        '<Info>
        Dim nodoInfo As MSXML2.IXMLDOMNode
        Set nodoInfo = Documento.createElement("Info")
        nodoInfo.Text = info
        Documento.documentElement.appendChild nodoInfo
    
    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & "message.xml")
    
    Exit Sub
ErrorHandler:
End Sub

Private Function KDS_Obtener_PosTerminal()
    KDS_Obtener_PosTerminal = val(sCaja)
End Function

Private Function KDS_Obtener_TransType()
    '<!-- 1 Añadir nuevo orden, para anexar a la última posición-->
    '<!-- 2 Eliminar esta orden. Si utiliza este valor, el KDS sólo hay etiquetas para identificación.
            'Otras etiquetas pueden ser cualquier valor, o no las transferencias -->
    '<!-- 3 Modificar este orden. Sólo tranfer todas las etiquetas para cambiar.
            'Si el valor de la etiqueta está en blanco, KDS tratar con él como sin cambios -->
    '<!-- 4 Reservado para uso futuro -->
    '<!-- 5 Pregunta que este estado de la orden -->
    KDS_Obtener_TransType = "1"
End Function

Private Function KDS_Obtener_OrderStatus()
    '<!-- 0 no pagado -->
    '<!-- 1 pagados -->
    '<!-- 2 En proceso -->
    KDS_Obtener_OrderStatus = "1"
End Function

Private Function KDS_Obtener_OrderType()
    '<!-- ""- Orden normal -->
    '<!-- RUSH- Orden Rush -->
    '<!-- Fire- fire order-->
    KDS_Obtener_OrderType = ""
End Function

Private Function KDS_Obtener_Destination(ByVal tCodigo As String)
    '<!-- El destino de esta orden -->
On Error GoTo ErrorKDSHandler
    Dim destino As String
    destino = Lib.OpenRecordset("USP_KDS_ObtenerTipoPedido '" & tCodigo & "'", Cn)!tDetallado
    KDS_Obtener_Destination = destino
    Exit Function
ErrorKDSHandler:
    Obtener_Destination = ""
End Function

Private Function KDS_Obtener_Category(ByVal tCodigoGrupo As String, ByVal tCodigoSubGrupo As String)
On Error GoTo ErrorKDSHandler
    'Obtener_Category = "A Full Dente Personal"
    Dim categoria As String
    categoria = Lib.OpenRecordset("USP_KDS_ObtenerCategoria '" & tCodigoGrupo & "','" & tCodigoSubGrupo & "'", Cn)!categoria
    KDS_Obtener_Category = categoria
    Exit Function
ErrorKDSHandler:
    KDS_Obtener_Category = ""
End Function

Private Function KDS_ObtenerDetallePedido(ByVal tCodigoPedido As String, ByVal accion As Integer) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsDetalle As Recordset
    Set kdsRsDetalle = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedido '" & tCodigoPedido & "', '" & accion & "'", Cn)
    Set KDS_ObtenerDetallePedido = kdsRsDetalle
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerDetallePedido = New Recordset
End Function

Private Function KDS_ObtenerDetallePedidox(ByVal tCodigoPedido As String, ByVal accion As Integer, ByVal xkds As Integer) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsDetalle As Recordset
    Set kdsRsDetalle = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedidox '" & tCodigoPedido & "', '" & accion & "', '" & xkds & "'", Cn)
    Set KDS_ObtenerDetallePedidox = kdsRsDetalle
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerDetallePedidox = New Recordset
End Function
Private Function KDS_ObtenerDetalleCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal lImprime As String, ByVal accion As Integer) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsDetalleCombo As Recordset
    Set kdsRsDetalleCombo = Lib.OpenRecordset("USP_KDS_ObtenerDetalleCombo '" & tCodigoPedido & "','" & tItem & "','" & lImprime & "','" & accion & "'", Cn)
    Set KDS_ObtenerDetalleCombo = kdsRsDetalleCombo
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerDetalleCombo = New Recordset
End Function

Private Function KDS_ObtenerPropiedadesProducto(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tItemCombo As String, ByVal tProducto As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsPropiedadProducto As Recordset
    Set kdsRsPropiedadProducto = Lib.OpenRecordset("USP_KDS_ObtenerPropiedadesProducto '" & tCodigoPedido & "','" & tItem & "','" & tItemCombo & "','" & tProducto & "'", Cn)
    Set KDS_ObtenerPropiedadesProducto = kdsRsPropiedadProducto
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerPropiedadesProducto = New Recordset
End Function

Private Function KDS_ObtenerInfoPropiedadProducto(ByVal tEnlace As String, ByVal tCodigoProducto As String, ByVal tCodigoPropiedad As String) As String
On Error GoTo ErrorKDSHandler
    Dim info As Recordset
    Set info = Lib.OpenRecordset("USP_RD_ObtenerInfoPropiedadProducto '" & tEnlace & "','" & tCodigoPropiedad & "','" & tCodigoProducto & "'", Cn)
    info.MoveFirst
    KDS_ObtenerInfoPropiedadProducto = info!Operador + " " + info!Propiedad
    Exit Function
ErrorKDSHandler:
    KDS_ObtenerInfoPropiedadProducto = ""
End Function

Public Function KDS_ValidarProductoArea(ByVal tCodigoProducto As String, ByVal tArea As String) As Boolean
On Error GoTo ErrorKDSHandler
    Dim RsArea As Recordset
    Set RsArea = Lib.OpenRecordset("USP_KDS_ObtenerArea '" & tArea & "'", Cn)
    RsArea.MoveFirst
    
    If (RsArea!kds = 1) Then
        Dim rsproductoArea As Recordset
        Set rsproductoArea = Lib.OpenRecordset("USP_KDS_ObtenerProductoArea '" & tCodigoProducto & "', '" & tArea & "'", Cn)
        If (rsproductoArea!Cantidad > 0) Then
            KDS_ValidarProductoArea = True 'no se puede insertar
        Else
            KDS_ValidarProductoArea = False 'permiso para insertar
        End If
    Else
        KDS_ValidarProductoArea = False 'permiso para inseratr
    End If
    
    Exit Function
ErrorKDSHandler:
    KDS_ValidarProductoArea = True
End Function

Public Function KDS_ObtenerAreaImpresionKDS(ByVal tCodigoProducto As String, ByVal tItemCombo As String, ByVal tCodigoPedido As String, ByVal tItem As String)
On Error GoTo ErrorKDSHandler
    Dim RsArea As Recordset
    Set RsArea = Lib.OpenRecordset("USP_KDS_ObtenerAreaImpresionKDS '" & tCodigoProducto & "','" & tItemCombo & "','" & tCodigoPedido & "','" & tItem & "'", Cn)
    KDS_ObtenerAreaImpresionKDS = RsArea!tArea
    Exit Function
ErrorKDSHandler:
    KDS_ObtenerAreaImpresionKDS = ""
End Function

Public Function KDS_ObtenerProductoPedido(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim rsProductopedido As Recordset
    Set rsProductopedido = Lib.OpenRecordset("USP_KDS_ObtenerProductoPedido '" & tCodigoPedido & "','" & tItem & "'", Cn)
    Set KDS_ObtenerProductoPedido = rsProductopedido
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerProductoPedido = New Recordset
End Function

Public Function KDS_ObtenerProductoPedidoDeCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal xItem As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim rsProductopedido As Recordset
    Set rsProductopedido = Lib.OpenRecordset("USP_KDS_ObtenerProductoPedidoDeCombo '" & tCodigoPedido & "','" & tItem & "','" & xItem & "'", Cn)
    Set KDS_ObtenerProductoPedidoDeCombo = rsProductopedido
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerProductoPedidoDeCombo = New Recordset
End Function

Public Function KDS_ObtenerProductoPedidoImpresos(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal accion As Integer) As Recordset
On Error GoTo ErrorKDSHandler
    Dim rsProductopedido As Recordset
    Set rsProductopedido = Lib.OpenRecordset("USP_KDS_ObtenerProductoPedidoImpresos '" & tCodigoPedido & "','" & tItem & "', '" & accion & "'", Cn)
    Set KDS_ObtenerProductoPedidoImpresos = rsProductopedido
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerProductoPedidoImpresos = New Recordset
End Function

Private Function KDS_Obtener_Mesa(ByVal tMesa As String) As String
On Error GoTo ErrorKDSHandler
    Dim RsMesa As Recordset
    Set RsMesa = Lib.OpenRecordset("USP_KDS_ObtenerNombreMesaxCodigo '" & tMesa & "'", Cn)
    KDS_Obtener_Mesa = RsMesa!tResumido
    Exit Function
ErrorKDSHandler:
    KDS_Obtener_Mesa = ""
End Function


Public Function KDS_Obtener_InicialesDeNombre(ByVal nombre As String) As String
On Error GoTo ErrorKDSHandler
    Dim NombreTemp As String
    Dim PosNombreTemo As String
    PosNombreTemo = 0
    NombreTemp = Mid(nombre, 1, 1)
    Do
        PosNombreTemo = InStr(PosNombreTemo + 1, nombre, " ", vbTextCompare)
        If (PosNombreTemo) = 0 Then
            Exit Do
        End If
        NombreTemp = NombreTemp & Mid(nombre, PosNombreTemo + 1, 1)
    Loop
    KDS_Obtener_InicialesDeNombre = NombreTemp
    Exit Function
ErrorKDSHandler:
    KDS_Obtener_InicialesDeNombre = ""
End Function

Public Function KDS_ProcesarBumpNotification(ByVal NombArchivo As String)
    Dim strAnio As String
    Dim strMes As String
    Dim strDia As String
    Dim strHora As String
    Dim strMinuto As String
    Dim strSegundo As String
    Dim strPedido As String
    Dim strItemId As String
    
    strAnio = Mid(NombArchivo, 0, 4)
    strMes = Mid(NombArchivo, 5, 2)
    strDia = Mid(NombArchivo, 8, 2)
    strHora = Mid(NombArchivo, 11, 2)
    strMinuto = Mid(NombArchivo, 14, 2)
    strSegundo = Mid(NombArchivo, 0, 4)
    strPedido = Mid(NombArchivo, 0, 4)
    strItemId = Mid(NombArchivo, 0, 4)
End Function
'---------------------------------------------------
    'Agregar lña referencia a Microsoft Scripting Runtime
'---------------------------------------------------
Public Sub KDS_ListarBumpNotification(ByVal sBump As String)
On Error GoTo Err_Sub
    Dim El_Archivo As File
    Dim El_Directorio As Folder
    Set fso = New FileSystemObject
    Set El_Directorio = fso.GetFolder(sBump)
    For Each El_Archivo In El_Directorio.Files
        Dim fFecha As Date
        'El_Archivo.DateCreated
        Cargar_XML sBump & "\" & El_Archivo.name, El_Archivo.DateCreated
        El_Archivo.Move (sBump & "\Historial\" & El_Archivo.name)
    Next El_Archivo
Exit Sub
Err_Sub:
MsgBox err.Description
End Sub

Private Sub Cargar_XML(Path_XML As String, FechaCreacion As Date)
On Error GoTo Err_Sub
    Dim objPeopleRoot As IXMLDOMElement
    Dim objPersonElement As IXMLDOMElement
    Dim ObjElement As IXMLDOMNode
    Dim X As IXMLDOMNodeList
    If Len(Dir(Path_XML)) = 0 Then
       MsgBox "El archivo " & Path_XML & _
               " No está en el directorio ." & vbNewLine & _
               " Compruebe la ruta", vbCritical
       Exit Sub
    End If
    'Seteamos la variable
    Set m_objDOMPeople = New DOMDocument60
    m_objDOMPeople.resolveExternals = True
    'Para que valide el documento xml
    m_objDOMPeople.validateOnParse = True
    'Carga el documento
    m_objDOMPeople.async = False
    Call m_objDOMPeople.Load(Path_XML)
    'Comprobamos si se carga
    If m_objDOMPeople.parseError.reason <> "" Then
        ' si hay un error muestra el fallo
        MsgBox m_objDOMPeople.parseError.reason
        Exit Sub
    End If
    Set objPeopleRoot = m_objDOMPeople.documentElement
    Dim Index As Integer
    Dim Lista As IXMLDOMNodeList
    
    Dim ytCodigoPedido As String
    ytCodigoPedido = objPeopleRoot.childNodes.nextNode.childNodes.Item(0).Text
    Dim ytItem As String
    ytItem = objPeopleRoot.childNodes.nextNode.childNodes.Item(8).childNodes.Item(0).Text
    Dim ytCodigoProducto As String
        
    Call USP_KDS_GrabarTiempoSalidaDPedido(ytCodigoPedido, ytItem, ytCodigoProducto, FechaCreacion)
Exit Sub
Err_Sub:
MsgBox err.Description
End Sub

Private Function USP_KDS_GrabarTiempoSalidaDPedido(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tCodigoProducto As String, ByVal fSalida As Date)
On Error GoTo ErrorKDSHandler
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("USP_KDS_GrabarTiempoSalidaDPedido", Cn) Then
        Set oComando = Nothing
        Exit Function
    End If
    oComando.CreateParameter "@tCodigoPedido", adVarChar, adParamInput, 10, tCodigoPedido
    oComando.CreateParameter "@tItem", adVarChar, adParamInput, 3, tItem
    oComando.CreateParameter "@fSalida", adDate, adParamInput, 10, fSalida
    
    If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Function
   End If

   oComando.ExecSP
    Exit Function
ErrorKDSHandler:
    USP_KDS_GrabarTiempoSalidaDPedido = ""
End Function


