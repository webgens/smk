<%
	configPath = "D:/inetPub/lgdacom"  'LG���÷������� ������ ȯ������("/conf/lgdacom.conf, /conf/mall.conf") ��ġ ����.  
 
    '/*
    ' * [LG�ڷ��� ���� Ȯ�� ������]
    ' *
    ' * ���������� LG�ڷ��ް��� ������ �׽�Ʈ �ϴ� ������ �Դϴ�.(�������� ������.)
    ' */
    CST_PLATFORM 	= trim(request("CST_PLATFORM"))         
    CST_MID 		= trim(request("CST_MID"))  't �� �߰����� ���� ���Կ�û�� ���̵� �Է¹ٶ��ϴ�.
            
    if CST_PLATFORM = "" then
    	Response.Write("[TX_PING error] CST_PLATFORM �Ķ���� ����<br>")
        Response.End
    end if

    if CST_MID = "" then
    	Response.Write("[TX_PING error] CST_MID �Ķ���� ����<br>")        
        Response.End
    end if
    
    if CST_PLATFORM = "test" then
        LGD_MID = "t" & CST_MID                                   
    else
        LGD_MID = CST_MID                                         
    end if     
    
	Dim xpay
	Dim i, j
	Dim itemName
	
	Set xpay = server.CreateObject("XPayClientCOM.XPayClient")	
    xpay.Init configPath, CST_PLATFORM   
    xpay.Init_TX(LGD_MID)
    xpay.Set "LGD_TXNAME", "Ping"
    xpay.Set "LGD_RESULTCNT", "3"
    xpay.TX()	
	
	Response.Write("ResCode = " & xpay.resCode & "<P>")
	Response.Write("ResMsg = " & xpay.resMsg & "<P>")

	Response.Write("Response <P>")

    Dim itemCount
    Dim resCount
    itemCount = xpay.resNameCount
    resCount = xpay.resCount
	
    For i = 0 To itemCount - 1
        itemName = xpay.ResponseName(i)
		Response.Write(itemName & "<P>")
        For j = 0 To resCount - 1
			Response.Write(xpay.Response(itemName, j) & "<P>")
        Next
    Next
%>

