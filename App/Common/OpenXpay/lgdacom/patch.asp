<%
    configPath = "D:/inetPub/lgdacom"  'LG���÷������� ������ ȯ������("/conf/lgdacom.conf, /conf/mall.conf") ��ġ ����.  

    '/*
    ' * [LG�ڷ��� ȯ������ UPDATE]
    ' *
    ' * �� �������� LG�ڷ��� ȯ�������� UPDATE �մϴ�.(�������� ������.)
    ' */ 
    CST_PLATFORM 	= trim(request("CST_PLATFORM"))         
    CST_MID 		= trim(request("CST_MID"))
    
    if CST_PLATFORM = "" then
    	Response.Write("[TX_PATCH error] CST_PLATFORM �Ķ���� ����<br>")
        Response.End
    end if

    if CST_MID = "" then
    	Response.Write("[TX_PATCH error] CST_MID �Ķ���� ����<br>")        
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
    
	Response.Write("patch result = " & xpay.Patch("lgdacom.conf") & "<P>")
%>
