<%
    '/*
    ' * [���ݿ����� �߱� ��û ������]
    ' *
    ' * �Ķ���� ���޽� POST�� ����ϼ���
    ' */
    CST_PLATFORM               = trim(request("CST_PLATFORM"))       	'LG���÷��� ���� ���� ����(test:�׽�Ʈ, service:����)
    CST_MID                    = trim(request("CST_MID"))            	'�������̵�(LG���÷������� ���� �߱޹����� �������̵� �Է��ϼ���)
                                                                     	'�׽�Ʈ ���̵�� 't'�� �ݵ�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                                    	'�������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_METHOD                 = trim(request("LGD_METHOD"))         	 '�޼ҵ�('AUTH':����, 'CANCEL' ���)
    LGD_OID                    = trim(request("LGD_OID"))            	 '�ֹ���ȣ(�������� ����ũ�� �ֹ���ȣ�� �Է��ϼ���)
	LGD_PAYTYPE                = trim(request("LGD_PAYTYPE"))         	 '�������� �ڵ� (SC0030:������ü, SC0040:�������, SC0100:�������Ա� �ܵ�)
    LGD_AMOUNT                 = trim(request("LGD_AMOUNT"))         	 '�ݾ�("," �� ������ �ݾ��� �Է��ϼ���)
    LGD_CASHCARDNUM            = trim(request("LGD_CASHCARDNUM"))    	 '�߱޹�ȣ(�ֹε�Ϲ�ȣ,���ݿ�����ī���ȣ,�޴�����ȣ ���)
    LGD_CUSTOM_MERTMAME        = trim(request("LGD_CUSTOM_MERTNAME"))    '������
    LGD_CUSTOM_BUSINESSNUM     = trim(request("LGD_CUSTOM_BUSINESSNUM")) '����ڵ�Ϲ�ȣ
    LGD_CUSTOM_MERTPHONE       = trim(request("LGD_CUSTOM_MERTPHONE")) 	 '���� ��ȭ��ȣ
    LGD_CASHRECEIPTUSE     	   = trim(request("LGD_CASHRECEIPTUSE")) 	 '���ݿ������߱޿뵵('1':�ҵ����, '2':��������)
    LGD_PRODUCTINFO     	   = trim(request("LGD_PRODUCTINFO")) 	     '��ǰ��
    LGD_TID     	   		   = trim(request("LGD_TID")) 	 		     'LG���÷��� �ŷ���ȣ
    
    configPath = "D:/inetPub/lgdacom"  'LG���÷������� ������ ȯ������("/conf/lgdacom.conf, /conf/mall.conf") ��ġ ����.  
    
	Dim xpay
	Dim i, j
	Dim itemName
	
	Set xpay = server.CreateObject("XPayClientCOM.XPayClient")	
    xpay.Init configPath, CST_PLATFORM    
    xpay.Init_TX(LGD_MID)
    xpay.Set "LGD_TXNAME", "CashReceipt"
	xpay.Set "LGD_METHOD", LGD_METHOD
	xpay.Set "LGD_PAYTYPE", LGD_PAYTYPE
	
	if LGD_METHOD = "AUTH" then              '���ݿ����� �߱� ��û 
		xpay.Set "LGD_OID", LGD_OID 
		xpay.Set "LGD_AMOUNT", LGD_AMOUNT
		xpay.Set "LGD_CASHCARDNUM", LGD_CASHCARDNUM
		xpay.Set "LGD_CUSTOM_MERTMAME", LGD_CUSTOM_MERTMAME
		xpay.Set "LGD_CUSTOM_BUSINESSNUM", LGD_CUSTOM_BUSINESSNUM
		xpay.Set "LGD_CUSTOM_MERTPHONE", LGD_CUSTOM_MERTPHONE
		xpay.Set "LGD_CASHRECEIPTUSE", LGD_CASHRECEIPTUSE
		
		if LGD_PAYTYPE = "SC0030" then		'������� ������ü�� ���ݿ����� �߱޿�û�� �ʼ� 
			xpay.Set "LGD_TID", LGD_TID
		elseIf	LGD_PAYTYPE = "SC0040" then	'������� ������°� ���ݿ����� �߱޿�û�� �ʼ�  
			xpay.Set "LGD_TID", LGD_TID
			xpay.Set "LGD_SEQNO", "001"
		else								'�������Ա� �ܵ��� �߱޿�û  
			xpay.Set "LGD_PRODUCTINFO", LGD_PRODUCTINFO
		end if
	
	else									'���ݿ����� ��� ��û
		xpay.Set "LGD_TID", LGD_TID
		
		if	LGD_PAYTYPE = "SC0040" then		'������°� ���ݿ����� �߱���ҽ� �ʼ�
			xpay.Set "LGD_SEQNO", "001"
		end if	
	end if
 
    xpay.TX()	
	Response.Write("���ݿ����� ó���� �Ϸ�Ǿ����ϴ�. <br>")	
	Response.Write("TX Response_code = " & xpay.resCode & "<br>")
	Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")	
    
    Response.Write("����ڵ� : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
    Response.Write("����޼��� : " & xpay.Response("LGD_RESPMSG", 0) & "<p>")

    Response.Write("[��� �Ķ����]<br>")

   '�Ʒ��� ��� �Ķ���͸� ��� ��� �ݴϴ�.   
	Dim itemCount
    Dim resCount
    itemCount = xpay.resNameCount
    resCount = xpay.resCount
	
    For i = 0 To itemCount - 1
        itemName = xpay.ResponseName(i)
		Response.Write(itemName & "&nbsp:&nbsp")
        For j = 0 To resCount - 1
			Response.Write(xpay.Response(itemName, j) & "<br>")
        Next
    Next	
%>
