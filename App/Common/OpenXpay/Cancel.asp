<%
    '/*
    ' * [������� ��û ������]
    ' *
    ' * LG���÷������� ���� �������� �ŷ���ȣ(LGD_TID)�� ������ ��� ��û�� �մϴ�.(�Ķ���� ���޽� POST�� ����ϼ���)
    ' * (���ν� LG���÷������� ���� �������� PAYKEY�� ȥ������ ������.)
    ' */

    CST_PLATFORM         = trim(request("CST_PLATFORM"))        ' LG���÷��� �������� ����(test:�׽�Ʈ, service:����)
    CST_MID              = trim(request("CST_MID"))             ' LG���÷������� ���� �߱޹����� �������̵� �Է��ϼ���.
                                                                ' �׽�Ʈ ���̵�� 't'�� �����ϰ� �Է��ϼ���.
    if CST_PLATFORM = "test" then                               ' �������̵�(�ڵ�����)
        LGD_MID = "t" & CST_MID
    else
        LGD_MID = CST_MID
    end if
    LGD_TID                 = trim(request("LGD_TID"))                  ' LG���÷������� ���� �������� �ŷ���ȣ(LGD_TID)
    LGD_CANCELREASON        = trim(request("LGD_CANCELREASON"))         ' ��һ���
    LGD_CANCELREQUESTER     = trim(request("LGD_CANCELREQUESTER"))      ' ��ҿ�û��
    LGD_CANCELREQUESTERIP   = trim(request("LGD_CANCELREQUESTERIP"))    ' ��ҿ�ûIP
    

    configPath = "D:/inetPub/lgdacom"  'LG���÷������� ������ ȯ������("/conf/lgdacom.conf, /conf/mall.conf") ��ġ ����.  


    Set xpay = CreateObject("XPayClientCOM.XPayClient")
    xpay.Init configPath, CST_PLATFORM
    xpay.Init_TX(LGD_MID)

    xpay.Set "LGD_TXNAME", "Cancel"
    xpay.Set "LGD_TID", LGD_TID
    xpay.Set "LGD_CANCELREASON", LGD_CANCELREASON
    xpay.Set "LGD_CANCELREQUESTER", LGD_CANCELREQUESTER
    xpay.Set "LGD_CANCELREQUESTERIP", LGD_CANCELREQUESTERIP
 

    '/*
    ' * 1. ������� ��û ���ó��
    ' *
    ' * ��Ұ�� ���� �Ķ���ʹ� �����޴����� �����Ͻñ� �ٶ��ϴ�.
	' *
	' * [[[�߿�]]] ���翡�� ������� ó���ؾ��� �����ڵ�
	' * 1. �ſ�ī�� : 0000, AV11  
	' * 2. ������ü : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (ȯ�������� ����-> ȯ�Ұ���ڵ�.xls ����)
	' * 3. ������ ���������� ��� 0000(����) �� ��Ҽ��� ó��
	' *
    ' */

    if xpay.TX() then
        '1)������Ұ�� ȭ��ó��(����,���� ��� ó���� �Ͻñ� �ٶ��ϴ�.)
        Response.Write("������� ��û�� �Ϸ�Ǿ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    else
        '2)API ��û ���� ȭ��ó��
        Response.Write("������� ��û�� �����Ͽ����ϴ�. <br>")
        Response.Write("TX Response_code = " & xpay.resCode & "<br>")
        Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
    end if
%>
