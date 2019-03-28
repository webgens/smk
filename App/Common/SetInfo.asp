<%
'**********************************************************************************'
'현재 날짜 START
'----------------------------------------------------------------------------------'
DIM R_YEAR, R_MONTH, R_DAY
DIM R_HOUR, R_MIN, R_SEC
DIM U_TIME, U_DATE
R_YEAR				 = Year(Date)
R_MONTH				 = Month(Date)
R_DAY				 = Day(Date)
R_HOUR				 = Hour(Time)
R_MIN				 = Minute(Time)
R_SEC				 = Second(Time)

IF LEN(R_MONTH)		 = 1 THEN R_MONTH	 = "0"&R_MONTH
IF LEN(R_DAY)		 = 1 THEN R_DAY		 = "0"&R_DAY
IF LEN(R_HOUR)		 = 1 THEN R_HOUR	 = "0"&R_HOUR
IF LEN(R_MIN)		 = 1 THEN R_MIN		 = "0"&R_MIN
IF LEN(R_SEC)		 = 1 THEN R_SEC		 = "0"&R_SEC

U_TIME = R_HOUR & R_MIN & R_SEC
U_DATE = R_YEAR & R_MONTH & R_DAY
'----------------------------------------------------------------------------------'
'현재 날짜 END
'----------------------------------------------------------------------------------'




'**********************************************************************************'
'회원 정보 쿠키 값 START
'----------------------------------------------------------------------------------'
DIM U_CARTID		'# CARTID
DIM U_IP			'# 접속 아이피
DIM U_CIP			'# 접속 아이피(쿠키값)
DIM U_ID			'# 회원 아이디
DIM U_NUM			'# 회원 번호
DIM U_NAME			'# 회원 이름
DIM U_MFLAG			'# 회원여부(Y/N)
DIM U_EFLAG			'# 임직원여부(Y/N)
DIM U_ETYPE			'# 임직원회사구분(P:일반회원/S:슈마커/J:JD스포츠)
DIM U_GROUP			'# 회원 그룹
DIM U_SNSKIND		'# SNS 간편로그인 종료
	
DIM U_GuestInfo		'# 3일간 저장되는 쿠키

DIM N_NAME			'# 비회원 이름
DIM N_HP			'# 비회원 휴대폰번호
DIM N_EMAIL			'# 비회원 이메일

DIM U_ISAPP			'# 앱여부
DIM U_DEVICEID		'# 핸드폰 ID
DIM U_DEVICE		'# 핸드폰 OS
DIM U_PUSHKEY		'# 푸쉬키
DIM U_MODELNAME		'# 핸드폰 기종
DIM U_APPVERSION	'# 앱버전
	

IF TRIM(Decrypt(Request.Cookies("USESSIONID"))) = "" THEN
		Response.Cookies("UCARTID")		 = Encrypt(Session.SessionID)
END IF
IF TRIM(Decrypt(Request.Cookies("USESSIONID"))) = "" THEN
		Response.Cookies("USESSIONID")	 = Encrypt(Session.SessionID)
END IF
IF TRIM(Decrypt(Request.Cookies("GuestInfo"))) = "" THEN
		Response.Cookies("GuestInfo") = Encrypt(U_DATE&U_TIME&Session.SessionID)
		Response.Cookies("GuestInfo").Expires = Now() + 3
End IF


U_CARTID			 = TRIM(Decrypt(Request.Cookies("UCARTID")))

U_GuestInfo			 = TRIM(Decrypt(Request.Cookies("GuestInfo")))


U_IP				 = Request.ServerVariables("REMOTE_ADDR")

U_CIP				 = TRIM(Decrypt(Request.Cookies("UIP")))
U_ID				 = TRIM(Decrypt(Request.Cookies("UID")))
U_NUM				 = TRIM(Decrypt(Request.Cookies("UNUM")))
U_NAME				 = TRIM(Decrypt(Request.Cookies("UNAME")))
U_MFLAG				 = TRIM(Decrypt(Request.Cookies("UMFLAG")))
U_EFLAG				 = TRIM(Decrypt(Request.Cookies("UEFLAG")))
U_ETYPE				 = TRIM(Decrypt(Request.Cookies("UETYPE")))
U_GROUP				 = TRIM(Decrypt(Request.Cookies("UGROUP")))
U_SNSKIND			 = TRIM(Decrypt(Request.Cookies("USNSKIND")))
IF U_GROUP = "" THEN U_GROUP = "0"

'//비회원주문일때
N_NAME				 = Request.Cookies("N_NAME")
N_HP				 = Request.Cookies("N_HP")
N_EMAIL				 = Request.Cookies("N_EMAIL")
IF IsNull(N_NAME)	OR IsEmpty(N_NAME)	THEN N_NAME		= ""
IF IsNull(N_HP)		OR IsEmpty(N_HP)	THEN N_HP		= ""
IF IsNull(N_EMAIL)	OR IsEmpty(N_EMAIL)	THEN N_EMAIL	= ""


U_ISAPP				 = TRIM(Decrypt(Request.Cookies("U_ISAPP")))
U_DEVICEID			 = TRIM(Decrypt(Request.Cookies("U_DEVICEID")))
U_DEVICE			 = TRIM(Decrypt(Request.Cookies("U_DEVICE")))
U_PUSHKEY			 = TRIM(Decrypt(Request.Cookies("U_PUSHKEY")))
U_MODELNAME			 = TRIM(Decrypt(Request.Cookies("U_MODELNAME")))
U_APPVERSION		 = TRIM(Decrypt(Request.Cookies("U_APPVERSION")))
'----------------------------------------------------------------------------------'
'관리자 정보 쿠키 값 END
'----------------------------------------------------------------------------------'




'**********************************************************************************'
'URL START
'----------------------------------------------------------------------------------'
DIM HOME_DOMAIN
DIM HOME_DOMAIN_HTTS
DIM HOME_DOMAIN_NOHTTPS
DIM HOME_URL
DIM HOME_URL1
DIM FRONT_URL
DIM IMAGE_URL
DIM LOGIN_URL

HOME_DOMAIN					= "https://app.shoemarker.co.kr"
HOME_DOMAIN_HTTS			= "https://app.shoemarker.co.kr"
HOME_DOMAIN_NOHTTPS			= "http://app.shoemarker.co.kr"
HOME_URL					= "https://app.shoemarker.co.kr"
HOME_URL1					= HOME_URL & "/"
FRONT_URL					= "https://www.shoemarker.co.kr"
IMAGE_URL					= "https://app.shoemarker.co.kr"

LOGIN_URL					= HOME_URL & "/ASP/Member/Login.asp"
'----------------------------------------------------------------------------------'
'URL END
'----------------------------------------------------------------------------------'




	
'**********************************************************************************'
'FILE UPLOAD 설정 START
'----------------------------------------------------------------------------------'
CONST DENY_EXT		 = "html,htm,asp,aspx,exe,com,bat,dll"
DIM ALLOW_EXT_IMG		'# UPLOAD 허용 이미지 확장자
DIM ALLOW_EXT_DOCU	'# UPLOAD 허용 문서 확장자
ALLOW_EXT_IMG		 = "jpg/gif/png/jpeg"
ALLOW_EXT_DOCU		 = "jpg/gif/png/jpeg/doc/xls/ppt/pdf/hwp/zip/txt"
'----------------------------------------------------------------------------------'
'FILE UPLOAD 설정 END
'----------------------------------------------------------------------------------'





'**********************************************************************************'
'파일 경로 START
'----------------------------------------------------------------------------------'
DIM D_HOME							'# 몰 관리자 DIR.
DIM D_UPLOAD						'# 업로드 기본 경로.
DIM D_BRAND							'# 브랜드 이미지
DIM D_REVIEW						'# 상품후기 이미지
DIM D_MTMQNA						'# 1:1상담 이미지
DIM D_NOTICE						'# 공지사항 파일
DIM D_PARTNERSHIP					'# 입점/제휴문의, 단체구매 파일
DIM D_ORDERAS						'# A/S
DIM D_EMAIL							'# 이메일 폼 파일

D_HOME								 = ""
D_UPLOAD							 = "/Upload/"	
D_BRAND								 = D_UPLOAD & "Brand/"
D_REVIEW							 = D_UPLOAD & "Community/ProductReview/"
D_MTMQNA							 = D_UPLOAD & "Community/MtmQna/"
D_NOTICE							 = D_UPLOAD & "Community/Notice/"
D_PARTNERSHIP						 = D_UPLOAD & "Customer/PartnerShip/"
D_ORDERAS							 = D_UPLOAD & "OrderAS/"
D_EMAIL								 = "/Common/Mail/Html/"
'----------------------------------------------------------------------------------'
'파일 경로 END
'----------------------------------------------------------------------------------'





'**********************************************************************************'
'API 정보 START
'----------------------------------------------------------------------------------'
DIM NICE_H_ID						'# NICE 휴대폰인증 아이디
DIM NICE_H_PWD						'# NICE 휴대폰인증 비밀번호

NICE_H_ID							 = "G7715"
NICE_H_PWD							 = "QIWPJS0ONF5O"

DIM IPIN_H_ID						'# IPIN 인증 아이디
DIM IPIN_H_PWD						'# IPIN 인증 비밀번호

IPIN_H_ID							 = "K790"
IPIN_H_PWD							 = "17069643"


DIM NAVER_MAP_CLIENTID				'# 네이버 MAP CLIENTID
DIM NAVER_MAP_CLIENTSECRET			'# 네이버 MAP CLIENT SECRET
NAVER_MAP_CLIENTID					 = "FCZrIdYaPdkKhA13dHFM"
NAVER_MAP_CLIENTSECRET				 = "kKZV4e2igk"


DIM NAVER_LOGIN_CLIENTID			'# 네이버 로그인 CLIENTID
DIM NAVER_LOGIN_CLIENTSECRET		'# 네이버 로그인 CLIENT SECRET
NAVER_LOGIN_CLIENTID				 = "btIix99vJZubLIoFVrYo"
NAVER_LOGIN_CLIENTSECRET			 = "S20UuXwlLV"


DIM GOOGLE_LOGIN_CLIENTID			'# 구글 로그인 CLIENTID
DIM GOOGLE_LOGIN_CLIENTSECRET		'# 구글 로그인 CLIENT SECRET
GOOGLE_LOGIN_CLIENTID				 = "477472571397-npfiub75halm3toglhhoohptg6k39l2k.apps.googleusercontent.com"
GOOGLE_LOGIN_CLIENTSECRET			 = "aeooHtiOXcTZSXYUSchmMcXn"



DIM KAKAO_LOGIN_CLIENTID			'# 카카오 SNS KEY
KAKAO_LOGIN_CLIENTID				 = "5947a50c87d320700fbba9d90ff16a56"

DIM KAKAO_JAVASCRIPT_KEY			'# 카카오 JAVASCRIPT KEY
KAKAO_JAVASCRIPT_KEY				 = "abd31b6fe9130f34beb0779b5e7fcb00"


DIM FACEBOOK_LOGIN_CLIENTID			'# 페이스북 API ID
DIM FACEBOOK_LOGIN_APPSECRET		'# 페이스북 APP SECRET
FACEBOOK_LOGIN_CLIENTID				 = "384834005263410"
FACEBOOK_LOGIN_APPSECRET			 = "c9b7192d7393d82ee89a512131143724"


DIM KAKAOTALK_SENDER_KEY			 '# 카카오톡 SENDER KEY
KAKAOTALK_SENDER_KEY				 = "faa9951dd4d2c0b59acd9cd451c60576490d8cb1"
'----------------------------------------------------------------------------------'
'API 정보 END
'----------------------------------------------------------------------------------'





'**********************************************************************************'
'PHP 암호화 키 START
'----------------------------------------------------------------------------------'
DIM PHP_KEY							'# PHP 암호화 키

PHP_KEY								 = "35e80f121fcae9fdb4d9a4d342e04f76"
'----------------------------------------------------------------------------------'
'NICE INFO END
'----------------------------------------------------------------------------------'



	

'**********************************************************************************'
'슈마커창고코드 START
'----------------------------------------------------------------------------------'
DIM B2B_WH_CODE						'# 슈마커 B2B 창고코드
DIM B2C_OUT_WH_CODE					'# 슈마커 B2C(외부몰) 창고코드
DIM B2C_SHOP_WH_CODE				'# 슈마커 B2C(온라인) 창고코드
DIM OUTLET_WH_CODE					'# 아울렛매장 코드
DIM DISUSE_WH_CODE					'# 슈마커 불용 창고코드
DIM LOSS_WH_CODE					'# 슈마커 조정 창고코드
DIM AS_WH_CODE						'# 슈마커 A/S반품 창고코드
B2B_WH_CODE							 = "11"
B2C_OUT_WH_CODE						 = "67"
B2C_SHOP_WH_CODE					 = "66"
OUTLET_WH_CODE						 = "63"
LOSS_WH_CODE						 = "69"
DISUSE_WH_CODE						 = "70"
AS_WH_CODE							 = "00"
'----------------------------------------------------------------------------------'
'파일 경로 END
'----------------------------------------------------------------------------------'





'**********************************************************************************'
'슈마커몰코드 START
'----------------------------------------------------------------------------------'
DIM SHOP_MALL_CODE					'# 슈마커 자사몰 코드
DIM OUTLET_MALL_CODE				'# 아울렛매장 코드
DIM OUT_MALL_CODE					'# 입점몰 통합매장 코드
SHOP_MALL_CODE						 = "006740"
OUTLET_MALL_CODE					 = "006750"
OUT_MALL_CODE						 = "009800"
'----------------------------------------------------------------------------------'
'파일 경로 END
'----------------------------------------------------------------------------------'





'**********************************************************************************'
'ERP DB & TABLE 정보 START
'----------------------------------------------------------------------------------'
DIM ERP_CON_STR
DIM ERP_LNK_SRV
DIM ERP_SYS_DAT

DIM ERP_ORD_TBL
DIM ERP_APP_TBL
DIM ERP_RST_TBL
DIM ERP_PRD_TBL
DIM ERP_PRM_TBL
DIM ERP_SHP_TBL
DIM ERP_DIF_TBL
DIM ERP_DIS_TBL
DIM ERP_STK_TBL
DIM ERP_STR_TBL
DIM ERP_OST_TBL
DIM ERP_SCL_TBL
DIM ERP_WRH_TBL
DIM ERP_WRD_TBL
DIM ERP_WQH_TBL
DIM ERP_WQD_TBL


'#ERP_CON_STR		 = "DSN=HOTTNEXT32DEV;UID=hottweb;PWD=webhott"
ERP_CON_STR		 = "DSN=HOTTNEXT32;UID=hottweb;PWD=webhott"
'#ERP_LNK_SRV		 = "HOTTNEXTDEV"
ERP_LNK_SRV		 = "HOTTNEXT"
ERP_ORD_TBL		 = "NEXTERP.IF_ONLINE_ORDER"
ERP_APP_TBL		 = "NEXTERP.IF_ONLINE_ORDER_APP"
ERP_RST_TBL		 = "NEXTERP.IF_ONLINE_ORDER_RESULT"
ERP_PRD_TBL		 = "NEXTERP.IF_ONLINE_PRODINFO"
ERP_PRM_TBL		 = "NEXTERP.PRODUCT"
ERP_SHP_TBL		 = "NEXTERP.IF_ONLINE_SHOPINFO"
ERP_DIF_TBL		 = "NEXTERP.DISTRIBUTION_INF"
ERP_DIS_TBL		 = "NEXTERP.DISTRIBUTION"
ERP_STK_TBL		 = "NEXTERP.STOCKREALREPAY"
ERP_STR_TBL		 = "NEXTERP.STOREREALREPAY"
ERP_OST_TBL		 = "NEXTERP.ONLINE_ORDER_V"
ERP_SCL_TBL		 = "NEXTERP.IF_ONLINE_STAFFCARDLIMIT_V"
ERP_WRH_TBL		 = "NEXTERP.IF_WMS_RETURNINFO_H"
ERP_WRD_TBL		 = "NEXTERP.IF_WMS_RETURNINFO_D"
ERP_WQH_TBL		 = "NEXTERP.IF_WMS_RETURNREQUEST_H"
ERP_WQD_TBL		 = "NEXTERP.IF_WMS_RETURNREQUEST_D"
ERP_SYS_DAT		 = "sysdate"



'# ERP_CON_STR		 = ConnectionString		'# /ADO/ADODBCommon.asp
'# ERP_LNK_SRV		 = "NEXTERP"
'# ERP_ORD_TBL		 = "NEXTERP.dbo.TEST_IF_ONLINE_ORDER"
'# ERP_APP_TBL		 = "NEXTERP.dbo.TEST_IF_ONLINE_ORDER_APP"
'# ERP_RST_TBL		 = "NEXTERP.dbo.TEST_IF_ONLINE_ORDER_RESULT"
'# ERP_PRD_TBL		 = "NEXTERP.dbo.IF_ONLINE_PRODINFO"
'# ERP_PRM_TBL		 = "NEXTERP.dbo.PRODUCT"
'# ERP_SHP_TBL		 = "NEXTERP.dbo.IF_ONLINE_SHOPINFO"
'# ERP_DIF_TBL		 = "NEXTERP.dbo.TEST_DISTRIBUTION_INF"
'# ERP_DIS_TBL		 = "NEXTERP.dbo.TEST_DISTRIBUTION"
'# ERP_STK_TBL		 = "NEXTERP.dbo.TEST_STOCKREALREPAY"
'# ERP_STR_TBL		 = "NEXTERP.dbo.TEST_STOREREALREPAY"
'# ERP_OST_TBL		 = "NEXTERP.dbo.TEST_ONLINE_ORDER_V"
'# ERP_SYS_DAT		 = "getdate()"
'----------------------------------------------------------------------------------'
'파일 경로 END
'----------------------------------------------------------------------------------'


'=================================================================================================='
'주문관련 정보 START
'--------------------------------------------------------------------------------------------------'
DIM MALL_MIN_ORDERPRICE				 '# 상품별 최소 주문금액
DIM MALL_REVIEW_POINT_B				 '# 일반상품후기 작성시 지급 포인트				
DIM MALL_REVIEW_POINT_P				 '# 포토상품후기 작성시 지급 포인트				
DIM MALL_OPENXPAY_NOTEURL			 '# 결제 처리  url(OpenXpay  멀티플랫폼)
DIM MALL_OPENXPAY_CASNOTEURL		 '# 가상계좌 결제 처리 url (OpenXpay  멀티플랫폼)
DIM MALL_OPENXPAY_RETURNURL			 '# 결제 리턴 url
DIM MALL_KVPMISPNOREURL				 '# ISP 카드결제 연동중 모바일ISP방식(고객세션을 유지하지않는 비동기방식)의 경우
DIM MALL_KVPMISPWAPURL				 '# ISP 카드결제 연동중 모바일ISP방식(고객세션을 유지하지않는 비동기방식)의 경우
DIM MALL_KVPMISPCANCELURL			 '# ISP 카드결제 연동중 모바일ISP방식(고객세션을 유지하지않는 비동기방식)의 경우
DIM MALL_LGD_ACCOUNTOWNER			 '# 계상계좌 입금계좌주명
DIM MALL_RECEIPT_LINK
DIM MALL_RECEIPT_LINK_TEST
DIM MALL_CLOSEDATE					 '# 가상계좌 입금시 입금 마감일
DIM MALL_ESCROW_LINK				 '# 에스크로 구매 연동 주소
DIM MALL_ESCROW_LINK_TEST			 '# 에스크로 구매 연동 주소(테스트)
DIM MALL_ESCROW_DELIVERY_URL_TEST	 '# 에스크로 배송 처리 테스트 url	
DIM MALL_ESCROW_DELIVERY_URL		 '# 에스크로 배송 처리 url
DIM MALL_PARTIALCANCEL_PATH			 '# 부분취소 환경설정 파일 위치

MALL_MIN_ORDERPRICE					 = 1000
MALL_REVIEW_POINT_B					 = 500
MALL_REVIEW_POINT_P					 = 1000

MALL_OPENXPAY_NOTEURL				 = HOME_DOMAIN & "/Common/OpenXpay/note_url.asp"
MALL_OPENXPAY_CASNOTEURL			 = HOME_DOMAIN & "/Common/OpenXpay/cas_noteurl.asp"
MALL_OPENXPAY_RETURNURL				 = HOME_DOMAIN & "/Common/OpenXpay/returnurl.asp"
MALL_KVPMISPNOREURL					 = HOME_DOMAIN & "/Common/OpenXpay/note_url.asp"
MALL_KVPMISPWAPURL					 = HOME_DOMAIN & "/Common/OpenXpay/mispwapurl.asp"
MALL_KVPMISPCANCELURL				 = HOME_DOMAIN & "/Common/OpenXpay/cancel_url.asp"

MALL_LGD_ACCOUNTOWNER				 = "(주)슈마커"
MALL_RECEIPT_LINK_TEST				 = "http://pgweb.dacom.net:7085/WEB_SERVER/js/receipt_link.js"
MALL_RECEIPT_LINK					 = "//pgweb.lgtelecom.com/WEB_SERVER/js/receipt_link.js"
MALL_CLOSEDATE						 = 7	'//입금마감일은 기본 7일로 설정

MALL_ESCROW_LINK					 = "//pgweb.dacom.net/js/DACOMEscrow_UTF8.js"
MALL_ESCROW_LINK_TEST				 = "//pgweb.dacom.net:7085/js/DACOMEscrow_UTF8.js"
MALL_ESCROW_DELIVERY_URL_TEST		 = "//pgweb.dacom.net:7085/pg/wmp/mertadmin/jsp/escrow/rcvdlvinfo.jsp"
MALL_ESCROW_DELIVERY_URL			 = "//pgweb.dacom.net/pg/wmp/mertadmin/jsp/escrow/rcvdlvinfo.jsp"

MALL_PARTIALCANCEL_PATH				 = "/Common/Pay/PartialCancel/lgdacom"
'--------------------------------------------------------------------------------------------------'
'주문관련 정보 END
'--------------------------------------------------------------------------------------------------'


'=================================================================================================='
'LGU+ 정보 START
'--------------------------------------------------------------------------------------------------'
DIM CST_MID					'# LGU+ 상점아이디
DIM CST_MID_TEST			'# LGU+ 상점아이디
DIM LGD_MERTKEY				'# 머트키값
DIM PAY_PLATFORM

CST_MID						 = "shoemarker01"
CST_MID_TEST				 = "tshoemarker01"
LGD_MERTKEY					 = "4733a4aacbd9937e32e6f48d2734d1bd"
PAY_PLATFORM				 = "service"
'PAY_PLATFORM				 = "test"
'--------------------------------------------------------------------------------------------------'
'LGU+ 정보 END
'--------------------------------------------------------------------------------------------------'


'=================================================================================================='
'네이버페이 정보 START
'--------------------------------------------------------------------------------------------------'
DIM NAVER_PAY_FLAG					'# 네이버 페이 사용여부
DIM NAVER_PAY_API_DOMAIN			'# 네이버 페이 API 도메인
DIM NAVER_PAY_ID					'# 네이버 페이 파트너ID
DIM NAVER_PAY_CLIENTID				'# 네이버 페이 CLIENTID
DIM NAVER_PAY_CLIENTSECRET			'# 네이버 페이 CLIENT SECRET
DIM NAVER_PAY_PLATFORM				'# 연동모드
DIM NAVER_PAY_PAYMENTURL			'# 네이버페이 결제 승인 URL
DIM NAVER_PAY_CANCELURL				'# 네이버페이 결제 취소 URL
DIM NAVER_PAY_SAVEURL				'# 네이버페이 포인트 적립요청 URL
DIM NAVER_PAY_CONFIRMURL			'# 네이버페이 거래완료 URL
DIM NAVER_PAY_RETURNURL				'# 네이버페이 결제 리턴 URL

NAVER_PAY_FLAG						= "N"
NAVER_PAY_ID						= "shoemarker01"
NAVER_PAY_CLIENTID					= "6Y13_s20u_fyw7VKPjEG"
NAVER_PAY_CLIENTSECRET				= "iGpluNa9bC"
NAVER_PAY_PLATFORM					= "production"		'# 운영모드
'NAVER_PAY_PLATFORM					= "development"		'# 개발모드

IF NAVER_PAY_PLATFORM = "production" THEN
		NAVER_PAY_API_DOMAIN				= "https://apis.naver.com/"
ELSE
		NAVER_PAY_API_DOMAIN				= "https://dev.apis.naver.com/"
END IF

NAVER_PAY_PAYMENTURL				= NAVER_PAY_API_DOMAIN & NAVER_PAY_ID & "/naverpay/payments/v2.2/apply/payment"
NAVER_PAY_CANCELURL					= NAVER_PAY_API_DOMAIN & NAVER_PAY_ID & "/naverpay/payments/v1/cancel"
NAVER_PAY_SAVEURL					= NAVER_PAY_API_DOMAIN & NAVER_PAY_ID & "/naverpay/payments/v1/naverpoint-save"
NAVER_PAY_CONFIRMURL				= NAVER_PAY_API_DOMAIN & NAVER_PAY_ID & "/naverpay/payments/v1/purchase-confirm"

NAVER_PAY_RETURNURL					= HOME_DOMAIN & "/Common/NaverPay/ReturnUrl.asp"

IF U_IP = "1.215.226.130" THEN
		NAVER_PAY_FLAG	= "Y"
END IF
'# 네이버페이 검수용 IP
IF U_ID = "211.249.70.25" OR U_ID = "211.249.71.218" OR U_ID = "220.230.184.40" THEN
		NAVER_PAY_FLAG	= "Y"
END IF
'--------------------------------------------------------------------------------------------------'
'LGU+ 정보 END
'--------------------------------------------------------------------------------------------------'




'=================================================================================================='
'UClick 정보 START
'--------------------------------------------------------------------------------------------------'
DIM USAFE_FLAG       '# USafe 사용여부
DIM USAFE_PAYTYPE	 '# USafe 결제수단
DIM USAFE_ID		 '# UClick 아이디
DIM USAFE_AMT		 '# 보증보험 적용 최저금액

USAFE_FLAG           = "N"
USAFE_PAYTYPE        = "B"
USAFE_ID			 = "shoemarker"
USAFE_AMT			 = 0
'--------------------------------------------------------------------------------------------------'
'UClick 정보 END
'--------------------------------------------------------------------------------------------------'
%>