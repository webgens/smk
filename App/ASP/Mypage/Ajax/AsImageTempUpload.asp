<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'AsImageTempUpload.asp - A/S 이미지 임시저장 파일 삭제
'Date		: 2019.01.23
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM UF
Dim UFImage
Dim Status
DIM SaveFolder
DIM FSO
DIM FileExt
DIM SaveFile


DIM Orientation
DIM Rotate
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


SET UF			 = Server.CreateObject("TABSUpload4.Upload")
UF.CodePage		 = "65001"
UF.Start Server.MapPath(D_UPLOAD)

SET UFImage		 = Server.CreateObject("TABSUpload4.Image")

SaveFolder = D_ORDERAS & "Temp/"

IF UF("FileName").FileSize > 0 THEN
		FileExt = Mid(UF("FileName").FileName, Instr(UF("FileName").FileName, ".") + 1)
		SaveFile = UF("FileName").SaveAs(Server.MapPath(SaveFolder) & "\" & U_DATE & U_TIME & right("000" & (timer * 1000) Mod 1000, 3) & "." & FileExt, False)
		SaveFile = UF("FileName").ShortSaveName

		Status = UFImage.Load(Server.MapPath(SaveFolder) & "\" & SaveFile)
		IF Status = 0 THEN
				Orientation = UFImage.Metadata.GetExifValue(274)
				'# 이미지 회전각도 설정
				Rotate = 0
				IF Orientation = "3" THEN
					Ratate = 90
				ELSEIF Orientation = "8" THEN
					Rotate = 180
				ELSEIF Orientation = "6" THEN
					Rotate = 270
				END IF
				UFImage.Rotate Rotate, "#00000000"
				UFImage.Save Server.MapPath(SaveFolder) & "\" & SaveFile, 100, True
		ELSE
				UFImage.Close
			
				SET UFImage = Nothing
				SET UF = Nothing
			
				Response.Write "FAIL|||||이미지 처리 도중 오류가 발생하였습니다."
				Response.End
		END IF

		Response.Write "OK|||||" & SaveFolder & "^^^^^" & SaveFile
ELSE
		Response.Write "FAIL|||||선택된 이미지가 없습니다."
END IF

UFImage.Close
SET UFImage = Nothing
SET UF = Nothing
%>