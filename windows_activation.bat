:: ----------------------------------- 1、初始化 -----------------------------------
@echo off
::mode con cols=75 lines=300
title Windows/Office KMS激活工具
setlocal EnableDelayedExpansion&color 70 & cd /d "%~dp0"
%1 %2
ver|find "5.">nul && goto :start
mshta vbscript:createobject("shell.application").shellexecute("%~s0","goto :start","","runas",1)(window.close)&goto :eof
 
:start
echo =========================== 运行说明 ==================================
echo Welcome 
echo =======================================================================
 
echo 【A】一键激活Windows和多版本Office（使用沧水KMS）
echo 【B】一键激活Windows和多版本Office（使用自定义KMS）
echo 【C】激活Windows
echo 【D】激活多版本Office
echo 【E】清除Windows的KMS
echo 【F】清除多版本Office的KMS
echo 【G】退出
 
set defaultKMS=http://kms.cangshui.net/
set /p xuanze= 请输入你的选择:
if /i "%xuanze%"=="a" goto startActiveAllWithDefaultKMS
if /i "%xuanze%"=="b" goto startActiveAllWithPrivateKMS
if /i "%xuanze%"=="c" goto startActiveWindows
if /i "%xuanze%"=="d" goto startActiveOffice
if /i "%xuanze%"=="e" goto startCleanWindowsKMS
if /i "%xuanze%"=="f" goto startCleanOfficeKMS
if /i "%xuanze%"=="g" exit
cls
goto start
 
:: ----------------------------------- 2、通用函数 -----------------------------------
:info
echo.
echo *************************** %1 ***************************
echo %2
echo ****************************************************************
echo.
goto :EOF
 
:infoExit
call :info %1 %2
goto end
 
:end
pause
exit 
 
:setKMSServer
if DEFINED KMSServer (echo KMS服务器为：!KMSServer! ) else ( 
  set /p KMSServer= 请输入KMS服务器域名或IP，如不输入默认直接使用沧水KMS:
  if "!KMSServer!"=="" set KMSServer=%defaultKMS%
  echo KMS服务器成功设置为：!KMSServer!
)
goto :EOF
 
:checkNetwork
echo 正在检查本机网络情况耐心等待……
ping https://www.baidu.com | find "超时"  > NUL && call :infoExit 网络故障 无法连接到服务器，请检查网络是否通畅。
ping https://www.baidu.com | find "目标主机"  > NUL &&  call :infoExit 网络故障 无法连接到服务器，请检查网络是否通畅。
echo 本机网络良好
goto :EOF
 
 
:: ----------------------------------- 3、主要脚本 -----------------------------------
:startActiveAllWithDefaultKMS
set KMSServer=%defaultKMS%
call :windows
call :office
goto end
 
:startActiveAllWithPrivateKMS
set KMSServer=
call :windows
call :office
goto end
 
:startActiveWindows
call :windows
goto end
 
 
:startActiveOffice
call :office
goto end
 
:: ----------------------------------- 4、Windows激活 -----------------------------------
:windows
call :checkNetwork
call :setKMSServer
if not DEFINED finished ( ver | find "6.0." > NUL &&  call :winvista )
if not DEFINED finished ( ver | find "6.1." > NUL &&  call :win7 )
if not DEFINED finished ( ver | find "6.2." > NUL &&  call :win8 )
if not DEFINED finished ( ver | find "6.3." > NUL &&  call :win81 )
if not DEFINED finished ( ver | find "10.0." > NUL &&  call :win10 )
if not DEFINED finished ( call :infoExit 激活失败 Windows激活失败，未找到合适的NT6系统，可能是WinXP或Win2003。 )
goto :EOF
 
:winvista
echo 当前操作系统为Windows Vista/2008。
set Business=YFKBB-PQJJV-G996G-VWGXY-2V3X8
set BusinessN=HMBQG-8H2RH-C77VX-27R82-VMQBT
set Enterprise=VKK3X-68KWM-X2YGT-QR4M6-4BWMV
set EnterpriseN=VTC42-BM838-43QHV-84HX6-XJXKV
set ServerWeb=WYR28-R7TFJ-3X2YQ-YCY4H-M249D
set ServerStandard=TM24T-X9RMF-VWXK6-X8JC9-BFGM2
set ServerStandardV=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
set ServerEnterprise=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
set ServerEnterpriseV=39BXF-X8Q23-P2WWT-38T2F-G3FPG
set ServerWeb=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
set ServerDatacenter=7M67G-PC374-GR742-YH8V4-TCBY3
set ServerDatacenterV=22XQ2-VRXRG-P8D42-K34TD-G3QQC
set ServerEnterpriseIA64=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
call :ActWindows
goto :EOF
:win7
echo 当前操作系统为Windows 7/2008 R2。
set Professional=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
set ProfessionalN=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
set ProfessionalE=W82YF-2Q76Y-63HXB-FGJG9-GF7QX
set Enterprise=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
set EnterpriseN=YDRBP-3D83W-TY26F-D46B2-XCKRJ
set EnterpriseE=C29WB-22CC8-VJ326-GHFJW-H9DH4
set ServerWeb=6TPJF-RBVHG-WBW2R-86QPH-6RTM4
set ServerHPC=TT8MH-CG224-D3D7Q-498W2-9QCTX
set ServerStandard=YC6KT-GKW9T-YTKYR-T4X34-R7VHC
set ServerEnterprise=489J6-VHDMP-X63PK-3K798-CPX3Y
set ServerDatacenter=74YFP-3QFB3-KQT8W-PMXWJ-7M648
set ServerEnterpriseIA64=GT63C-RJFQ3-4GMB6-BRFB9-CB83V
call :ActWindows
goto :EOF
:win8
echo 当前操作系统为Windows 8/2012。
set Professional=NG4HW-VH26C-733KW-K6F98-J8CK4
set ProfessionalN=XCVCF-2NXM9-723PB-MHCB7-2RYQQ
set Core=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
set Enterprise=32JNW-9KQ84-P47T8-D8GGY-CWCK7
set EnterpriseN=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
set CoreN=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
set CoreSingleLanguage=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
set CoreCountrySpecific=4K36P-JN4VD-GDC6V-KDT89-DYFKP
set ServerMultiPointPremium=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
set ServerMultiPointStandard=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
set ServerStandard=XC9B7-NBPP2-83J2H-RHMBY-92BT4
set ServerDatacenter=48HP8-DN98B-MYWDG-T2DCC-8W83P
call :ActWindows
goto :EOF
:win81
echo 当前操作系统为Windows 8.1。
set Professional=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
set ProfessionalN=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
set Enterprise=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
set EnterpriseN=TT4HM-HN7YT-62K67-RGRQJ-JFFXW
set ServerSolution=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
set ServerStandard=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
set ServerDatacenter=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
set EmbeddedIndustry=32JNW-9KQ84-P47T8-D8GGY-CWCK7
call :ActWindows
goto :EOF
:win10
echo 当前操作系统为Windows 10/Server 2016。
set Core=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
set CoreCountrySpecific=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
set CoreN=3KHY7-WNT83-DGQKR-F7HPR-844BM
set CoreSingleLanguage=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
set Professional=W269N-WFGWX-YVC9B-4J6C9-T83GX
set ProfessionalN=MH37W-N47XK-V7XM9-C7227-GCQG9
set Enterprise=NPPR9-FWDCX-D2C8J-H872K-2YT43
set EnterpriseN=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
set Education=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
set EducationN=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
set EnterpriseS=WNMTR-4C88C-JK8YV-HQ7T2-76DF9
set EnterpriseSN=2F77B-TNFGY-69QQF-B8YKP-D69TJ
set ServerDatacenter=CB7KF-BWN84-R7R2Y-793K2-8XDDG
set ServerStandard=WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY
set ServerEssentials=JCKRF-N37P4-C2D82-9YXRT-4M63B
set EnterpriseG=YYVX9-NTFWV-6MDM3-9PT4T-4M68B
set EnterpriseGN=44RPN-FTY23-9VTTB-MP9BX-T84FV
set ProfessionalEducation=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
set ProfessionalEducationN=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
set ProfessionalWorkstation=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstations=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstationsN=9FNHH-K3HBT-3W4TD-6383H-6XYWF
call :ActWindows
goto :EOF
 
:ActWindows
for /f "tokens=3 delims= " %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionID"') do set EditionID=%%i
echo !%EditionID%!
if defined %EditionID% (
	cscript //Nologo %windir%\system32\slmgr.vbs /ipk !%EditionID%!
	cscript //Nologo %windir%\system32\slmgr.vbs /skms !KMSServer!
	cscript //Nologo %windir%\system32\slmgr.vbs /ato
) else (
	call :infoExit 激活失败 Windows激活失败，找不到序列号，可能是旗舰版之类的系统。
)
call :info 激活成功 您的Windows已激活成功。
set finished="True"
goto :EOF
 
:: ----------------------------------- 5、Office激活 -----------------------------------
:office
call :checkNetwork
call :setKMSServer
echo 检查安装的office……
call :getOfficePath 14 Office2010
call :ActOffice 14 Office2010
call :getOfficePath 15 Office2013
call :ActOffice 15 Office2013
call :getOfficePath 16 Office2016
call :ActOffice 16 Office2016
goto :EOF
 
:ActOffice
if DEFINED _Office%1Path (
    cd /d "!_Office%1Path!"
	if %1 EQU 14 call :license2010
	if %1 EQU 15 call :license2013
    if %1 EQU 16 call :license2016
    echo 尝试激活 %2 ……
    cscript //nologo ospp.vbs /sethst:!KMSServer! >nul
    cscript //nologo ospp.vbs /act | find /i "successful" && (call :info 激活成功 您的%2已经激活成功) || (call :infoExit 激活失败 %2激活失败)
)    
 
cd /d "%~dp0"
goto :EOF
 
:getOfficePath
echo 正在检测 %2 系列产品的安装路径……
set _Office%1Path=
set _Reg32=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot
set _Reg64=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot
reg query "%_Reg32%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg32%" /v "Path"') do SET "_OfficePath1=%%b"
reg query "%_Reg64%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg64%" /v "Path"') do SET "_OfficePath2=%%b"
if DEFINED _OfficePath1 (if exist "%_OfficePath1%ospp.vbs" set _Office%1Path=!_OfficePath1!)
if DEFINED _OfficePath2 (if exist "%_OfficePath2%ospp.vbs" set _Office%1Path=!_OfficePath2!)
set _OfficePath1=
set _OfficePath2=
if DEFINED _Office%1Path (echo 已发现 %2) else (echo 未发现 %2)
goto :EOF
 
:license2010
cscript ospp.vbs /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >nul
cscript ospp.vbs /inpkey:V7QKV-4XVVR-XYV4D-F7DFM-8R6BM >nul
cscript ospp.vbs /inpkey:D6QFG-VBYP2-XQHM7-J97RH-VVRCK >nul
cscript ospp.vbs /inpkey:V7Y44-9T38C-R2VJK-666HK-T7DDX >nul
cscript ospp.vbs /inpkey:H62QG-HXVKF-PP4HP-66KMR-CW9BM >nul
cscript ospp.vbs /inpkey:QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4 >nul
cscript ospp.vbs /inpkey:K96W8-67RPQ-62T9Y-J8FQJ-BT37T >nul
cscript ospp.vbs /inpkey:Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX >nul
cscript ospp.vbs /inpkey:7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ >nul
cscript ospp.vbs /inpkey:RC8FX-88JRY-3PF7C-X8P67-P4VTT >nul
cscript ospp.vbs /inpkey:YGX6F-PGV49-PGW3J-9BTGG-VHKC6 >nul
cscript ospp.vbs /inpkey:BFK7F-9MYHM-V68C7-DRQ66-83YTP >nul
cscript ospp.vbs /inpkey:HVHB3-C6FV7-KQX9W-YQG79-CRY7T >nul
cscript ospp.vbs /inpkey:D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ >nul
cscript ospp.vbs /inpkey:7MCW8-VRQVK-G677T-PDJCM-Q8TCP >nul
cscript ospp.vbs /inpkey:767HD-QGMWX-8QTDB-9G3R2-KHFGJ >nul
goto :EOF
 
:license2013
cscript ospp.vbs /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT >nul
cscript ospp.vbs /inpkey:KBKQT-2NMXY-JJWGP-M62JB-92CD4 >nul
cscript ospp.vbs /inpkey:FN8TT-7WMH6-2D4X9-M337T-2342K >nul
cscript ospp.vbs /inpkey:6NTH3-CW976-3G3Y2-JK3TX-8QHTT >nul
cscript ospp.vbs /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3 >nul
cscript ospp.vbs /inpkey:J484Y-4NKBF-W2HMG-DBMJC-PGWR7 >nul
cscript ospp.vbs /inpkey:NG2JY-H4JBT-HQXYP-78QH9-4JM2D >nul
cscript ospp.vbs /inpkey:VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB >nul
cscript ospp.vbs /inpkey:DKT8B-N7VXH-D963P-Q4PHY-F8894 >nul
cscript ospp.vbs /inpkey:2MG3G-3BNTT-3MFW9-KDQW3-TCK7R >nul
cscript ospp.vbs /inpkey:TGN6P-8MMBC-37P2F-XHXXK-P34VW >nul
cscript ospp.vbs /inpkey:QPN8Q-BJBTJ-334K3-93TGY-2PMBT >nul
cscript ospp.vbs /inpkey:4NT99-8RJFH-Q2VDH-KYG2C-4RD4F >nul
cscript ospp.vbs /inpkey:PN2WF-29XG2-T9HJ7-JQPJR-FCXK4 >nul
cscript ospp.vbs /inpkey:6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7 >nul
goto :EOF
 
 
:license2016
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul
cscript ospp.vbs /inpkey:NYH39-6GMXT-T39D4-WVXY2-D69YY >nul
cscript ospp.vbs /inpkey:7WHWN-4T7MP-G96JF-G33KR-W8GF4 >nul
cscript ospp.vbs /inpkey:RBWW7-NTJD4-FFK2C-TDJ7V-4C2QP >nul
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT >nul
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK >nul
goto :EOF
 
:: ----------------------------------- 6、清除Office激活 -----------------------------------
 
:startCleanOfficeKMS
set /p xuanze= 是否真的要清除Office的KMS激活？（【Y】继续/【N】关闭）：
if /i "%xuanze%"=="y" goto nextun
if /i "%xuanze%"=="n" goto end
 
 
:nextun
call :getOfficePath 14 Office2010
call :cleanOffice 14 Office2010
call :getOfficePath 15 Office2013
call :cleanOffice 15 Office2013
call :getOfficePath 16 Office2016
call :cleanOffice 16 Office2016
call :infoExit 清除完成 您的OfficeKMS信息已从电脑上成功移除。
  
:cleanOffice
if DEFINED _Office%1Path (
    echo 清理%2的Key……
    cd /d "!_Office%1Path!"
	for /f "tokens=8 delims= " %%a in ('cscript //nologo ospp.vbs /dstatus ^| find "product key"') do ( 
	    cscript //nologo ospp.vbs /unpkey:%%a >nul
	)
	echo 移除KMS……
	cscript //nologo ospp.vbs /remhst >nul
)
goto :EOF
 
:: ----------------------------------- 7、清除Windows KMS -----------------------------------
 
:startCleanWindowsKMS
set /p xuanze= 是否真的要清除Windows的KMS激活？（【Y】继续/【N】关闭）：
if /i "%xuanze%"=="y" goto nextunw
if /i "%xuanze%"=="n" goto end
:nextunw
slmgr /upk
::slmgr /ckms
::slmgr /rearm
call :infoExit 清除完成 您的WindowsKMS信息已从电脑上成功移除。
