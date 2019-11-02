# kms_server
kms server

# depend on vlmcsd
https://github.com/Wind4/vlmcsd

# install
```shell
git clone https://github.com/easonjim/kms_server.git
cd kms_server
bash install_kms_server.sh
```

# uninstall
```shell
bash uninstall_kms_server.sh
```

# vlmcsd service command
```shell
# start
service vlmcsd start
# stop
service vlmcsd stop
# status
service vlmcsd status
```

# use
```shell
# check
{VLMCSD_PATH}\Windows\intel>vlmcs-Windows-x64.exe -v -l 3 ${IP/Domain}
# windows
cd /d "%SystemRoot%\system32"
## set gvlk key
slmgr /ipk ${GVLK Key}
## set kms server
slmgr /skms ${IP/Domain}
## activation
slmgr /ato
## get activation status
slmgr /xpr
# office
## 64 bit office
cd /d "%ProgramFiles%\Microsoft Office\${Office15}"
## 32 bit office
cd /d "%ProgramFiles(x86)%\Microsoft Office\#{Office15}"
## set gvlk key
cscript ospp.vbs /inpkey:${GVLK KEY}
## set kms server
cscript ospp.vbs /sethst:${IP/Domain}
## activation
cscript ospp.vbs /act
## get activation status
cscript ospp.vbs /dstatus
```

# GVLK key
## windows
https://docs.microsoft.com/zh-cn/windows-server/get-started/kmsclientkeys
## office
https://docs.microsoft.com/zh-cn/previous-versions/office/office-2010/ee624355(v=office.14%29
https://docs.microsoft.com/zh-cn/previous-versions/office/dn385360(v=office.15%29
https://docs.microsoft.com/en-us/deployoffice/vlactivation/gvlks
