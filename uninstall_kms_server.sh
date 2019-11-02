#!/bin/bash
# 
# kms server by vlmcsd

# 解决相对路径问题
cd `dirname $0`

# 检查是否为root用户，脚本必须在root权限下运行
if [[ "$(whoami)" != "root" ]]; then
    echo "please run this script as root !" >&2
    exit 1
fi

# remove vlmcsd
service vlmcsd stop
chkconfig vlmcsd off
rm /etc/init.d/vlmcsd
rm -rf /opt/vlmcsd
