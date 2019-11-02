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

# get latest version
get_latest_release_number() {
  curl --silent "https://github.com/$1/releases/latest" | sed 's#.*tag/\(.*\)".*#\1#'
}
VLMCSD_LATEST_VERSION=`get_latest_release_number "Wind4/vlmcsd"`

# download file
wget https://github.com/Wind4/vlmcsd/releases/download/${VLMCSD_LATEST_VERSION}/binaries.tar.gz -O ${VLMCSD_LATEST_VERSION}_binaries.tar.gz

# install
tar -zxvf ${VLMCSD_LATEST_VERSION}_binaries.tar.gz
mv ${VLMCSD_LATEST_VERSION}_binaries /opt/vlmcsd
chmod -R 777 /opt/vlmcsd

# install service
cat > /etc/init.d/vlmcsd <<-"EOF"
#!/usr/bin/env bash
# chkconfig: 2345 90 10
# description: kms server by vlmcsd.

### BEGIN INIT INFO
# Provides:          vlmcsd
# Default-Start:     2 3 4 5
# Default-Stop:      0 1 6
# Short-Description: kms server by vlmcsd
# Description:       Start or stop the kms server
### END INIT INFO

DAEMON=/opt/vlmcsd/Linux/intel/static/vlmcsd-x64-musl-static
NAME=vlmcsd-x64-musl-static
PID_DIR=/var/run
PID_FILE=$PID_DIR/vlmcsd-x64-musl-static.pid
RET_VAL=0

[ -x $DAEMON ] || exit 0

if [ ! -d $PID_DIR ]; then
    mkdir -p $PID_DIR
    if [ $? -ne 0 ]; then
        echo "Creating PID directory $PID_DIR failed"
        exit 1
    fi
fi

check_running() {
    if [ -r $PID_FILE ]; then
        read PID < $PID_FILE
        if [ -d "/proc/$PID" ]; then
            return 0
        else
            rm -f $PID_FILE
            return 1
        fi
    else
        return 2
    fi
}

do_status() {
    check_running
    case $? in
        0)
        echo "$NAME (pid $PID) is running..."
        ;;
        1|2)
        echo "$NAME is stopped"
        RET_VAL=1
        ;;
    esac
}

do_start() {
    if check_running; then
        echo "$NAME (pid $PID) is already running..."
        return 0
    fi
    $DAEMON
    $PID=`ps -ef | grep vlmcsd-x64-musl-static | grep -v grep | awk '{print $2}'`
    echo $PID > $PID_FILE
    if check_running; then
        echo "Starting $NAME success"
    else
        echo "Starting $NAME failed"
        RET_VAL=1
    fi
}

do_stop() {
    if check_running; then
        kill -9 $PID
        rm -f $PID_FILE
        echo "Stopping $NAME success"
    else
        echo "$NAME is stopped"
        RET_VAL=1
    fi
}

do_restart() {
    do_stop
    sleep 0.5
    do_start
}

case "$1" in
    start|stop|restart|status)
    do_$1
    ;;
    *)
    echo "Usage: $0 { start | stop | restart | status }"
    RET_VAL=1
    ;;
esac

exit $RET_VAL
EOF
chmod 777 /etc/init.d/vlmcsd
chkconfig vlmcsd on

# start
service vlmcsd start
service vlmcsd status


