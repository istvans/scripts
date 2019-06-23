#!/bin/bash

BRANCH=${1:-Leia}
SEP="==============================================================="
CMAKE_COMMAND="cmake -DADDONS_TO_BUILD=pvr.waipu -DADDON_SRC_PREFIX=../.. -DCMAKE_BUILD_TYPE=Debug -DCMAKE_INSTALL_PREFIX=../../xbmc/addons -DPACKAGE_ZIP=1 ../../xbmc/cmake/addons"

sudo apt-get install build-essential  git autoconf cmake zip -y
git clone --branch $BRANCH https://github.com/xbmc/xbmc.git
git clone --branch $BRANCH https://github.com/flubshi/pvr.waipu.git
cd pvr.waipu && mkdir build && cd build

echo "### expected cmake failure, but we get the directory structure we need..."
$CMAKE_COMMAND
echo $SEP

echo "#### adding waipu to the supported addons' list..."
cd ../../xbmc/cmake/addons/addons
cp -R pvr.zattoo pvr.waipu
mv pvr.waipu/pvr.zattoo.txt pvr.waipu/pvr.waipu.txt
echo "pvr.waipu https://github.com/flubshi/pvr.waipu $BRANCH" > pvr.waipu/pvr.waipu.txt
echo $SEP

cd ../../../../pvr.waipu/build
echo "### make up the missing rapidjson.sha256 file..."
wget https://raw.githubusercontent.com/rbuehlma/pvr.zattoo/$BRANCH/depends/common/rapidjson/rapidjson.sha256 -P ../depends/common/rapidjson/
echo $SEP

$CMAKE_COMMAND
make
