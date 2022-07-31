#!/bin/bash
set -x
############
# SETTINGS #
############

PYTHON_PATH="`find /usr/local/Cellar/python@3.9 -type f -name python3.9 | head -n1`"
PIP_PATH="`find /usr/local/Cellar/python@3.9 -type f -name pip3.9 | head -n1`"
APP_NAME='Calendar'

PYTHON_VERSION="`${PYTHON_PATH} --version | cut -d' ' -f2`"
PYTHON_EXEC_VERSION="`echo ${PYTHON_VERSION} | cut -d. -f1-2`"

########
# INFO #
########

# print some info for debugging failed builds
uname -a
sw_vers
which python2
python2 --version
which python3
python3 --version
${PYTHON_PATH} --version
echo $PATH
pwd
ls -lah

###################
# INSTALL DEPENDS #
###################

brew install sdl2 sdl2_image sdl2_ttf sdl2_mixer
${PIP_PATH} install setuptools wheel kivy openpyxl icalendar
${PIP_PATH} install pyinstaller==3.6

#####################
# PYINSTALLER BUILD #
#####################

mkdir pyinstaller
pushd pyinstaller

cat >> ${APP_NAME}.spec <<EOF
# -*- mode: python ; coding: utf-8 -*-
block_cipher = None
a = Analysis(['../src/main.py'],
             pathex=['./'],
             binaries=[],
             datas=[],
             hiddenimports=['pkg_resources.py2_warn'],
             hookspath=[],
             runtime_hooks=[],
             excludes=['_tkinter', 'Tkinter', 'enchant', 'twisted'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='${APP_NAME}',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False )
coll = COLLECT(exe, Tree('../src/'),
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='${APP_NAME}')
app = BUNDLE(coll,
             name='${APP_NAME}.app',
             icon=None,
             bundle_identifier=None)
EOF

${PYTHON_PATH} -m PyInstaller -y --clean --windowed "${APP_NAME}.spec"

pushd dist
hdiutil create ./${APP_NAME}.dmg -srcfolder ${APP_NAME}.app -ov
popd

#####################
# PREPARE ARTIFACTS #
#####################

# create the dist dir for our result to be uploaded as an artifact
mkdir -p ../dist
cp "dist/${APP_NAME}.dmg" ../dist/

#######################
# OUTPUT VERSION INFO #
#######################


uname -a
sw_vers
which python2
python2 --version
which python3
python3 --version
${PYTHON_PATH} --version
echo $PATH
pwd
ls -lah
ls -lah dist

##################
# CLEANUP & EXIT #
##################

# exit cleanly
exit 0