#!/bin/sh
# DigiCalc build (Linux):  ./build.sh [hbmk2 options, e.g. -rebuild]
# HBMK2  hbmk2 to use (default: hbmk2 in the PATH, else /opt/harbour-core/bin/linux/gcc/hbmk2)
# Needs ../Harbour-CUI (the hbtui library built: cd ../Harbour-CUI && hbmk2 hbtui.hbp)
# HBLIB  a folder with contrib libraries (hbct) when Harbour was built without them
#        (HB_BUILD_CONTRIBS=no); default ~/hblib when it is there
cd "$(dirname "$0")" || exit 1
if [ -z "$HBMK2" ]; then
    if command -v hbmk2 >/dev/null 2>&1; then HBMK2=hbmk2; else HBMK2=/opt/harbour-core/bin/linux/gcc/hbmk2; fi
fi
[ -f ../Harbour-CUI/lib/linux/gcc/libhbtui.a ] || { echo "Build Harbour-CUI first: (cd ../Harbour-CUI && $HBMK2 hbtui.hbp)"; exit 1; }
HBLIB=${HBLIB:-$HOME/hblib}
LOPT=
[ -d "$HBLIB" ] && LOPT="-L$HBLIB"
exec "$HBMK2" digicalc.hbp $LOPT "$@"
