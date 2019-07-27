#!/bin/bash

#Here we define pretty colors

RESTORE='\033[0m'
LGREEN='\033[01;32m'

RESET='\e[0m'

echo -e "${LGREEN}"
echo "╔════════════════════════════════════════════════════════════════╗"
echo "║      Enter a file path you wish to smuggle through macro       ║"              
echo "╚════════════════════════════════════════════════════════════════╝"
echo -e "${RESTORE}"
read TARGET

cat $TARGET | base64 > base64_output.txt
echo -e "${LGREEN}"
echo "[+]Base64 encoded!"
RESTORE='\033[0m'

echo -e "${LGREEN}"
echo "╔════════════════════════════════════════════════════════════════╗"
echo "║              Adding VB variables to the file                   ║"              
echo "╚════════════════════════════════════════════════════════════════╝"
echo -e "${RESTORE}"

sed 's/^/''base = base + "''/' base64_output.txt > variables
sed -e 's/$/"/' variables > macro
sed -i '1iDim var1' macro
echo -e "${LGREEN}"
echo "[+]Added variables and saved to a macro file"
echo -e "${RESTORE}"
rm -f variables
rm -f encoded*

echo -e "${LGREEN}"
echo "╔════════════════════════════════════════════════════════════════╗"
echo "║     Macro is ready, copy and paste it in the VBA editor        ║"              
echo "╚════════════════════════════════════════════════════════════════╝"
echo -e "${RESTORE}"
rm -f base*

