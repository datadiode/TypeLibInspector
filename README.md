Main Script
-----------
./src/au3/TypeLibInspector.au3 - 32-bit
./src/au3/TypeLibInspector64.au3 - 64-bit

Directory Structure
-------------------
./build - used for compiled scripts
./conf - containes configuration files
./conf/plugins - configuration for TypeLibInspector plugins
./res - application resources
./res/css - cascading stylesheet files
./res/html - HTML files
./res/icons - icon files
./res/xsl - XML stylesheets
./src - program source files
./src/au3 - AU3 sources for TypeLibInspector
./src/pkg - distribution packager sources

The project's directory structure is purposed to be used in "development mode",
i.e. application script runs within this structure out-of-the-box only uncompiled
and executed from ./src/au3 directory. When compiled TypeLibInspector expects
its resources to be structured as follows

Distribution Structure
----------------------
./TypeLibInspector.exe
./plugins - configuration for TypeLibInspector plugins
./css - cascading stylesheet files
./html - HTML files
./icons - icon files
./xsl - XML stylesheets

The script ./src/pkg/Packager.au3 (resp. ./src/pkg/Packager64.au3 for 64-bit)
when compiled with AutoIt3Wrapper produces an SFX package (./build/TypeLibInspector-sfx.exe)
that contains all required resources structured as described.
