## Altium Library Loader v.2.3 (unofficial)
Modified Altium Library Loader scripts from Samacsys to be used with Designer 26+. Might also work with previous versions. Unpacked from the installer 2.2 directly with Innoextract 1.9 (https://github.com/dscharrer/innoextract/releases/tag/1.9), re-packed with Inno Setup 6.7.1 by https://www.innosetup.com.

One-click installation, as original.
All rights belong to https://supplyframe.com/ and SamacSys.

# Compilation
1. Install InnoSetup:
```bash
winget install JRSoftware.InnoSetup 
```
2. Clone this repo:
```bash
git clone https://github.com/lethargicpillow/Altium-Library-Loader.git
```
3.1 Compile via GUI: 
Go to the cloned folder, double-click .iss and compile in GUI (button on ribbon or Ctrl+F9), or:

3.2 Compile via cmd

cd to the cloned folder and run  
```bash
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" AltiumLL_23.iss
```
Compiled .exe will be in the root, whichever approach you choose.