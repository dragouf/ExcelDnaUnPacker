ExcelDnaUnPacker
================

A program to unpack .xll add-in pack with Excel-DNA project : https://exceldna.codeplex.com/


Usage
------------------

ExcelDnaUnPack is a command-line utility to un-pack ExcelDna add-ins into a single .dll file.

Usage: ExcelDnaUnPack.exe xllPath [/O outputPath] [/Y] 

  xllPath      The path to the .xll file for the ExcelDna add-in.
  /Y           If the output .dll exists, overwrite without prompting.
  /O outPath   Output path - default is <dnaPath>-unpacked.dll.

Example: ExcelDnaUnPack.exe MyAddins\FirstAddin.xll
		 The unpacked add-in file will be created as MyAddins\FirstAddin-unpacked.dll.

The Excel-Dna integration assembly (ExcelDna.Integration.dll) is searched for 
  1. in the same directory as the .xll file, and if not found there, 
  2. in the same directory as the ExcelDnaUnPack.exe file.