using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Reflection;

using System.IO;
using System.Diagnostics;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace ExcelDnaPack
{
	class PackProgram
	{
		static string usageInfo =
@"ExcelDnaUnPack Usage
------------------
ExcelDnaUnPack is a command-line utility to un-pack ExcelDna add-ins into a single .dll file.

Usage: ExcelDnaUnPack.exe xllPath [/O outputPath] [/Y] 

  xllPath      The path to the .xll file for the ExcelDna add-in.
  /Y           If the output .dll exists, overwrite without prompting.
  /O outPath   Output path - default is <dnaPath>-unpacked.dll.

Example: ExcelDnaUnPack.exe MyAddins\FirstAddin.sll
		 The unpacked add-in file will be created as MyAddins\FirstAddin-unpacked.dll.

The template add-in host file (the copy of ExcelDna.dll renamed to FirstAddin.xll) is 
searched for in the same directory as FirstAddin.dna.

The Excel-Dna integration assembly (ExcelDna.Integration.dll) is searched for 
  1. in the same directory as the .xll file, and if not found there, 
  2. in the same directory as the ExcelDnaUnPack.exe file.
";
		
		static void Main(string[] args)
		{
            // Force jit-load of ExcelDna.Integration assembly
            int unused = XlCall.xlAbort;

            if (args.Length < 1)
            {
                Console.Write("No .xll file specified.\r\n\r\n" + usageInfo);
                return;
            }

            string xllPath = args[0];
            string xllDirectory = Path.GetDirectoryName(xllPath);           
            string xllFilePrefix = Path.GetFileNameWithoutExtension(xllPath);
            string dllOutputPath = Path.Combine(xllDirectory, xllFilePrefix + "-unpacked.dll");

            bool overwrite = false;

            if (!File.Exists(xllPath))
            {
                Console.Write("Add-in .xll file " + xllPath + " not found.\r\n\r\n" + usageInfo);
                return;
            }

            // TODO: Replace with an args-parsing routine.
            if (args.Length > 1)
            {
                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].ToUpper() == "/O")
                    {
                        if (i >= args.Length - 1)
                        {
                            // Too few args.
                            Console.Write("Invalid command-line arguments.\r\n\r\n" + usageInfo);
                            return;
                        }
                        dllOutputPath = args[i + 1];
                    }
                    else if (args[i].ToUpper() == "/Y")
                    {
                        overwrite = true;
                    }
                }
            }

            if (File.Exists(dllOutputPath))
            {
                if (overwrite == false)
                {
                    Console.Write("Output .dll file " + dllOutputPath + " already exists. Overwrite? [Y/N] ");
                    string response = Console.ReadLine();
                    if (response.ToUpper() != "Y")
                    {
                        Console.WriteLine("\r\nNot overwriting existing file.\r\nExiting ExcelDnaUnPack.");
                        return;
                    }
                }

                try
                {
                    File.Delete(dllOutputPath);
                }
                catch
                {
                    Console.Write("Existing output .dll file " + dllOutputPath + "could not be deleted. (Perhaps loaded somewhere ?)\r\n\r\nExiting ExcelDnaUnPack.");
                    return;
                }
            }

            // Look inside an XLL with PEView. Each dll is packed in SECTION .rsrc as ASSEMBLY_LZMA <name> where name is the
            // DLL name without .dll suffix
            string lzmaName = dllOutputPath.Substring( 0, dllOutputPath.Length - 4);
            Console.WriteLine( "Unpacking " + lzmaName );
            var hModule = ResourceHelper.LoadLibraryEx(xllPath, IntPtr.Zero, LoadLibraryFlags.LOAD_LIBRARY_AS_DATAFILE | LoadLibraryFlags.LOAD_LIBRARY_AS_IMAGE_RESOURCE);
            var content = ResourceHelper.LoadResourceBytes( hModule, "ASSEMBLY_LZMA", lzmaName);

            if (null != content)
            {
                using (BinaryWriter binWriter = new BinaryWriter(File.Open(dllOutputPath, FileMode.Create)))
                {
                    binWriter.Write(content);
                }
            }
		}
	}

}
