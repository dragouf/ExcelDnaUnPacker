using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using SevenZip.Compression.LZMA;
using System.Reflection;
using System.IO;
using System.Collections.Generic;


[Flags]
enum LoadLibraryFlags : uint
{
    DONT_RESOLVE_DLL_REFERENCES = 0x00000001,
    LOAD_IGNORE_CODE_AUTHZ_LEVEL = 0x00000010,
    LOAD_LIBRARY_AS_DATAFILE = 0x00000002,
    LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = 0x00000040,
    LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x00000020,
    LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008
}


internal unsafe static class ResourceHelper
{
	// TODO: Learn about locales
	private const ushort localeNeutral		= 0;
	private const ushort localeEnglishUS	= 1033;
	private const ushort localeEnglishSA	= 7177;

	[DllImport("kernel32.dll")]
	private static extern IntPtr BeginUpdateResource(
		string pFileName,
		bool bDeleteExistingResources);

	[DllImport("kernel32.dll")]
	private static extern bool EndUpdateResource(
		IntPtr hUpdate,
		bool fDiscard);
	
	//, EntryPoint="UpdateResourceA", CharSet=CharSet.Ansi,
	[DllImport("kernel32.dll", SetLastError=true)]
	private static extern bool UpdateResource(
		IntPtr hUpdate,
		string lpType,
		string lpName,
		ushort wLanguage,
		IntPtr lpData,
		uint cbData);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr FindResource(
        IntPtr hModule,
        string lpName,
        string lpType);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadResource(
        IntPtr hModule,
        IntPtr hResInfo);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SizeofResource(
        IntPtr hModule,
        IntPtr hResInfo);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LockResource(
        IntPtr hResData);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("kernel32.dll")]
    public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, LoadLibraryFlags dwFlags);

    [DllImport("kernel32.DLL")]
	private static extern uint GetLastError();


    // Load the resource, trying also as compressed if no uncompressed version is found.
    // If the resource type ends with "_LZMA", we decompress from the LZMA format.
    internal static byte[] LoadResourceBytes(IntPtr hModule, string typeName, string resourceName)
    {
        Debug.Print("LoadResourceBytes for resource {0} of type {1}", resourceName, typeName);
        IntPtr hResInfo = FindResource(hModule, resourceName, typeName);
        if (hResInfo == IntPtr.Zero)
        {
            // We expect this null result value when the resource does not exists.

            if (!typeName.EndsWith("_LZMA"))
            {
                // Try the compressed name.
                typeName += "_LZMA";
                hResInfo = FindResource(hModule, resourceName, typeName);
            }
            if (hResInfo == IntPtr.Zero)
            {
                Debug.Print("Resource not found - resource {0} of type {1}", resourceName, typeName);
                // Return null to indicate that the resource was not found.
                return null;
            }
        }
        IntPtr hResData = LoadResource(hModule, hResInfo);
        if (hResData == IntPtr.Zero)
        {
            // Unexpected error - this should not happen
            Debug.Print("Unexpected errror loading resource {0} of type {1}", resourceName, typeName);
            throw new Win32Exception();
        }
        uint size = SizeofResource(hModule, hResInfo);
        IntPtr pResourceBytes = LockResource(hResData);
        byte[] resourceBytes = new byte[size];
        Marshal.Copy(pResourceBytes, resourceBytes, 0, (int)size);

        if (typeName.EndsWith("_LZMA"))
            return Decompress(resourceBytes);
        else
            return resourceBytes;
    }

    private static byte[] Decompress(byte[] inputBytes)
    {
        MemoryStream newInStream = new MemoryStream(inputBytes);
        Decoder decoder = new Decoder();
        newInStream.Seek(0, 0);
        MemoryStream newOutStream = new MemoryStream();
        byte[] properties2 = new byte[5];
        if (newInStream.Read(properties2, 0, 5) != 5)
            throw (new Exception("input .lzma is too short"));
        long outSize = 0;
        for (int i = 0; i < 8; i++)
        {
            int v = newInStream.ReadByte();
            if (v < 0)
                throw (new Exception("Can't Read 1"));
            outSize |= ((long)(byte)v) << (8 * i);
        }
        decoder.SetDecoderProperties(properties2);
        long compressedSize = newInStream.Length - newInStream.Position;
        decoder.Code(newInStream, newOutStream, compressedSize, outSize, null);
        byte[] b = newOutStream.ToArray();
        return b;
    }
}