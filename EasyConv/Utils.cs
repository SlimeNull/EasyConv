using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyConv
{
    internal static class Utils
    {
        public static readonly string RegistryBase = @"HKEY_CLASSES_ROOT\*\shell\converter_ffmpeg";
        public static string? FindFile(string filename)
        {
            string[]? paths = Environment.GetEnvironmentVariable("PATH")?.Split(";");
            if (paths == null)
                return null;

            foreach (string path in paths)
            {
                string[] files = Directory.GetFiles(path);
                foreach (string file in files)
                {
                    if (Path.GetFileName(file).Equals(filename, StringComparison.OrdinalIgnoreCase))
                    {
                        return file;
                    }
                }
            }

            return null;
        }

        internal static bool RegistryExisted => Registry.ClassesRoot.OpenSubKey(@"*\shell\converter_ffmpeg") != null;

        internal static bool TryRegister(string converterPath)
        {
            void CreateShellCommand(RegistryKey shell, string format, int[]? sizes = null)
            {
                if (sizes != null)
                {
                    RegistryKey subkey = shell.CreateSubKey($"to_{format}");
                    subkey.SetValue("MUIVerb", $"To {format}");
                    subkey.SetValue("Subcommands", string.Empty);

                    RegistryKey subshell = subkey.CreateSubKey("shell");
                    foreach (int size in sizes)
                    {
                        string cmdline = $"wscript.exe \"{converterPath}\" \"%1\" {format} {size}";

                        RegistryKey subkeysub = subshell.CreateSubKey($"{size}x{size}");
                        RegistryKey cmdkeysub = subkeysub.CreateSubKey("command");
                        subkeysub.SetValue(null, $"{size}x");
                        cmdkeysub.SetValue(null, cmdline);
                    }
                }
                else
                {
                    string cmdline = $"wscript.exe \"{converterPath}\" \"%1\" {format}";

                    RegistryKey subkey = shell.CreateSubKey($"to_{format}");
                    RegistryKey cmdkey = subkey.CreateSubKey("command");
                    subkey.SetValue(null, $"To {format}");
                    cmdkey.SetValue(null, cmdline);
                }
            }

            try
            {
                RegistryKey converter = Registry.ClassesRoot.CreateSubKey(@"*\shell\converter_ffmpeg");
                converter.SetValue("MUIVerb", "MediaConvert");
                converter.SetValue("Subcommands", "");
                //converter.SetValue("Icon", "imageres.dll,229");

                RegistryKey shell = converter.CreateSubKey("shell");

                string[] formats = new string[]
                {
                    "jpg", "png", "bmp", "gif", "webp", "mp3", "mp4"
                };

                foreach (string format in formats)
                {
                    CreateShellCommand(shell, format);
                }

                CreateShellCommand(shell, "ico", new int[] { 8, 16, 32, 64, 128, 256 });

                return true;
            }
            catch
            {
                return false;
            }
        }

        internal static bool TryUnregister()
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKeyTree(@"*\shell\converter_ffmpeg");
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
