using EasyConv;
using Microsoft.Win32;
using System.Text;

string vbscript = """
Set obj = CreateObject("wscript.shell")

Set args = WScript.Arguments
origin = args(0)
format = args(1)
If args.Count = 3 Then
    vfinsert = "-vf scale=" + args(2) + ":" + args(2) + " "
Else
    vfinsert = ""
End If

splited = Split(origin, ".")
ext = splited(UBound(splited))
namelen = Len(origin) - 1 - Len(ext)

If namelen < 1 Then
    nameonly = origin
Else
    nameonly = Left(origin, namelen)
End If

outname = Chr(34) + nameonly + "." + format + Chr(34)
cmdline = "ffmpeg -y -i " + Chr(34) + origin + Chr(34) + " " + vfinsert + outname
errcode = obj.run(cmdline, 0, True)

If errcode <> 0 Then
    MsgBox "无法转换"
End If
""";

Console.WriteLine("EasyConv 管理. 按下按键:");
Console.WriteLine("  安装(Install): A");
Console.WriteLine("  卸载(Uninstall): B");

bool next;

do
{
    next = false;
    var key = Console.ReadKey(true).Key;
    if (key == ConsoleKey.A)
    {
        if (Utils.FindFile("ffmpeg.exe") is not string ffmpegPath)
        {
            Console.WriteLine("程序无法从 PATH 中找到 ffmpeg.exe (This program cannot find ffmpeg.exe from PATH)");
            Environment.ExitCode = -1;
            break;
        }

        Console.WriteLine($"FFmpeg 路径 (FFmpeg path):");
        Console.WriteLine($"  {ffmpegPath}");

        string? ffmpegDir = Path.GetDirectoryName(ffmpegPath);
        Console.WriteLine($"FFmpeg 目录 (FFmpeg directory):");
        Console.WriteLine($"  {ffmpegDir}");

        if (ffmpegDir == null)
        {
            Console.WriteLine("出现了不可预知的错误, 无法获取 ffmpeg.exe 的父级目录");
            Environment.ExitCode = -2;
            break;
        }


        string converterScriptPath = Path.Combine(ffmpegDir, "converter.vbs");

        File.WriteAllText(converterScriptPath, vbscript, Encoding.Unicode);
        Console.WriteLine($"辅助脚本已导出 (Helper script extracted):");
        Console.WriteLine($"  {converterScriptPath}");

        if (Utils.TryRegister(converterScriptPath))
        {
            Console.WriteLine("注册表写入完成 (Registry wrote):");
            Console.WriteLine($"  {Utils.RegistryBase}");

            Console.WriteLine("安装成功 (Install successfully");
        }
        else
        {
            Console.WriteLine("安装失败 (Install failed)");
        }
    }
    else if (key == ConsoleKey.B)
    {
        if (Utils.FindFile("ffmpeg.exe") is not string ffmpegPath)
        {
            Console.WriteLine("程序无法从 PATH 中找到 ffmpeg.exe (This program cannot find ffmpeg.exe from PATH)");
            Environment.ExitCode = -1;
            break;
        }

        string? ffmpegDir = Path.GetDirectoryName(ffmpegPath);
        if (ffmpegDir == null)
        {
            Console.WriteLine("出现了不可预知的错误, 无法获取 ffmpeg.exe 的父级目录");
            Environment.ExitCode = -2;
            break;
        }

        string converterScriptPath = Path.Combine(ffmpegDir, "converter.vbs");

        if (File.Exists(converterScriptPath))
        {
            File.Delete(converterScriptPath);
            Console.WriteLine($"辅助脚本已删除 (Helper script deleted):");
            Console.WriteLine($"  {converterScriptPath}");
        }
        else
        {
            Console.WriteLine($"未找到辅助脚本程序 (Helper script connot be found):");
            Console.WriteLine($"  {converterScriptPath}");
        }

        if (Utils.TryUnregister())
        {
            Console.WriteLine("注册表移除完成 (Registry removed):");
            Console.WriteLine($"  {Utils.RegistryBase}");

            Console.WriteLine("卸载成功 (Uninstall successfully");
        }
        else
        {
            if (Utils.RegistryExisted)
            {
                Console.WriteLine(@"卸载失败, 请检查注册表 (Uninstall failed, please check registry table)");
                Console.WriteLine($"  {Utils.RegistryBase}");
            }
            else
            {
                Console.WriteLine(@"没有找到目标注册表项 (Cannot find target registry):");
                Console.WriteLine($"  {Utils.RegistryBase}");
            }
        }
    }
    else
    {
        Console.WriteLine("未知选项(Unknown choice)");
        next = true;
    }
}
while (next);

Console.WriteLine("Press any key to continue");
Console.ReadKey(true);
