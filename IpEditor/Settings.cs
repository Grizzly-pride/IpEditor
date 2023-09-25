using IpEditor.Entity.ExcelData;
namespace IpEditor;

internal class Settings
{
    [NonSerialized]
    private const string _logo = @"
  _____ _____    ______    _ _ _              
 |_   _|  __ \  |  ____|  | (_) |             
   | | | |__) | | |__   __| |_| |_ ___  _ __  
   | | |  ___/  |  __| / _` | | __/ _ \| '__| 
  _| |_| |      | |___| (_| | | || (_) | |    
 |_____|_|      |______\__,_|_|\__\___/|_|   v 0.1
                                                                                           
 Tool for edit eNodeB transport in Huawei bulk configuration file.

";
    public SourceFile SourceFile { get; init; }
    public TargetFile TargetFile { get; init; }

    public static void PrintLogo(ConsoleColor color)
    {
        Console.CursorVisible = false;  
        Console.ForegroundColor = color;
        Console.WriteLine(_logo);
        Console.ForegroundColor = ConsoleColor.White;
    }

    public static void TaskCompletedMessage(ConsoleColor color)
    {
        Console.ForegroundColor = color;
        Console.WriteLine();
        Console.WriteLine(" Таск has been completed.");
        Console.ForegroundColor = ConsoleColor.White;
    }
}
