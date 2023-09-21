namespace IpEditor;

internal static class Logger
{

    public static void Info(string message)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("[info] ");
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine($"{message}");
    }

    public static void Error(string message)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Write("[error] ");
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine($"{message}");
    }

    public static void Warning(string message)
    {
        Console.ForegroundColor = ConsoleColor.Blue;
        Console.Write("[warning] ");
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine($"{message}");
    }
}
