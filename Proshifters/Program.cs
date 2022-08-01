// See https://aka.ms/new-console-template for more information

using Traherom.Proshifters;

class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 1) {
            Console.Error.WriteLine($"Usage: proshifter <schedule-path>"); 
            return;
        }
        
        // Put the resulting calculation in the same directory as the original file
        var originalPath = args[0];
        var originalDir = Path.GetDirectoryName(originalPath) ?? ".";
        var resultPath = Path.Combine(originalDir, "result.xlsx");
        // Console.WriteLine($"Writing results to {resultPath}");

        var shifter = new Proshifter();
        shifter.Calculate(originalPath, resultPath);
    }
}
