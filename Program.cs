internal static class Program
{
    /// <summary>
    /// Entry poin of the applicaton.
    /// Expects two arguments:
    /// 1) Input Excel file (.xlsx)
    /// 2) Output directory for generated PNG files
    /// </summary>
    /// <param name="args">
    /// Command-line arguments.
    /// </param>
    /// <returns>
    /// Process exit code:
    /// 0 = success,
    /// 1 = invalid arguments / input not found,
    /// 2 = runtime error.
    /// </returns>
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length < 2 )
        {
            Console.WriteLine("Usage: <input.xlsx> <output.pdf>");
            return 1;
        }

        string inputPath = Path.GetFullPath(args[0]);
        string pngDir = Path.GetFullPath(args[1]);

        Console.WriteLine("Input : " + inputPath);
        Console.WriteLine("Output: " + pngDir);

        if (!File.Exists(inputPath)) 
        {
            Console.WriteLine("Input file not found");
            return 1;
        }

        try
        {
            var converter = new ExcelToUhdConverter();
            converter.Convert(inputPath, pngDir, message => Console.WriteLine(message));

            Console.WriteLine("Conversion successful.");
            return 0;

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            return 2;   
        }
    }
}