using System;

namespace ClosedXML_Sandbox
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Running {0}", nameof(PerformanceRunner.OpenTestFile));
            PerformanceRunner.TimeAction(PerformanceRunner.OpenTestFile);
            Console.WriteLine();

            // Disable this block by default - I don't use it often
#if false

            Console.WriteLine("Running {0}", nameof(PerformanceRunner.RunInsertTable));
            PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTable);
            Console.WriteLine();
#endif

            Console.WriteLine("Running {0}", nameof(PerformanceRunner.RunInsertTableWithStyles));
            PerformanceRunner.TimeAction(PerformanceRunner.RunInsertTableWithStyles);
            Console.WriteLine();

            Console.WriteLine("Press any key to continue");
            Console.ReadKey();
        }
    }
}
