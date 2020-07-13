using CommandLine;
using System.IO;

namespace BypassAv
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
              .WithParsed(RunTweaker);
        }

        private static void RunTweaker(Options options)
        {
            string inputPath = options.InputFilePath;
            string outputPath = options.Overwrite ? inputPath : options.OutputFilePath;

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                outputPath = GetOutputPathWithSuffix(options.InputFilePath);
            }

            OfficeXTweaker tweaker = new OfficeXTweaker(inputPath);
            tweaker.RemoveProperties(outputPath);
        }

        /// <summary>
        /// Based on inputPath, creates an output path.
        /// For example, if inputPath is "c:\temp\a.docx", output path will be "c:\temp\a.out.docx"
        /// </summary>
        /// <param name="inputPath">Path to input file</param>
        /// <returns>Output path based on input path.</returns>
        private static string GetOutputPathWithSuffix(string inputPath)
        {
            string folderPath = Path.GetDirectoryName(inputPath);
            string nameWitoutExtension = Path.GetFileNameWithoutExtension(inputPath);
            string extension = Path.GetExtension(inputPath);
            string outputFileName = nameWitoutExtension + Consts.DEFAULT_OUTPUT_SUFFIX + extension;
            return Path.Combine(folderPath, outputFileName);
        }

        public class Options
        {
            [Option('i', "input", Required = true, HelpText = "Path to OfficeX file to tweak.")]
            public string InputFilePath { get; set; }

            [Option('o', "output", Required = false, HelpText = "Output path of tweaked OfficeX file. If not exists, will be <input path>.out.<ext>")]
            public string OutputFilePath { get; set; }

            [Option('w', "overwrite", Required = false, Default = false, HelpText = "Ignore output file path and overwrite input path.")]
            public bool Overwrite { get; set; }
        }
    }
}