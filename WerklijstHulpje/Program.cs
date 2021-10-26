using CommandLine;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WerklijstHulpje
{
    internal class Program
    {
        #region Private Methods

        private static void Execute(IEnumerable<string> OriginalFiles, IEnumerable<string> SheetsMonths, IEnumerable<string> RangesToCopyValuesFrom, string TemplateFile)
        {
            var Log = new StringBuilder($"{DateTime.Now.ToLongDateString()} | {DateTime.Now.ToLongTimeString()}" + Environment.NewLine);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var workDir = Path.GetDirectoryName(TemplateFile);
            var logfile = workDir + "\\werklijsthulpje.log.txt";

            _ = Log.AppendLine($"We have {OriginalFiles.Count()} to convert.");

            _ = Parallel.ForEach(OriginalFiles, (originalFile) =>
              {
                  StringBuilder itemLog = new StringBuilder(originalFile);
                  using (var destinationPackage = new ExcelPackage(new FileInfo(GetTempTemplateFilePath(TemplateFile, originalFile))))
                  using (var originalFileilePackage = new ExcelPackage(new FileInfo(originalFile)))

                  {
                      var originWorkbook = originalFileilePackage.Workbook;
                      var destinationWorkbook = destinationPackage.Workbook;
                      _ = itemLog.AppendLine($"  Converting --> {originalFileilePackage.File.Name}");
                      foreach (var month in SheetsMonths)
                      {
                          var originMonthSheet = originWorkbook.Worksheets[month];
                          var desinationMonthSheet = destinationWorkbook.Worksheets[month];
                          _ = itemLog.AppendLine($"   Month sheet: {month}");
                          foreach (var range in RangesToCopyValuesFrom)
                          {
                              var cellsToCopy = originMonthSheet.Cells[range];
                              var rangeLog = RangeCopyValuesBetweenSheets(cellsToCopy, desinationMonthSheet);
                              if (!string.IsNullOrWhiteSpace(rangeLog))
                                  _ = itemLog.Append(rangeLog);
                          }
                      }

                      destinationPackage.Save();
                      _ = itemLog.AppendLine($"  Saved --> {destinationPackage.File.FullName}");
                  }
                  lock (Log)
                  {
                      _ = Log.Append(itemLog);
                  }
              });
            Console.WriteLine($"Succes: {OriginalFiles.Count()} have been processed succesfully: Logfile => {logfile}");

            System.IO.File.WriteAllText(logfile, Log.ToString());
        }

        /// <summary>
        /// Copies the template file to the new destination.
        /// </summary>
        /// <param name="filename">    Template file path </param>
        /// <param name="destination"> Destination file path </param>
        /// <returns>
        /// Destination file path with its filename appended with the .new.xlsm
        /// </returns>
        private static string GetTempTemplateFilePath(string filename, string destination)
        {
            // todo: make this nice and clean!

            var t = destination.Replace(".xlsm", ".new.xlsm");
            File.Copy(filename, t, true);

            return t;
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
            //handle errors
            Console.WriteLine("Some parsing errors occuer, please try again with valid paramaters.");
        }

        private static void Main(string[] args)
        {
            _ = CommandLine.Parser.Default.ParseArguments<Options>(args)
         .WithParsed(RunOptions)
         .WithNotParsed(HandleParseError);
        }

        private static string RangeCopyValuesBetweenSheets(ExcelRange originalValues, ExcelWorksheet newTemplateSheet)
        {
            var log = new StringBuilder();
            int skippedLines = 0;
            foreach (var sourceCell in originalValues)
            {
                var destinationCell = newTemplateSheet.Cells[sourceCell.Address];

                // Hacky safety that can be removed later on, a lot of people seem to place
                // this value in white last mont days cells.
                if (sourceCell.Text.Equals("555") && destinationCell.Formula != "")
                {
                    _ = log.AppendLine($"Skipped value 555: Range[{sourceCell.Address}]; Formula[{sourceCell.Formula}]; Value[{sourceCell.Value}]; ");
                    continue;
                }
                // Do not update what is already the same value
                if (destinationCell.Text.Equals(sourceCell.Text))
                {
                    skippedLines++;
                    continue;
                }

                // If we get here we probably want to copy value's
                destinationCell.Value = sourceCell.Value;
                _ = log.AppendLine($"Copied value: Range[{sourceCell.Address}]; Value[{sourceCell.Value}]; ");

            }
            _ = log.AppendLine($"SkippedCells = {skippedLines};");
            return log.ToString().Trim();
        }

        private static void RunOptions(Options opts)
        {
            string[] SheetsMonths = new string[] {
            "januari", "februari", "maart", "april", "mei", "juni",
            "juli","augustus","september", "oktober", "november", "december"};

            string[] RangesToCopyValuesFrom = "U1;E3;C8:I38;N8:R38;T8:T37;V41;V44;E49;E50;E51;E52;G53;G54;D42;D43;D44;D45;D46".Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            //handle options
            try
            {
                Execute(opts.InputFiles, SheetsMonths, RangesToCopyValuesFrom, opts.TemplateFilePath);
            }
            catch (Exception ex)
            {
                // throw;
                Console.Write($"Failed: {ex.Message}");
            }
        }

        #endregion

        #region Private Classes

        private class Options
        {
            #region Public Properties

            [Option('o', "original", Required = true, HelpText = "Input files to be processed.")]
            public IEnumerable<string> InputFiles { get; set; }

            [Option('t', "template", Required = true, HelpText = "Templatefile to be used.")]
            public string TemplateFilePath { get; set; }

            // Omitting long name, defaults to name of property, ie "--verbose"
            [Option(
              Default = false,
              HelpText = "Prints all messages to standard output.")]
            public bool Verbose { get; set; }

            #endregion
        }

        #endregion
    }
}