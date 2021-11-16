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
            var isSucces = false;

            _ = Log.AppendLine($"We have {OriginalFiles.Count()} to convert.");
            try
            {
                _ = Parallel.ForEach(OriginalFiles, (originalFile) =>
              {
                  // Just to be save: Throw error when user assumes an XLS-file is OK.
                  if (originalFile.EndsWith("XLS", StringComparison.CurrentCultureIgnoreCase))
                      throw new Exception("Only XLSM files are supported!");

                  var itemLog = new StringBuilder(originalFile);

                  // The original excel-file with values typed in by the user.
                  using (var originalFilePackage = new ExcelPackage(new FileInfo(originalFile)))
                  using (var destinationFilePackage = new ExcelPackage(new FileInfo(GetTempTemplateFilePath(TemplateFile, originalFile))))


                  {
                      var originWorkbook = originalFilePackage.Workbook;
                      var destinationWorkbook = destinationFilePackage.Workbook;
                      
                      _ = itemLog.AppendLine($"  Converting --> {originalFilePackage.File.Name}");
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

                      destinationFilePackage.Save();
                      _ = itemLog.AppendLine($"  Saved --> {destinationFilePackage.File.FullName}");
                  }

                  lock (Log)
                  {
                      _ = Log.Append(itemLog);
                  }
              });
                isSucces = true;
            }
            catch (Exception exception)
            {
                _ = Log.AppendLine($"Failed to read file --> {exception.Message} --> {exception.InnerException?.Message}");

            }

            var logString = Log.ToString();
            if (isSucces)
                Console.WriteLine($"Success: {OriginalFiles.Count()} have been processed successfully: Log-file => {logfile}");
            else
                Console.WriteLine($"Process failed: Log-file => {logfile}");

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
            Console.WriteLine("Some parsing errors occurred, please try again with valid parameters.");
        }

        private static void Main(string[] args)
        {
            _ = CommandLine.Parser.Default.ParseArguments<Options>(args)
         .WithParsed(RunOptions)
         .WithNotParsed(HandleParseError);
        }

        private static string RangeCopyValuesBetweenSheets(ExcelRange originalValues, ExcelWorksheet newTemplateSheet)
        {
            var log = new StringBuilder($"--> Sheet[{newTemplateSheet.Name}] Range [{originalValues.Address}] <--\r\n");
            int skippedCelles = 0;
            foreach (var sourceCell in originalValues)
            {
                var destinationCell = newTemplateSheet.Cells[sourceCell.Address];

                // fix: Hack -> safety that can be removed later on, a lot of people seem to place
                // this value in white last month days cells.
                if (sourceCell.Text.Equals("555") && destinationCell.Formula != "")
                {
                    _ = log.AppendLine($"   |--> Skipped value 555: SourceCell address[{sourceCell.Address}]; formula[{sourceCell.Formula}]; value[{sourceCell.Value}]; ");
                    continue;
                }
                // Do not update values are equal
                if (destinationCell.Text.Equals(sourceCell.Text, StringComparison.CurrentCultureIgnoreCase))
                {
                    skippedCelles++;
                    continue;
                }

                // This mitigates polluting the destination excel with empty string values from source file.
                if (sourceCell.Formula == "" && string.IsNullOrWhiteSpace(sourceCell.Text))
                {
                    destinationCell.Value = "";
                    _ = log.AppendLine($"   |--> Set value: SourceCell address[{sourceCell.Address}]; to empty string; ");
                    continue;
                }

                // Soft set of value, formula is not broken.
                if (destinationCell.Formula == "" && !string.IsNullOrWhiteSpace(sourceCell.Text))
                {
                    destinationCell.Value = sourceCell.Value;
                    _ = log.AppendLine($"   |--> Copied value: SourceCell address[{sourceCell.Address}]; value[{sourceCell.Value}]; SOFT SET destination: no formula.");
                    continue;
                }

                // Hard set destination, break formula.
                if (string.IsNullOrWhiteSpace(sourceCell.Formula) && !string.IsNullOrWhiteSpace(sourceCell.Text))
                {
                    destinationCell.Value = sourceCell.Value;
                    _ = log.AppendLine($"   |--> Set value: SourceCell address[{sourceCell.Address}]; value[{sourceCell.Value}]; HARD SET destination: formula broken!");
                }

                // No valid reason to copy values, so skip.
                skippedCelles++;

            }
            _ = log.AppendLine($"   |--> Total number of skippedCells = {skippedCelles};");
            return log.ToString();
        }

        private static void RunOptions(Options opts)
        {
            var SheetsMonths = Properties.Settings.Default.SheetMonths;

            var RangesToCopyValuesFrom = Properties.Settings.Default.Ranges;

            //handle options
            try
            {
                if (!File.Exists(opts.TemplateFilePath))
                    throw new FileNotFoundException($"File not found:{opts.TemplateFilePath}");
                foreach (var f in opts.InputFiles)
                    if (!File.Exists(f))
                        throw new FileNotFoundException($"File not found:{f}");

                Execute(
                    OriginalFiles: opts.InputFiles,
                    SheetsMonths: SheetsMonths.Cast<string>(),
                    RangesToCopyValuesFrom: RangesToCopyValuesFrom.Cast<string>(),
                    TemplateFile: opts.TemplateFilePath);
            }
            catch (Exception ex)
            {
                // throw;
                Console.Write($"Failed: {ex.GetType()} {ex.Message}");
            }
        }

        #endregion

        #region Private Classes

        private class Options
        {
            #region Public Properties

            [Option('o', "original", Required = true, HelpText = "Input files to be processed.")]
            public IEnumerable<string> InputFiles { get; set; }

            [Option('t', "template", Required = true, HelpText = "Template-file to be used.")]
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