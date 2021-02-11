using System;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace EmlTimeStamper
{
    class Program
    {
        /// <summary>
        /// Main entry point.
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        /// <returns></returns>
        static int Main(string[] args)
        {
            try
            {
                // Can optionally specify a directory name as a command line argument.
                var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (args.Length == 1)
                    path = args[0];
                TimeStampFiles(path);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.ToString());
                return 1;
            }
            return 0;
        }

        /// <summary>
        /// Change the timestamp of all files in a directory and all child directories.
        /// </summary>
        /// <param name="path">The directory path.</param>
        static void TimeStampFiles(string path)
        {
            // Change timestamp of all .eml files.
            foreach (var fileName in Directory.GetFiles(path, "*.eml"))
            {
                var emlContents = File.ReadAllText(fileName);
                var match = Regex.Match(emlContents, @"Date: \w+, (.+).+\n");
                if (!match.Success)
                    Console.WriteLine($"ERROR: Cannot find date in file: {fileName}");
                else
                {
                    var dateString = match.Groups[1].Value;
                    int posBracket = dateString.IndexOf('(');
                    if (posBracket != -1)
                        dateString = dateString.Remove(posBracket);
                    dateString = dateString.Replace(" UT", "");
                    var date = DateTime.Parse(dateString);

                    // Set timestamp.
                    File.SetCreationTime(fileName, date);
                    File.SetLastWriteTime(fileName, date);

                    // Add date to filename.
                    var newName = date.ToString("yyyy-MM-dd HH.mm.ss ") + Path.GetFileName(fileName);
                    var newFullFileName = Path.Combine(Path.GetDirectoryName(fileName), newName);
                    File.Move(fileName, newFullFileName);

                    Console.WriteLine($"Set timestamp of {fileName}");
                }
            }

            // Recurse through all child directories.
            foreach (var directoryName in Directory.GetDirectories(path))
                TimeStampFiles(directoryName);
        }
    }
}