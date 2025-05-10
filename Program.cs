using System;
using System.IO;
using System.Reflection;
using System.Text;
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

                // Look for .mbox files and convert to .eml files in sub directories named by years.
                bool mboxFound = ConvertMBoxToEml(path);

                // If no mbox was found then look for .eml files and change their name and date time stamp.
                if (!mboxFound)
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
                var date = GetDateFromMessage(emlContents, fileName);
                if (date != DateTime.MinValue)
                {
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

        /// <summary>
        /// Get the date from an email message.
        /// </summary>
        /// <param name="emlContents">The contents of the email message.</param>
        /// <param name="fileName">The file name the email came from.</param>
        /// <returns>The found date or DateTime.MinValue if not found.</returns>
        private static DateTime GetDateFromMessage(string emlContents, string fileName)
        {
            var stringToFind = Environment.NewLine + "Date: ";
            int posDate = emlContents.IndexOf(stringToFind);
            if (posDate == -1)
            {
                Console.WriteLine($"ERROR: Cannot find date in file: {fileName}");
                return DateTime.MinValue;
            }   
            else
            {
                posDate += stringToFind.Length;
                int posEndDate = emlContents.IndexOf(Environment.NewLine, posDate);
                if (posEndDate == -1)
                {
                    Console.WriteLine($"ERROR: Cannot find date in file: {fileName}");
                    return DateTime.MinValue;
                }

                
                var dateString = emlContents.Substring(posDate, posEndDate - posDate);
                int posBracket = dateString.IndexOf('(');
                if (posBracket != -1)
                    dateString = dateString.Remove(posBracket);
                dateString = dateString.Replace(" UT", "");
                return DateTime.Parse(dateString);
            }
        }

        /// <summary>
        /// Look for .mbox files and convert to .eml files.
        /// </summary>
        /// <param name="path">The path to search.</param>
        /// <returns>true if mbox was converted.</returns>
        private static bool ConvertMBoxToEml(string path)
        {
            bool mboxFound = false;
            foreach (var fileName in Directory.GetFiles(path, "*.mbox"))
            {
                mboxFound = true;
                StringBuilder message = new StringBuilder();
                string subject = null;
                using var reader = new StreamReader(fileName);
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line.StartsWith("From "))
                    {
                        SaveMessageToFile(message.ToString(), subject, path);
                        message.Clear();
                    }
                    message.AppendLine(line);

                    if (line.StartsWith("Subject: "))
                    {
                        subject = line.Replace("Subject: ", "");
                        if (subject.ToLower().Contains("?utf-"))
                            subject = null;
                    }
                }

                // At end of file - save the message we were building.
                SaveMessageToFile(message.ToString(), subject, Path.GetDirectoryName(path));
            }
            return mboxFound;
        }

        private static void SaveMessageToFile(string emlContents, string subject, string parentDirectoryName)
        {
            if (emlContents.Length > 0)
            {
                var date = GetDateFromMessage(emlContents, null);

                string fileName = $"{date.ToString("yyyy-MM-dd HH.mm.ss")} {subject}.eml";

                // Sanititise file name
                var pattern = ":|\\[|\\]|\\\\|\"|'|/|<|>|\\\\|\\?|=|!|,|\\*|\t|\\|";
                fileName = Regex.Replace(fileName, pattern, "");
                fileName = Regex.Replace(fileName, @"<\w+>", "");

                string directoryName = Path.Combine(parentDirectoryName, date.Year.ToString());
                string path = Path.Combine(directoryName, fileName);
                Directory.CreateDirectory(directoryName);
                if (File.Exists(path))
                    Console.WriteLine($"Ignoring {Path.GetFileName(path)}. File already exists.");
                else
                {
                    Console.WriteLine($"Writing {path}");
                    File.WriteAllText(path, emlContents);
                }
            }
        }
    }
}