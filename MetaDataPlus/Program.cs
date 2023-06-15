using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace MetadataPlus
{
    class Program
    {
        //Works on xlsx, xlsm, xltx, xltm, docx, docm, dotm, dotx, ppt, pptx, potm, potx
        //Created by Chris Nevin at NCCGroup
        public static List<Document> documents = new List<Document>();
        public static string workingFolder;
        public static bool askedHelp = false;
        public static bool grep = true;
        public static string additionalGrep;
        public static bool everyFileInFolder = false;
        public static bool extractMedia = false;
        public static bool extractEmbedded = false;
        public static int embedDocNumber = 1;
        public static int mediaDocNumber = 1;

        static void Main(string[] args)
        {
            //Deal with args
            bool setInput = false;
            foreach (string arg in args)
            {
                if (arg.Contains("-inputFolder") || arg.Contains("-i"))
                {
                    workingFolder = arg.Replace("-inputFolder=", "").Replace("-i=", "");
                    setInput = true;
                }
                if (arg.Contains("-Search") || arg.Contains("-s"))
                {
                    additionalGrep = arg.Replace("-Search=", "").Replace("-s=", "");
                }
                if (arg.Contains("-help") || arg.Contains("-h"))
                {
                    askedHelp = true;
                }
                if (arg.Contains("-Media") || arg.Contains("-m"))
                {
                    extractMedia = true;
                }
                if (arg.Contains("-Embed") || arg.Contains("-e"))
                {
                    extractEmbedded = true;
                }
                if (arg.Contains("-All") || arg.Contains("-a"))
                {
                    everyFileInFolder = true;
                }
            }
            if (!setInput)
            {
                Console.WriteLine("No options specified - using current working directory...");
                workingFolder = Directory.GetCurrentDirectory() + "\\";
            }
            if (askedHelp)
            {
                Console.WriteLine("MetadataPlus v1.0");
                Console.WriteLine("    Please specify the folder containing the documents to be analysed, or the current working directory will be used...");
                Console.WriteLine("    MetadataPlus can currently analyse XLSX/XLSM (Excel), XLTX/XLTM (Excel Template), DOCX/DOCM (Word), DOTX/DOTM (Word Template), PPT/PPTX (PowerPoint), POTX/POTM (PowerPoint Templates)");
                Console.WriteLine(Environment.NewLine);
                Console.WriteLine("Specify input folder: -inputFolder= / -i=");
                Console.WriteLine("   Example: -i=\"C:\\Temp");
                Console.WriteLine("Do not include grepping of internal files for sensitive strings: -NoGrep / -n");
                Console.WriteLine("   By default MetadataPlus searches specific places for usernames and sensitive data, and also searches all of each individual internal file. If this is returning a large amount of noise, you can use this to restrict searches to known locations.");
                Console.WriteLine("Supply additional grep string (this will work even if NoGrep is specified above): -Search= / -s=");
                Console.WriteLine("   Example: -s=\"apikey\"");
                Console.WriteLine("Try every file in folder: -All / -a");
                Console.WriteLine("   By default MetadataPlus will only try files known to work, including other files may lead to lots of corrupted file errors.");
                Console.WriteLine("Extract media for manual exif checking: -Media / -m");
                Console.WriteLine("   This will extract images to a Media folder for manual exif review. This is not on by default.");
                Console.WriteLine("Extract embedded documents: -Embed / -e");
                Console.WriteLine("   Extract embedded documents and objects to a Embedded folder for additional manual review. This is not on by default.");
                Console.WriteLine("Help: -help / -h");
            }
            else
            {
                if (extractEmbedded)
                {
                    if (Directory.Exists(workingFolder + "Embed\\"))
                    {
                        Directory.Delete(workingFolder + "Embed\\", true);
                        Directory.CreateDirectory(workingFolder + "Embed\\");
                    }
                    else
                    {
                        Directory.CreateDirectory(workingFolder + "Embed\\");
                    }
                }
                if (extractMedia)
                {
                    if (Directory.Exists(workingFolder + "Media\\"))
                    {
                        Directory.Delete(workingFolder + "Media\\", true);
                        Directory.CreateDirectory(workingFolder + "Media\\");
                    }
                    else
                    {
                        Directory.CreateDirectory(workingFolder + "Media\\");
                    }
                }

                //Process files
                if (everyFileInFolder)
                {
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.*", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                }
                else
                {
                    //Doesn't seem to be an inbuilt way to give multiple file types...
                    //Analyse each xlsx in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.xlsx", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each xlsm in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.xlsm", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each xltx in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.xltx", SearchOption.AllDirectories))
                    {
                        RunTheJewels(file);
                    }
                    //Analyse each xltm in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.xltm", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each docx in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.docx", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each docm in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.docm", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each dotm in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.dotm", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each dotx in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.dotx", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each ppt in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.ppt", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each pptx in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.pptx", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each potm in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.potm", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                    //Analyse each potx in folder
                    foreach (string file in Directory.EnumerateFiles(workingFolder + "\\", "*.potx", SearchOption.AllDirectories))
                    {
                        Console.WriteLine("Processing file: " + Path.GetFileName(file));
                        RunTheJewels(file);
                    }
                }
                //Final cleanup
                try
                {
                    //Delete temp folder if exists
                    if (Directory.Exists(workingFolder + "\\Temp\\"))
                    {
                        Directory.Delete(workingFolder + "\\Temp\\", true);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                Console.WriteLine();

                //Print Results
                Console.WriteLine("All Results:");
                //Get all docs that have metadata in
                var allDocsWithMetadata = documents.Where(p => p.MetaDataFound == true);
                if (allDocsWithMetadata.Count() > 0)
                {
                    //Get all hidden sheets
                    var allDocsWithSheets = documents.Where(p => p.HiddenSheet.Count() > 0);
                    if (allDocsWithSheets.Count() > 0)
                    {
                        Console.WriteLine("Hidden Sheets:");
                        foreach (var doc in allDocsWithSheets)
                        {
                            foreach (HiddenSheet sheet in doc.HiddenSheet)
                            {
                                Console.WriteLine("  " + sheet.HiddenSheetName + " in file: " + Path.GetFileName(sheet.HiddenSheetLocation));
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all external links
                    var allDocsWithLinks = documents.Where(p => p.ExternalLinks.Count() > 0);
                    if (allDocsWithLinks.Count() > 0)
                    {
                        Console.WriteLine("External Links:");
                        foreach (var doc in allDocsWithLinks)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (string link in doc.ExternalLinks)
                            {
                                Console.WriteLine("  " + link);
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all password docs
                    var allDocsWithPasswords = documents.Where(p => p.ContainsPassword == true);
                    if (allDocsWithPasswords.Count() > 0)
                    {
                        Console.WriteLine("Docs containing the word \"password\":");
                        foreach (var doc in allDocsWithPasswords)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, "") + " in inner files:");
                            foreach (var innerfile in doc.PasswordContainingInnerFileNames)
                            {
                                Console.WriteLine("  " + Path.GetFileName(innerfile));
                            }
                            Console.WriteLine("   in strings:");
                            List<string> alreadyPassword = new List<string>();
                            foreach (string pass in doc.PasswordContainingString)
                            {
                                if (alreadyPassword.Contains(pass))
                                {

                                }
                                else
                                {
                                    Console.WriteLine("    " + pass);
                                    alreadyPassword.Add(pass);
                                }
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all results for user provided grep
                    var allDocsWithUserGrep = documents.Where(p => p.UserGrep.Count() > 0);
                    if (allDocsWithUserGrep.Count() > 0)
                    {
                        Console.WriteLine("Docs containing user provided string:");
                        foreach (var doc in allDocsWithUserGrep)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (GrepObject grepped in doc.UserGrep)
                            {
                                Console.WriteLine("  Inner file: " + Path.GetFileName(grepped.InnerFile));
                                Console.WriteLine("  String: " + grepped.GrepString);
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all filepaths
                    var allDocsWithFilepaths = documents.Where(p => p.FilePaths.Count() > 0);
                    if (allDocsWithFilepaths.Count() > 0)
                    {
                        Console.WriteLine("File paths:");
                        foreach (var doc in allDocsWithFilepaths)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (string filepath in doc.FilePaths)
                            {
                                Console.WriteLine("  " + filepath);
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all printers
                    var allDocsWithPrinters = documents.Where(p => p.Printers.Count() > 0);
                    if (allDocsWithPrinters.Count() > 0)
                    {
                        Console.WriteLine("Printers:");
                        foreach (var doc in allDocsWithPrinters)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (string printer in doc.Printers)
                            {
                                Console.WriteLine("  " + printer);
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get hostnames
                    var allDocsWithHostnames = documents.Where(p => p.Hostnames.Count() > 0);
                    if (allDocsWithHostnames.Count() > 0)
                    {
                        Console.WriteLine("Hostnames:");
                        foreach (var doc in allDocsWithHostnames)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (string host in doc.Hostnames)
                            {
                                Console.WriteLine("  " + host);
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Emails
                    var allDocsWithEmails = documents.Where(p => p.Emails.Count() > 0);
                    if (allDocsWithEmails.Count() > 0)
                    {
                        Console.WriteLine("Emails:");
                        foreach (var doc in allDocsWithEmails)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (string email in doc.Emails)
                            {
                                Console.WriteLine("  " + email);
                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }


                    //Get all with both names and usernames
                    //This can be helpful to try and link usernames with names if they are ambiguous usernames
                    //But the usernames included may include commenters and editors, so will not always be definitive match
                    var allDocsWithBoth = documents.Where(p => p.ProbablyUsernames.Count() > 0 && p.ProbablyNames.Count() > 0);
                    if (allDocsWithBoth.Count() > 0)
                    {
                        Console.WriteLine("Docs containing usernames and names (possibly linked):");
                        foreach (var doc in allDocsWithBoth)
                        {
                            doc.AlreadySaidUsernamesAndNames = true;
                            Console.WriteLine(" File: " + doc.NameOfFile.Replace(workingFolder, ""));
                            Console.WriteLine("  Names:");
                            foreach (Name name in doc.ProbablyNames)
                            {
                                StringBuilder allLocations = new StringBuilder();
                                int HowMany = 0;
                                if (name.FromCreator)
                                {
                                    allLocations.Append("Creator");
                                    HowMany++;
                                }
                                if (name.FromFilepath)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Filepath");
                                    }
                                    else
                                    {
                                        allLocations.Append("Filepath");
                                    }
                                    HowMany++;
                                }
                                if (name.FromAuthor)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Author");
                                    }
                                    else
                                    {
                                        allLocations.Append("Author");

                                    }
                                    HowMany++;
                                }

                                if (name.FromExternalLink)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", External Links");
                                    }
                                    else
                                    {
                                        allLocations.Append("External Links");

                                    }
                                    HowMany++;
                                }
                                if (name.LastModifiedBy)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Last Modified By");
                                    }
                                    else
                                    {
                                        allLocations.Append("Last Modified By");

                                    }
                                    HowMany++;
                                }

                                if (name.PriorToVistaSet)
                                {
                                    if (name.PriorToVista)
                                    {
                                        Console.WriteLine("   " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Prior to Vista");
                                    }
                                    else
                                    {
                                        Console.WriteLine("   " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Later than Vista");
                                    }

                                }
                                else
                                {
                                    Console.WriteLine("   " + name.MyName + " [" + allLocations + "]");
                                }


                            }
                            Console.WriteLine("  Usernames:");
                            foreach (Name name in doc.ProbablyUsernames)
                            {
                                StringBuilder allLocations = new StringBuilder();
                                int HowMany = 0;
                                if (name.FromCreator)
                                {
                                    allLocations.Append("Creator");
                                    HowMany++;
                                }
                                if (name.FromFilepath)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Filepath");
                                    }
                                    else
                                    {
                                        allLocations.Append("Filepath");
                                    }
                                    HowMany++;
                                }
                                if (name.FromAuthor)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Author");
                                    }
                                    else
                                    {
                                        allLocations.Append("Author");

                                    }
                                    HowMany++;
                                }

                                if (name.FromExternalLink)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", External Links");
                                    }
                                    else
                                    {
                                        allLocations.Append("External Links");

                                    }
                                    HowMany++;
                                }
                                if (name.LastModifiedBy)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Last Modified By");
                                    }
                                    else
                                    {
                                        allLocations.Append("Last Modified By");

                                    }
                                    HowMany++;
                                }

                                if (name.PriorToVistaSet)
                                {
                                    if (name.PriorToVista)
                                    {
                                        Console.WriteLine("   " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Prior to Vista");
                                    }
                                    else
                                    {
                                        Console.WriteLine("   " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Later than Vista");
                                    }

                                }
                                else
                                {
                                    Console.WriteLine("   " + name.MyName + " [" + allLocations + "]");
                                }


                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all usernames that haven't been said above
                    var allDocsWithUsernames = documents.Where(p => p.ProbablyUsernames.Count() > 0 && p.AlreadySaidUsernamesAndNames != true);
                    if (allDocsWithUsernames.Count() > 0)
                    {
                        Console.WriteLine("Usernames:");
                        foreach (var doc in allDocsWithUsernames)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (Name name in doc.ProbablyUsernames)
                            {
                                StringBuilder allLocations = new StringBuilder();
                                int HowMany = 0;
                                if (name.FromCreator)
                                {
                                    allLocations.Append("Creator");
                                    HowMany++;
                                }
                                if (name.FromFilepath)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Filepath");
                                    }
                                    else
                                    {
                                        allLocations.Append("Filepath");
                                    }
                                    HowMany++;
                                }
                                if (name.FromAuthor)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Author");
                                    }
                                    else
                                    {
                                        allLocations.Append("Author");

                                    }
                                    HowMany++;
                                }

                                if (name.FromExternalLink)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", External Links");
                                    }
                                    else
                                    {
                                        allLocations.Append("External Links");

                                    }
                                    HowMany++;
                                }
                                if (name.LastModifiedBy)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Last Modified By");
                                    }
                                    else
                                    {
                                        allLocations.Append("Last Modified By");

                                    }
                                    HowMany++;
                                }

                                if (name.PriorToVistaSet)
                                {
                                    if (name.PriorToVista)
                                    {
                                        Console.WriteLine("  " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Prior to Vista");
                                    }
                                    else
                                    {
                                        Console.WriteLine("  " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Later than Vista");
                                    }

                                }
                                else
                                {
                                    Console.WriteLine("  " + name.MyName + " [" + allLocations + "]");
                                }


                            }

                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Get all names that haven't already been said
                    var allDocsWithNames = documents.Where(p => p.ProbablyNames.Count() > 0 && p.AlreadySaidUsernamesAndNames != true);
                    if (allDocsWithNames.Count() > 0)
                    {
                        Console.WriteLine("Names:");
                        foreach (var doc in allDocsWithNames)
                        {
                            Console.WriteLine(" " + doc.NameOfFile.Replace(workingFolder, ""));
                            foreach (Name name in doc.ProbablyNames)
                            {
                                StringBuilder allLocations = new StringBuilder();
                                int HowMany = 0;
                                if (name.FromCreator)
                                {
                                    allLocations.Append("Creator");
                                    HowMany++;
                                }
                                if (name.FromFilepath)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Filepath");
                                    }
                                    else
                                    {
                                        allLocations.Append("Filepath");
                                    }
                                    HowMany++;
                                }
                                if (name.FromAuthor)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Author");
                                    }
                                    else
                                    {
                                        allLocations.Append("Author");

                                    }
                                    HowMany++;
                                }

                                if (name.FromExternalLink)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", External Links");
                                    }
                                    else
                                    {
                                        allLocations.Append("External Links");

                                    }
                                    HowMany++;
                                }
                                if (name.LastModifiedBy)
                                {
                                    if (HowMany > 0)
                                    {
                                        allLocations.Append(", Last Modified By");
                                    }
                                    else
                                    {
                                        allLocations.Append("Last Modified By");

                                    }
                                    HowMany++;
                                }

                                if (name.PriorToVistaSet)
                                {
                                    if (name.PriorToVista)
                                    {
                                        Console.WriteLine("  " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Prior to Vista");
                                    }
                                    else
                                    {
                                        Console.WriteLine("  " + name.MyName + " [" + allLocations + "]" + " Filepath for OS Later than Vista");
                                    }

                                }
                                else
                                {
                                    Console.WriteLine("  " + name.MyName + " [" + allLocations + "]");
                                }


                            }

                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Print all unique names and usernames in a list for easy use in other tools
                    List<string> saidNames = new List<string>();
                    var allDocsWithNamesForList = documents.Where(p => p.ProbablyNames.Count() > 0);
                    if (allDocsWithNamesForList.Count() > 0)
                    {
                        Console.WriteLine("List of all names:");
                        foreach (Document c in allDocsWithNamesForList)
                        {
                            foreach (Name name in c.ProbablyNames)
                            {
                                if (saidNames.Contains(name.MyName))
                                {

                                }
                                else
                                {
                                    Console.WriteLine("  " + name.MyName);
                                    saidNames.Add(name.MyName);
                                }
                            }

                        }
                        Console.WriteLine(Environment.NewLine);
                    }
                    var allDocsWithUsernamesForList = documents.Where(p => p.ProbablyUsernames.Count() > 0);
                    if (allDocsWithUsernamesForList.Count() > 0)
                    {
                        Console.WriteLine("List of all usernames:");
                        foreach (Document c in allDocsWithUsernamesForList)
                        {
                            foreach (Name name in c.ProbablyUsernames)
                            {
                                if (saidNames.Contains(name.MyName))
                                {

                                }
                                else
                                {
                                    Console.WriteLine("  " + name.MyName);
                                    saidNames.Add(name.MyName);
                                }
                            }

                        }
                        Console.WriteLine(Environment.NewLine);
                    }

                    //Print all emails for easier copy/paste
                    //Emails
                    List<string> alreadyEmails = new List<string>();
                    var allDocsWithEmailsForList = documents.Where(p => p.Emails.Count() > 0);
                    if (allDocsWithEmailsForList.Count() > 0)
                    {
                        Console.WriteLine("Emails:");
                        foreach (var doc in allDocsWithEmails)
                        {
                            foreach (string email in doc.Emails)
                            {
                                if (alreadyEmails.Contains(email))
                                {

                                }
                                else
                                {
                                    Console.WriteLine("  " + email);
                                    alreadyEmails.Add(email);
                                }

                            }
                        }
                        Console.WriteLine(Environment.NewLine);
                    }
                }
                else
                {
                    Console.WriteLine("No metadata extracted from any document...");
                }

            }
            Console.WriteLine("Finished...");
        }


        static void RunTheJewels(string file)
        {
            Document myDoc = new Document() { NameOfFile = file };
            try
            {
                if (Directory.Exists(workingFolder + "\\Temp\\"))
                {
                    Directory.Delete(workingFolder + "\\Temp\\", true);
                }
                Directory.CreateDirectory(workingFolder + "\\Temp\\");
                ZipFile.ExtractToDirectory(file, workingFolder + "\\Temp\\");

                //Extract Media and embedded docs if chosen
                if (extractEmbedded)
                {

                    foreach (string dir in Directory.EnumerateDirectories(workingFolder, "*.*", SearchOption.AllDirectories))
                    {
                        if (dir.Contains("embed"))
                        {

                            Directory.CreateDirectory(workingFolder + "Embed\\Doc" + embedDocNumber + "\\");

                            foreach (string fileFromEmbed in Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories))
                            {
                                File.Copy(fileFromEmbed, workingFolder + "Embed\\Doc" + embedDocNumber + "\\" + Path.GetFileName(fileFromEmbed));
                            }
                            embedDocNumber++;
                        }
                    }
                }
                if (extractMedia)
                {

                    foreach (string dir in Directory.EnumerateDirectories(workingFolder, "*.*", SearchOption.AllDirectories))
                    {
                        if (dir.Contains("media"))
                        {

                            Directory.CreateDirectory(workingFolder + "Media\\Doc" + mediaDocNumber + "\\");

                            foreach (string fileFromMedia in Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories))
                            {
                                File.Copy(fileFromMedia, workingFolder + "Media\\Doc" + mediaDocNumber + "\\" + Path.GetFileName(fileFromMedia));
                            }
                            mediaDocNumber++;
                        }
                    }
                }


                //Search each extracted file
                foreach (string fileFromZip in Directory.EnumerateFiles(workingFolder + "\\Temp\\", "*.*", SearchOption.AllDirectories))
                {
                    //Added try catch here so if an individual internal file fails it keeps trying for others
                    try
                    {
                        //These are the expected file types from the extracted files
                        if (fileFromZip.Contains("xml") || fileFromZip.Contains(".rels"))
                        {
                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.Load(fileFromZip);

                            //Perform basic grepping for things we'd want
                            if (grep)
                            {
                                string t = xmlDoc.OuterXml;
                                //Look for patterns indicating users or hostnames just on all of every internal file
                                LookForUsers(t, myDoc);
                                LookForHostnames(t, myDoc);

                                //If the internal file contains the word password, record this.
                                if (t.Contains("password"))
                                {
                                    myDoc.ContainsPassword = true;
                                    myDoc.PasswordContainingInnerFileNames.Add(fileFromZip);
                                    //Construct the string of each tag
                                    foreach (XmlNode node in xmlDoc.SelectNodes("*"))
                                    {
                                        foreach (XmlNode child in node.ChildNodes)
                                        {
                                            if (child.OuterXml.Contains("password"))
                                            {
                                                if (child.HasChildNodes)
                                                {
                                                    XmlNodeList childNodes2 = child.ChildNodes;
                                                    foreach (XmlNode n2 in childNodes2)
                                                    {
                                                        if (n2.HasChildNodes)
                                                        {
                                                            if (n2.OuterXml.Contains("password"))
                                                            {
                                                                myDoc.PasswordContainingString.Add(n2.OuterXml);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            string full = "<" + child.Name + ">" + n2.OuterXml + "</" + child.Name + ">";
                                                            if (full.Contains("password"))
                                                            {
                                                                myDoc.PasswordContainingString.Add(full);
                                                            }
                                                        }

                                                    }
                                                }
                                                else
                                                {
                                                    myDoc.PasswordContainingString.Add(child.OuterXml);
                                                }
                                            }


                                        }
                                    }

                                    myDoc.MetaDataFound = true;
                                }

                                //Now just use this to add the file paths - don't need to look for users/hosts on this cos already doing on all of t
                                if (t.Contains("absPath"))
                                {
                                    Regex rx66 = new Regex("(?<=absPath url=\")(.*?\")");
                                    MatchCollection matches66 = rx66.Matches(t);
                                    if (matches66.Count > 0)
                                    {
                                        string getLocalFilePath = matches66[0].Value.Replace("\"", "");
                                        myDoc.FilePaths.Add(getLocalFilePath);
                                        myDoc.MetaDataFound = true;

                                    }
                                }

                                //User grep
                                if (additionalGrep != "" && additionalGrep != null)
                                {
                                    //Construct the string of each tag
                                    foreach (XmlNode node in xmlDoc.SelectNodes("*"))
                                    {
                                        foreach (XmlNode child in node.ChildNodes)
                                        {
                                            if (child.OuterXml.Contains(additionalGrep))
                                            {
                                                if (child.HasChildNodes)
                                                {
                                                    XmlNodeList childNodes2 = child.ChildNodes;
                                                    foreach (XmlNode n2 in childNodes2)
                                                    {
                                                        if (n2.HasChildNodes)
                                                        {
                                                            if (n2.OuterXml.Contains(additionalGrep))
                                                            {
                                                                myDoc.UserGrep.Add(new GrepObject() { GrepString = n2.OuterXml, InnerFile = fileFromZip });
                                                            }
                                                        }
                                                        else
                                                        {
                                                            string full = "<" + child.Name + ">" + n2.OuterXml + "</" + child.Name + ">";
                                                            if (full.Contains(additionalGrep))
                                                            {
                                                                myDoc.UserGrep.Add(new GrepObject() { GrepString = full, InnerFile = fileFromZip });
                                                            }
                                                        }

                                                    }
                                                }
                                                else
                                                {
                                                    myDoc.UserGrep.Add(new GrepObject() { GrepString = child.OuterXml, InnerFile = fileFromZip });
                                                }

                                                myDoc.MetaDataFound = true;
                                            }
                                        }
                                    }
                                }

                                //External Links
                                if (t.Contains("Target"))
                                {
                                    Regex rx00 = new Regex("(?<=Target=\")(.*)(.*?\" Target)");
                                    MatchCollection matches00 = rx00.Matches(t);
                                    if (matches00.Count > 0)
                                    {
                                        string getExternalLink = matches00[0].Value.Replace("\" Target", "");
                                        if (getExternalLink.StartsWith("../") || getExternalLink.StartsWith("docProps") || getExternalLink.StartsWith("styles.xml") || getExternalLink.StartsWith("theme/") || getExternalLink.StartsWith("worksheets/") || getExternalLink.StartsWith("presProps.xml") || getExternalLink.StartsWith("slide") || getExternalLink.StartsWith("tableStyles.xml") || getExternalLink.StartsWith("webSettings.xml") || getExternalLink.StartsWith("footnotes") || getExternalLink.StartsWith("media/") || getExternalLink.StartsWith("header") || getExternalLink.StartsWith("settings") || getExternalLink.StartsWith("footnotes") || getExternalLink.StartsWith("fontTable") || getExternalLink.StartsWith("endnotes") || getExternalLink.StartsWith("stylewitheffects") || getExternalLink.StartsWith("word") || getExternalLink.StartsWith("notesMasters") || getExternalLink.StartsWith("ppt") || getExternalLink.StartsWith("style"))
                                        {

                                        }
                                        else
                                        {
                                            if (getExternalLink.Contains("/>"))
                                            {
                                                string[] link = getExternalLink.Split('>');
                                                if (link[0].EndsWith("/"))
                                                {
                                                    myDoc.ExternalLinks.Add(link[0].Remove(link[0].Length - 1, 1));
                                                }

                                            }
                                            else
                                            {
                                                myDoc.ExternalLinks.Add(getExternalLink);
                                            }

                                            myDoc.MetaDataFound = true;
                                        }
                                    }

                                }
                            }

                            if(additionalGrep != "" && additionalGrep != null && !grep)
                            {
                                //User has supplied a string for us to search for - but this is not being handled by the above grep functionality as NoGrep was specified
                                //Construct the string of each tag
                                foreach (XmlNode node in xmlDoc.SelectNodes("*"))
                                {
                                    foreach (XmlNode child in node.ChildNodes)
                                    {

                                        XmlNodeList childNodes2 = child.ChildNodes;
                                        foreach (XmlNode n2 in childNodes2)
                                        {
                                            string full = "<" + child.Name + ">" + n2.OuterXml + "</" + child.Name + ">";
                                            if (full.Contains(additionalGrep))
                                            {
                                                myDoc.UserGrep.Add(new GrepObject() { GrepString = full, InnerFile = fileFromZip });
                                            }
                                        }

                                    }
                                }
                            }


                            //Continue with specific metadata extraction
                            //Get any hidden sheets
                            //This is already essentially acting like grep above so does not need to be added to grep
                            XmlNodeList sheetForHidden = xmlDoc.GetElementsByTagName("sheet");
                            foreach (XmlNode node in sheetForHidden)
                            {
                                if (node.OuterXml.Contains("hidden"))
                                {
                                    Regex rx = new Regex("(?<=name=\")(.*?\")");
                                    MatchCollection matches = rx.Matches(node.OuterXml);
                                    if (matches.Count > 0)
                                    {
                                        string getSheetName = matches[0].Value.Replace("\"", "");
                                        myDoc.HiddenSheet.Add(new HiddenSheet() { HiddenSheetName = getSheetName, HiddenSheetLocation = file });
                                        myDoc.MetaDataFound = true;
                                    }

                                }
                            }

                            //Image links
                            if (fileFromZip.Contains("document"))
                            {
                                XmlNodeList sheetForComments = xmlDoc.GetElementsByTagName("w:document");
                                if (sheetForComments.Count > 0)
                                {
                                    foreach (XmlNode node in sheetForComments)
                                    {
                                        XmlNodeList childNodes = node.ChildNodes;

                                        foreach (XmlNode n in childNodes)
                                        {
                                            XmlNodeList childNodes2 = node.ChildNodes;
                                            foreach (XmlNode n2 in childNodes2)
                                            {
                                                if (n2.OuterXml.Contains("descr"))
                                                {

                                                    LookForHostnames(n2.OuterXml, myDoc);
                                                    LookForUsers(n2.OuterXml, myDoc);
                                                }
                                            }
                                        }


                                    }
                                }
                            }

                            //Target from settings
                            if (fileFromZip.Contains("settings"))
                            {
                                XmlNodeList sheetForSettings = xmlDoc.GetElementsByTagName("Relationships");
                                if (sheetForSettings.Count > 0)
                                {
                                    foreach (XmlNode node in sheetForSettings)
                                    {
                                        XmlNodeList childNodes = node.ChildNodes;

                                        foreach (XmlNode n in childNodes)
                                        {
                                            XmlNodeList childNodes2 = node.ChildNodes;
                                            foreach (XmlNode n2 in childNodes2)
                                            {
                                                if (n2.OuterXml.Contains("Target"))
                                                {
                                                    LookForUsers(n2.OuterXml, myDoc);
                                                }
                                            }
                                        }

                                    }
                                }
                            }

                            //Creator
                            if (fileFromZip.Contains("core"))
                            {
                                XmlNodeList sheetForComments = xmlDoc.GetElementsByTagName("cp:coreProperties");
                                if (sheetForComments.Count > 0)
                                {
                                    foreach (XmlNode node in sheetForComments)
                                    {
                                        if (node.OuterXml.Contains("cp:lastModifiedBy"))
                                        {
                                            LookForUsers(node.OuterXml, myDoc);
                                        }

                                    }
                                }
                            }

                            //Get last saved location
                            XmlNodeList sheetForAbs = xmlDoc.GetElementsByTagName("workbook");
                            if (sheetForAbs.Count > 0)
                            {
                                foreach (XmlNode node in sheetForAbs)
                                {
                                    XmlNodeList childNodes = node.ChildNodes;

                                    foreach (XmlNode n in childNodes)
                                    {
                                        if (n.OuterXml.Contains("absPath"))
                                        {
                                            Regex rx = new Regex("(?<=absPath url=\")(.*?\")");
                                            MatchCollection matches = rx.Matches(n.OuterXml);
                                            if (matches.Count > 0)
                                            {
                                                string getLocalFilePath = matches[0].Value.Replace("\"", "");
                                                myDoc.FilePaths.Add(getLocalFilePath);

                                                if (getLocalFilePath.Contains("Users"))
                                                {
                                                    LookForUsers(getLocalFilePath, myDoc);
                                                }

                                                if (getLocalFilePath.Contains("\\\\"))
                                                {
                                                    LookForHostnames(getLocalFilePath, myDoc);
                                                }
                                            }
                                        }
                                    }

                                }
                            }

                            //Comments file
                            if (fileFromZip.Contains("comments"))
                            {
                                XmlNodeList sheetForComments = xmlDoc.GetElementsByTagName("comments");
                                if (sheetForComments.Count > 0)
                                {
                                    foreach (XmlNode node in sheetForComments)
                                    {
                                        XmlNodeList childNodes = node.ChildNodes;

                                        foreach (XmlNode n in childNodes)
                                        {
                                            if (n.OuterXml.Contains("author"))
                                            {
                                                LookForUsers(n.OuterXml, myDoc);
                                            }
                                        }
                                    }
                                }
                            }

                            //External links file
                            if (fileFromZip.Contains("externalLink") && fileFromZip.Contains("xml.rels"))
                            {
                                XmlNodeList sheetForExternalLinks = xmlDoc.GetElementsByTagName("Relationships");
                                if (sheetForExternalLinks.Count > 0)
                                {
                                    foreach (XmlNode node in sheetForExternalLinks)
                                    {
                                        XmlNodeList childNodes = node.ChildNodes;

                                        foreach (XmlNode n in childNodes)
                                        {
                                            if (n.OuterXml.Contains("Target"))
                                            {
                                                Regex rx = new Regex("(?<=Target=\")(.*)(.*?\" Target)");
                                                MatchCollection matches = rx.Matches(n.OuterXml);
                                                if (matches.Count > 0)
                                                {
                                                    string getExternalLink = matches[0].Value.Replace("\" Target", "");
                                                    if (getExternalLink.StartsWith("../") || getExternalLink.StartsWith("docProps") || getExternalLink.StartsWith("styles.xml") || getExternalLink.StartsWith("theme/") || getExternalLink.StartsWith("worksheets/") || getExternalLink.StartsWith("presProps.xml") || getExternalLink.StartsWith("slides/"))
                                                    {

                                                    }
                                                    else
                                                    {
                                                        myDoc.ExternalLinks.Add(getExternalLink);
                                                        myDoc.MetaDataFound = true;
                                                        LookForUsers(getExternalLink, myDoc);
                                                        LookForHostnames(getExternalLink, myDoc);
                                                    }
                                                }
                                            }
                                        }


                                    }
                                }
                            }
                        }

                        //Maybe refactor if I start getting other bits with like vba or whatever too!
                        if (fileFromZip.Contains(".bin"))
                        {
                            byte[] allBytes = File.ReadAllBytes(fileFromZip);
                            string decodedString2 = Encoding.UTF8.GetString(allBytes).Replace("\0", "");

                            Regex rxB = new Regex("(?<=)(.*?\\u0001)");
                            MatchCollection matchesB = rxB.Matches(decodedString2);
                            if (matchesB.Count > 0)
                            {
                                string printer = matchesB[0].Value.Replace("\u0001", "");
                                if (printer != "" && printer != null)
                                {
                                    if (printer.Contains("Microsoft Print to PDF"))
                                    {

                                    }
                                    else
                                    {
                                        myDoc.Printers.Add(printer);
                                        myDoc.MetaDataFound = true;
                                    }
                                }
                            }


                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.ToString());
                    }

                }
            }
            catch (Exception e)
            {
                if (e.ToString().Contains("End of Central Directory record could not be found") || e.ToString().Contains("Number of entries expected in End Of Central Directory does not correspond to number of entries in Central Directory") || e.ToString().Contains("Central Directory corrupt"))
                {
                    Console.WriteLine("*Corrupted file: " + Path.GetFileName(file));
                }
                else
                {
                    Console.WriteLine("  Error: " + e.ToString());
                }
            }


            documents.Add(myDoc);
        }

        private static void LookForHostnames(string t, Document myDoc)
        {
            try
            {
                Regex rxH22 = new Regex(@"(?<=\\\\)(.*?\\)");
                MatchCollection matchesH22 = rxH22.Matches(t);
                if (matchesH22.Count > 0)
                {
                    foreach (Match match in matchesH22)
                    {
                        string hostname = "\\\\" + match.Value.Replace("\\", "");
                        if (hostname != "")
                        {
                            myDoc.Hostnames.Add(hostname);
                            myDoc.MetaDataFound = true;
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private static void LookForUsers(string t, Document myDoc)
        {
            Regex rxU3333 = new Regex(@"(?<=documents%20and%20settings\\)(.*?\\)");
            MatchCollection matchesU3333 = rxU3333.Matches(t);
            if (matchesU3333.Count > 0)
            {
                foreach (Match match in matchesU3333)
                {
                    bool isUsername = false;
                    string userN = match.Value.Replace("\\", "");
                    if (userN != "" && userN != "." && userN != "<cp:keywords></cp:keywords><dc:description>")
                    {
                        if (userN.Contains(" "))
                        {
                            if (userN.StartsWith(" "))
                            {
                                myDoc.AddName(userN, true);
                                isUsername = true;
                                myDoc.MetaDataFound = true;
                            }
                            else
                            {
                                myDoc.AddName(userN, false);
                                myDoc.MetaDataFound = true;
                            }
                        }
                        else
                        {
                            myDoc.AddName(userN, true);
                            isUsername = true;
                            myDoc.MetaDataFound = true;
                        }
                    }
                    if (isUsername)
                    {
                        foreach (Name n in myDoc.ProbablyUsernames)
                        {
                            if (n.MyName == userN)
                            {
                                n.PriorToVista = true;
                                n.PriorToVistaSet = true;
                                n.FromFilepath = true;
                            }
                        }
                    }
                    else
                    {
                        foreach (Name n in myDoc.ProbablyNames)
                        {
                            if (n.MyName == userN)
                            {
                                n.PriorToVista = true;
                                n.PriorToVistaSet = true;
                                n.FromFilepath = true;
                            }
                        }
                    }
                }

            }

            //Docume
            Regex rxU33333 = new Regex(@"(?<=DOCUME~1/)(.*?/)");
            MatchCollection matchesU33333 = rxU33333.Matches(t);
            if (matchesU33333.Count > 0)
            {
                foreach (Match match in matchesU33333)
                {
                    bool isUsername = false;
                    string userN = match.Value.Replace("/", "");
                    if (userN != "" && userN != ".")
                    {
                        if (userN.Contains(" "))
                        {
                            if (userN.StartsWith(" "))
                            {
                                myDoc.AddName(userN, true);
                                isUsername = true;
                                myDoc.MetaDataFound = true;
                            }
                            else
                            {
                                myDoc.AddName(userN, false);
                                myDoc.MetaDataFound = true;
                            }
                        }
                        else
                        {
                            myDoc.AddName(userN, true);
                            isUsername = true;
                            myDoc.MetaDataFound = true;
                        }
                    }
                    if (isUsername)
                    {
                        foreach (Name n in myDoc.ProbablyUsernames)
                        {
                            if (n.MyName == userN)
                            {
                                n.FromExternalLink = true;
                                n.PriorToVista = true;
                                n.PriorToVistaSet = true;
                            }
                        }
                    }
                    else
                    {
                        foreach (Name n in myDoc.ProbablyNames)
                        {
                            if (n.MyName == userN)
                            {
                                n.FromExternalLink = true;
                                n.PriorToVista = true;
                                n.PriorToVistaSet = true;
                            }
                        }
                    }
                }

            }


            //Email from mailto
            Regex rxe = new Regex(@"(?<=mailto:)(.*? )");
            MatchCollection matchese = rxe.Matches(t);
            if (matchese.Count > 0)
            {
                foreach (Match match in matchese)
                {
                    string userN = match.Value.Replace("/", "").Replace("\"", "");
                    if (userN != "" && userN != ".")
                    {
                        myDoc.Emails.Add(userN);

                    }

                }

            }

            //Other Users
            Regex rxU33 = new Regex(@"(?<=Users\\)(.*?\\)");
            MatchCollection matchesU33 = rxU33.Matches(t);
            if (matchesU33.Count > 0)
            {
                foreach (Match match in matchesU33)
                {
                    bool isUsername = false;
                    string userN = matchesU33[0].Value.Replace("\\", "");
                    if (userN != "" && userN != ".")
                    {

                        if (userN.Contains(" "))
                        {
                            if (userN.StartsWith(" "))
                            {
                                myDoc.AddName(userN, true);
                                isUsername = true;
                                myDoc.MetaDataFound = true;
                            }
                            else
                            {
                                myDoc.AddName(userN, false);
                                myDoc.MetaDataFound = true;
                            }
                        }
                        else
                        {
                            myDoc.AddName(userN, true);
                            isUsername = true;
                            myDoc.MetaDataFound = true;
                        }
                    }
                    if (isUsername)
                    {
                        foreach (Name n in myDoc.ProbablyUsernames)
                        {
                            if (n.MyName == userN)
                            {
                                n.PriorToVista = false;
                                n.PriorToVistaSet = true;
                                n.FromFilepath = true;
                            }
                        }
                    }
                    else
                    {
                        foreach (Name n in myDoc.ProbablyNames)
                        {
                            if (n.MyName == userN)
                            {
                                n.PriorToVista = false;
                                n.PriorToVistaSet = true;
                                n.FromFilepath = true;
                            }
                        }
                    }
                }


            }

            //Users
            Regex rxU333 = new Regex(@"(?<=Users/)(.*?/)");
            MatchCollection matchesU333 = rxU333.Matches(t);
            if (matchesU333.Count > 0)
            {
                foreach (Match match in matchesU333)
                {
                    bool isUsername = false;
                    string userN = match.Value.Replace("/", "");
                    if (userN != "" && userN != "." && userN != "<cp:keywords></cp:keywords><dc:description>")
                    {
                        if (userN.Contains(" "))
                        {
                            if (userN.StartsWith(" "))
                            {
                                myDoc.AddName(userN, true);
                                isUsername = true;
                                myDoc.MetaDataFound = true;
                            }
                            else
                            {
                                myDoc.AddName(userN, false);
                                myDoc.MetaDataFound = true;
                            }
                        }
                        else
                        {
                            myDoc.AddName(userN, true);
                            isUsername = true;
                            myDoc.MetaDataFound = true;
                        }
                    }
                    if (isUsername)
                    {
                        foreach (Name n in myDoc.ProbablyUsernames)
                        {
                            if (n.MyName == userN)
                            {
                                n.PriorToVista = false;
                                n.PriorToVistaSet = true;
                                n.FromFilepath = true;
                            }
                        }
                    }
                    else
                    {
                        foreach (Name n in myDoc.ProbablyNames)
                        {
                            if (n.MyName == userN)
                            {
                                n.PriorToVista = false;
                                n.PriorToVistaSet = true;
                                n.FromFilepath = true;
                            }
                        }
                    }
                }


            }

            if (t.Contains("creator"))
            {
                Regex rxU44 = new Regex(@"(?<=<dc:creator>)(.*?</dc)");
                MatchCollection matchesU44 = rxU44.Matches(t);
                if (matchesU44.Count > 0)
                {
                    foreach (Match match in matchesU44)
                    {
                        bool isUsername = false;
                        bool weirdBool = false;
                        string userN1 = match.Value.Replace("</dc", "");
                        if (userN1 != "" && userN1 != "." && userN1 != "<cp:keywords></cp:keywords><dc:description>")
                        {
                            if (userN1.Contains("lastModifiedBy"))
                            {
                                Regex rxU551 = new Regex(@"(?<=cp:lastModifiedBy>)(.*?</cp:lastModifiedBy)");
                                MatchCollection matchesU551 = rxU551.Matches(userN1);
                                if (matchesU551.Count > 0)
                                {
                                    foreach (Match match1 in matchesU551)
                                    {
                                        bool isUsername1 = false;
                                        bool weirdBool1 = false;
                                        string userN11 = match1.Value.Replace("</cp:lastModifiedBy", "");
                                        if (userN11 != "")
                                        {
                                            if (userN11.Contains("-"))
                                            {
                                                string[] both = userN11.Split('-');
                                                string name = both[0];
                                                string username = both[1];
                                                if (username.StartsWith(" "))
                                                {
                                                    //Probably name in first bit and username in second bit!
                                                    myDoc.AddName(username, true);
                                                    myDoc.MetaDataFound = true;
                                                    myDoc.AddName(name, false);
                                                    weirdBool1 = true;
                                                    foreach (Name n in myDoc.ProbablyUsernames)
                                                    {
                                                        if (n.MyName == username)
                                                        {
                                                            n.FromAuthor = true;
                                                        }
                                                    }
                                                    foreach (Name n in myDoc.ProbablyNames)
                                                    {
                                                        if (n.MyName == name)
                                                        {
                                                            n.FromAuthor = true;
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    //So this might be a weird name that contains - anyway...
                                                    myDoc.AddName(userN1, false);
                                                    myDoc.MetaDataFound = true;
                                                    weirdBool1 = true;
                                                    foreach (Name n in myDoc.ProbablyNames)
                                                    {
                                                        if (n.MyName == userN1)
                                                        {
                                                            n.FromAuthor = true;
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {

                                                if (userN11.Contains(" "))
                                                {
                                                    if (userN11.StartsWith(" "))
                                                    {
                                                        myDoc.AddName(userN11, true);
                                                        isUsername1 = true;
                                                        myDoc.MetaDataFound = true;
                                                    }
                                                    else
                                                    {
                                                        myDoc.AddName(userN11, false);
                                                        myDoc.MetaDataFound = true;
                                                    }
                                                }
                                                else
                                                {
                                                    myDoc.AddName(userN11, true);
                                                    isUsername1 = true;
                                                    myDoc.MetaDataFound = true;
                                                }
                                            }
                                        }
                                        if (!weirdBool1)
                                        {
                                            if (isUsername1)
                                            {
                                                foreach (Name n in myDoc.ProbablyUsernames)
                                                {
                                                    if (n.MyName == userN11)
                                                    {
                                                        n.LastModifiedBy = true;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (Name n in myDoc.ProbablyNames)
                                                {
                                                    if (n.MyName == userN11)
                                                    {
                                                        n.LastModifiedBy = true;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                            else
                            {
                                if (userN1.Contains("-"))
                                {
                                    string[] both = userN1.Split('-');
                                    string name = both[0];
                                    string username = both[1];
                                    if (username.StartsWith(" "))
                                    {
                                        //Probably name in first bit and username in second bit!
                                        myDoc.AddName(username, true);
                                        myDoc.MetaDataFound = true;
                                        myDoc.AddName(name, false);
                                        weirdBool = true;
                                        foreach (Name n in myDoc.ProbablyUsernames)
                                        {
                                            if (n.MyName == username)
                                            {
                                                n.FromAuthor = true;
                                            }
                                        }
                                        foreach (Name n in myDoc.ProbablyNames)
                                        {
                                            if (n.MyName == name)
                                            {
                                                n.FromAuthor = true;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //So this might be a weird name that contains - anyway...
                                        myDoc.AddName(userN1, false);
                                        myDoc.MetaDataFound = true;
                                        weirdBool = true;
                                        foreach (Name n in myDoc.ProbablyNames)
                                        {
                                            if (n.MyName == userN1)
                                            {
                                                n.FromAuthor = true;
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (userN1.Contains(" "))
                                    {
                                        if (userN1.StartsWith(" "))
                                        {
                                            myDoc.AddName(userN1, true);
                                            isUsername = true;
                                            myDoc.MetaDataFound = true;
                                        }
                                        else
                                        {
                                            myDoc.AddName(userN1, false);
                                            myDoc.MetaDataFound = true;
                                        }
                                    }
                                    else
                                    {
                                        myDoc.AddName(userN1, true);
                                        isUsername = true;
                                        myDoc.MetaDataFound = true;
                                    }
                                }

                            }

                        }
                        if (!weirdBool)
                        {
                            if (isUsername)
                            {
                                foreach (Name n in myDoc.ProbablyUsernames)
                                {
                                    if (n.MyName == userN1)
                                    {
                                        n.FromCreator = true;
                                    }
                                }
                            }
                            else
                            {
                                foreach (Name n in myDoc.ProbablyNames)
                                {
                                    if (n.MyName == userN1)
                                    {
                                        n.FromCreator = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (t.Contains("cp:lastModifiedBy"))
            {
                Regex rxU55 = new Regex(@"(?<=cp:lastModifiedBy>)(.*?</cp:lastModifiedBy)");
                MatchCollection matchesU55 = rxU55.Matches(t);
                if (matchesU55.Count > 0)
                {
                    foreach (Match match in matchesU55)
                    {
                        bool isUsername = false;
                        bool weirdBool = false;
                        string userN1 = match.Value.Replace("</cp:lastModifiedBy", "");
                        if (userN1 != "" && userN1 != "." && userN1 != "<cp:keywords></cp:keywords><dc:description>")
                        {
                            if (userN1.Contains("-"))
                            {
                                string[] both = userN1.Split('-');
                                string name = both[0];
                                string username = both[1];
                                if (username.StartsWith(" "))
                                {
                                    //Probably name in first bit and username in second bit!
                                    myDoc.AddName(username, true);
                                    myDoc.MetaDataFound = true;
                                    myDoc.AddName(name, false);
                                    weirdBool = true;
                                    foreach (Name n in myDoc.ProbablyUsernames)
                                    {
                                        if (n.MyName == username)
                                        {
                                            n.FromAuthor = true;
                                        }
                                    }
                                    foreach (Name n in myDoc.ProbablyNames)
                                    {
                                        if (n.MyName == name)
                                        {
                                            n.FromAuthor = true;
                                        }
                                    }
                                }
                                else
                                {
                                    //So this might be a weird name that contains - anyway...
                                    myDoc.AddName(userN1, false);
                                    myDoc.MetaDataFound = true;
                                    weirdBool = true;
                                    foreach (Name n in myDoc.ProbablyNames)
                                    {
                                        if (n.MyName == userN1)
                                        {
                                            n.FromAuthor = true;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (userN1.Contains(" "))
                                {
                                    if (userN1.StartsWith(" "))
                                    {
                                        myDoc.AddName(userN1, true);
                                        isUsername = true;
                                        myDoc.MetaDataFound = true;
                                    }
                                    else
                                    {
                                        myDoc.AddName(userN1, false);
                                        myDoc.MetaDataFound = true;
                                    }
                                }
                                else
                                {
                                    myDoc.AddName(userN1, true);
                                    isUsername = true;
                                    myDoc.MetaDataFound = true;
                                }
                            }
                        }
                        if (!weirdBool)
                        {
                            if (isUsername)
                            {
                                foreach (Name n in myDoc.ProbablyUsernames)
                                {
                                    if (n.MyName == userN1)
                                    {
                                        n.LastModifiedBy = true;
                                    }
                                }
                            }
                            else
                            {
                                foreach (Name n in myDoc.ProbablyNames)
                                {
                                    if (n.MyName == userN1)
                                    {
                                        n.LastModifiedBy = true;
                                    }
                                }
                            }
                        }
                    }
                }


            }

            if (t.Contains("author"))
            {
                Regex rx99 = new Regex("(?<=<author>)(.*?</author>)");
                MatchCollection matches99 = rx99.Matches(t);
                if (matches99.Count > 0)
                {
                    foreach (Match match in matches99)
                    {
                        bool isUsername = false;
                        bool weirdBool = false;
                        string getUsername = match.Value.Replace("</author>", "");
                        if (getUsername != "" && getUsername != "." && getUsername != "<cp:keywords></cp:keywords><dc:description>")
                        {
                            if (getUsername.Contains(" "))
                            {
                                //NEED TO SPLIT - username FROM AUTHOR
                                if (getUsername.Contains("-"))
                                {
                                    string[] both = getUsername.Split('-');
                                    string name = both[0];
                                    string username = both[1];
                                    if (username.StartsWith(" "))
                                    {
                                        //Probably name in first bit and username in second bit!
                                        myDoc.AddName(username, true);
                                        myDoc.MetaDataFound = true;
                                        myDoc.AddName(name, false);
                                        weirdBool = true;
                                        foreach (Name n in myDoc.ProbablyUsernames)
                                        {
                                            if (n.MyName == username)
                                            {
                                                n.FromAuthor = true;
                                            }
                                        }
                                        foreach (Name n in myDoc.ProbablyNames)
                                        {
                                            if (n.MyName == name)
                                            {
                                                n.FromAuthor = true;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //So this might be a weird name that contains - anyway...
                                        myDoc.AddName(getUsername, false);
                                        myDoc.MetaDataFound = true;
                                        weirdBool = true;
                                        foreach (Name n in myDoc.ProbablyNames)
                                        {
                                            if (n.MyName == getUsername)
                                            {
                                                n.FromAuthor = true;
                                            }
                                        }
                                    }
                                }
                                else
                                {

                                    if (getUsername.StartsWith(" "))
                                    {
                                        myDoc.AddName(getUsername, true);
                                        isUsername = true;
                                        myDoc.MetaDataFound = true;
                                    }
                                    else
                                    {
                                        myDoc.AddName(getUsername, false);
                                        myDoc.MetaDataFound = true;
                                    }
                                }
                            }
                            else
                            {
                                myDoc.AddName(getUsername, true);
                                isUsername = true;
                                myDoc.MetaDataFound = true;
                            }
                        }
                        if (!weirdBool)
                        {
                            if (isUsername)
                            {
                                foreach (Name n in myDoc.ProbablyUsernames)
                                {
                                    if (n.MyName == getUsername)
                                    {
                                        n.FromAuthor = true;
                                    }
                                }
                            }
                            else
                            {
                                foreach (Name n in myDoc.ProbablyNames)
                                {
                                    if (n.MyName == getUsername)
                                    {
                                        n.FromAuthor = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }

        }
    }

    public class Name
    {
        string name;
        bool fromCreator = false;
        bool fromAuthor = false;
        bool fromFilepath = false;
        bool priorToVista = false;
        bool priorToVistaSet = false;
        bool fromExternalLink = false;
        bool lastModifiedBy = false;

        public string MyName { get => name; set => name = value; }
        public bool FromCreator { get => fromCreator; set => fromCreator = value; }
        public bool FromAuthor { get => fromAuthor; set => fromAuthor = value; }
        public bool FromFilepath { get => fromFilepath; set => fromFilepath = value; }
        public bool PriorToVista { get => priorToVista; set => priorToVista = value; }
        public bool PriorToVistaSet { get => priorToVistaSet; set => priorToVistaSet = value; }
        public bool FromExternalLink { get => fromExternalLink; set => fromExternalLink = value; }
        public bool LastModifiedBy { get => lastModifiedBy; set => lastModifiedBy = value; }
    }

    public class Document
    {
        string nameOfFile;
        bool metaDataFound = false;
        bool alreadySaidUsernamesAndNames = false;
        List<Name> probablyUsernames = new List<Name>();
        List<Name> probablyNames = new List<Name>();
        List<HiddenSheet> hiddenSheet = new List<HiddenSheet>();
        bool containsPassword = false;
        List<string> passwordContainingInnerFileNames = new List<string>();
        List<string> passwordContainingString = new List<string>();
        List<string> filePaths = new List<string>();
        List<string> externalLinks = new List<string>();
        List<string> hostnames = new List<string>();
        List<string> printers = new List<string>();
        List<string> emails = new List<string>();
        List<GrepObject> userGrep = new List<GrepObject>();

        public string NameOfFile { get => nameOfFile; set => nameOfFile = value; }
        public List<Name> ProbablyUsernames { get => probablyUsernames; set => probablyUsernames = value; }
        public List<Name> ProbablyNames { get => probablyNames; set => probablyNames = value; }
        public List<HiddenSheet> HiddenSheet { get => hiddenSheet; set => hiddenSheet = value; }
        public bool ContainsPassword { get => containsPassword; set => containsPassword = value; }
        public List<string> PasswordContainingString { get => passwordContainingString; set => passwordContainingString = value; }
        public List<string> FilePaths { get => filePaths; set => filePaths = value; }
        public List<string> ExternalLinks { get => externalLinks; set => externalLinks = value; }
        public List<string> Hostnames { get => hostnames; set => hostnames = value; }
        public List<string> Printers { get => printers; set => printers = value; }
        public bool MetaDataFound { get => metaDataFound; set => metaDataFound = value; }
        public List<string> PasswordContainingInnerFileNames { get => passwordContainingInnerFileNames; set => passwordContainingInnerFileNames = value; }
        public bool AlreadySaidUsernamesAndNames { get => alreadySaidUsernamesAndNames; set => alreadySaidUsernamesAndNames = value; }
        public List<string> Emails { get => emails; set => emails = value; }
        public List<GrepObject> UserGrep { get => userGrep; set => userGrep = value; }

        public void AddName(string name, bool isUsername)
        {
            if (isUsername)
            {
                var hasUsernameAlready = probablyUsernames.Where(p => p.MyName == name);
                if (hasUsernameAlready.Count() > 0)
                {

                }
                else
                {
                    //Not in list
                    probablyUsernames.Add(new Name() { MyName = name });
                }
            }
            else
            {
                var hasNameAlready = probablyNames.Where(p => p.MyName == name);
                if (hasNameAlready.Count() > 0)
                {

                }
                else
                {
                    //Not in list
                    probablyNames.Add(new Name() { MyName = name });
                }
            }
        }
    }

    public class GrepObject
    {
        string grepString;
        string innerFile;

        public string GrepString { get => grepString; set => grepString = value; }
        public string InnerFile { get => innerFile; set => innerFile = value; }
    }

    public class HiddenSheet
    {
        string hiddenSheetName;
        string hiddenSheetLocation;

        public string HiddenSheetName { get => hiddenSheetName; set => hiddenSheetName = value; }
        public string HiddenSheetLocation { get => hiddenSheetLocation; set => hiddenSheetLocation = value; }
    }
}
