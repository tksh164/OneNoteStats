﻿using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections.Generic;

namespace OneNoteStats
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionHandler;

            if (args.Length < 1)
            {
                showUsage();
                return;
            }
            string notebookNickName = args[0];
            string dumpListFilePath;
            if (args.Length >= 2)
            {
                dumpListFilePath = Path.GetFullPath(args[1]);
            }
            else
            {
                dumpListFilePath = Path.GetFullPath(notebookNickName + @".tsv");
            }

            if (!Directory.Exists(Path.GetDirectoryName(dumpListFilePath)))
            {
                Console.Error.WriteLine(@"Could not find the parent folder path ""{0}"".", dumpListFilePath);
                return;
            }
            if (Directory.Exists(dumpListFilePath))
            {
                Console.Error.WriteLine(@"""{0}"" is exists as directory.", dumpListFilePath);
                return;
            }
            if (File.Exists(dumpListFilePath))
            {
                Console.Error.WriteLine(@"""{0}"" is exists.", dumpListFilePath);
                return;
            }

            NotebookStats notebook = new NotebookStats(notebookNickName);

            Console.WriteLine(@"Notebook    : {0}", notebookNickName);
            Console.WriteLine(@"SectionGroup: {0}", notebook.SectionGroupCount);
            Console.WriteLine(@"Section     : {0}", notebook.SectionCount);
            Console.WriteLine(@"Page        : {0}", notebook.PageCount);
            Console.WriteLine(@"DumpListFile: {0}", dumpListFilePath);

            writeDumpListFile(notebook, dumpListFilePath);
        }

        private static void writeDumpListFile(NotebookStats notebook, string dumpListFilePath)
        {
            List<PageInfo> pageInfos = notebook.GetPageInfo();
            using (FileStream stream = new FileStream(dumpListFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                writeAsCsv(pageInfos, "\t", stream);
            }
        }

        private static void writeAsCsv(List<PageInfo> pageInfos, string separator, Stream stream)
        {
            using (StreamWriter writer = new StreamWriter(stream, Encoding.Unicode))
            {
                string[] headerFields = new string[] {
                        wrapDq(@"PageName"),
                        wrapDq(@"LastModifiedTime"),
                        wrapDq(@"DateTime"),
                        wrapDq(@"PageLevel"),
                        wrapDq(@"IsCurrentlyViewed"),
                        wrapDq(@"Location"),
                        wrapDq(@"Id"),
                    };
                writer.WriteLine(string.Join(separator, headerFields));

                foreach (PageInfo pageInfo in pageInfos)
                {
                    string[] fields = new string[] {
                        wrapDq(pageInfo.Name),
                        wrapDq(pageInfo.LastModifiedTime.ToString(@"yyyy/MM/dd hh:mm:ss")),
                        wrapDq(pageInfo.DateTime.ToString(@"yyyy/MM/dd hh:mm:ss")),
                        wrapDq(pageInfo.PageLevel.ToString()),
                        wrapDq(pageInfo.IsCurrentlyViewed),
                        wrapDq(pageInfo.LocationPath),
                        wrapDq(pageInfo.Id),
                    };

                    writer.WriteLine(string.Join(separator, fields));
                }

                writer.Flush();
            }
        }

        private static string wrapDq(string text)
        {
            return @"""" + text + @"""";
        }

        private static void showUsage()
        {
            Console.WriteLine(@"Usage: {0} <NotebookNickName> [DumpListFilePath]", Path.GetFileName(Assembly.GetEntryAssembly().Location));
        }

        private static void UnhandledExceptionHandler(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            int nestLevel = 0;

            while (ex != null)
            {
                Console.Error.WriteLine();
                Console.Error.WriteLine(@"**** EXCEPTION (Level:{0}) ****", nestLevel);
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(@"Exception: {0}", ex.GetType().FullName);
                Console.Error.WriteLine(@"**** STACK TRACE ****");
                Console.Error.WriteLine(ex.StackTrace);

                ex = ex.InnerException;
                nestLevel++;
            }

            Environment.Exit(-1);
        }
    }
}
