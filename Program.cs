using CommandLine;
using System;

namespace ConsoleApp1
{
    class Program
    {
        public class Options
        {
            [Option('p', "path", Required = true, HelpText = "The path to the outlook profile")]
            public string ProfilePath { get; set; }

            [Option('s', "store", Required = false, HelpText = "The target outlook store id")]
            public string StoreId { get; set; }

            [Option('f', "folder", Required = false, HelpText = "The target outlook folder id")]
            public string FolderId { get; set; }

            [Option('t', "threads", Required = false, HelpText = "Create conversation threads (doesn't seem to work)")]
            public bool CreateThreads { get; set; }
        }

        static void Main(string[] args)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                Parser.Default.ParseArguments<Options>(args)
                  .WithParsed<Options>(o =>
                  {
                      var loader = OutlookFileLoader.Create(o.ProfilePath);
                      var query = OutlookSqliteQuery.Create(o.ProfilePath);
                      using (var manager = OutlookPstManager.Create(o.StoreId, o.FolderId))
                      {
                          var i = 0;
                          var total = 0;
                          foreach (var file in loader.EnumerateMessages())
                          {
                              total++;

                              var msgId = manager.GetMessageId(file);

                              query.TryFindMessage(msgId, out var threadId, out var isFirstMsg, out var folderPath);

                              manager.AddEmailToFolder(file.FullName, o.CreateThreads ? threadId : -1, isFirstMsg, folderPath);

                              if (++i >= 100)
                              {
                                  Console.Write(".");
                                  i = 0;
                              }
                          }

                          Console.WriteLine("Done");

                          Console.WriteLine($"Found {total:N0} out of {query.MessageCount:N0} ({(total / query.MessageCount * 100d):N2}%)");
                      }
                  });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
