using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using Microsoft.Office.Interop.PowerPoint;
using System.CommandLine;

namespace ppt_diffmerge_tool
{ 
    internal static class Program
    {
        static readonly ManualResetEvent handle = new ManualResetEvent(false);

        [STAThread]
        static int Main(string[] args)
        {
            int result;
            try
            {
                // Create some options:
                var localFile = new Argument<FileInfo>("local", description: "path to local file");
                localFile.LegalFilePathsOnly();
                localFile.ExistingOnly();

                var remoteFile = new Argument<FileInfo>("remote", description: "path to remote file");
                remoteFile.LegalFilePathsOnly();
                remoteFile.ExistingOnly();

                var baseFile = new Argument<FileInfo>("base", description: "(Optional) path to base file");
                baseFile.LegalFilePathsOnly();
                baseFile.ExistingOnly();
                baseFile.SetDefaultValue(null);

                var resultFile = new Argument<FileInfo>("result", description: "(Optional) path to result file");
                resultFile.LegalFilePathsOnly();
                resultFile.SetDefaultValue(null);

                // Add the options to a root command:
                var rootCommand = new RootCommand();
                rootCommand.AddArgument(localFile);
                rootCommand.AddArgument(remoteFile);
                rootCommand.AddArgument(baseFile);
                rootCommand.AddArgument(resultFile);

                rootCommand.Description = "PowerPoint Diff/Merge tool";

                void handle1(FileInfo lf, FileInfo rf, FileInfo bf, FileInfo resf)
                {
                    PowerPointApplication app = null;
                    Presentation presentation = null;

                    try
                    {
                        app = new PowerPointApplication();
                        app.PresentationCloseFinal += (Presentation _) =>
                        {
                            handle.Set();
                        };

                        if (resf != null)
                        {
                            File.Copy(lf.FullName, resf.FullName, true);
                            lf = resf;
                        }

                        presentation = app.Presentations.Open(lf.FullName);

                        if (bf == null)
                        {
                            presentation.Merge(rf.FullName);
                        }
                        else
                        {
                            presentation.MergeWithBaseline(rf.FullName, bf.FullName);
                        }

                        handle.WaitOne();

                        // Ask to save?
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(presentation);
                        Marshal.ReleaseComObject(app);
                    }
                }

                rootCommand.SetHandler((Action<FileInfo, FileInfo, FileInfo, FileInfo>)handle1, localFile, remoteFile, baseFile, resultFile);

                // Parse the incoming args and invoke the handler
                return rootCommand.Invoke(args);
            }
            catch (Exception e)
            {
                result = -1;
                Console.Error.WriteLine(e.Message);
                if (e.InnerException != null)
                {
                    Console.Error.WriteLine(e.InnerException.Message);
                }
            }

            Environment.Exit(result);
            return result;
        }
    }
}
