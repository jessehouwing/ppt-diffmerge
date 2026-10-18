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
                var localFile = new Argument<FileInfo>("local") { Description = "path to local file" };
                localFile.AcceptLegalFilePathsOnly();
                localFile.AcceptExistingOnly();

                var remoteFile = new Argument<FileInfo>("remote") { Description = "path to remote file" };
                remoteFile.AcceptLegalFilePathsOnly();
                remoteFile.AcceptExistingOnly();

                var baseFile = new Argument<FileInfo>("base")
                {
                    Description = "(Optional) path to base file",
                    Arity = ArgumentArity.ZeroOrOne,
                    DefaultValueFactory = _ => null
                };
                baseFile.AcceptLegalFilePathsOnly();
                baseFile.AcceptExistingOnly();

                var resultFile = new Argument<FileInfo>("result")
                {
                    Description = "(Optional) path to result file",
                    Arity = ArgumentArity.ZeroOrOne,
                    DefaultValueFactory = _ => null
                };
                resultFile.AcceptLegalFilePathsOnly();

                var rootCommand = new RootCommand("PowerPoint Diff/Merge tool");
                rootCommand.Arguments.Add(localFile);
                rootCommand.Arguments.Add(remoteFile);
                rootCommand.Arguments.Add(baseFile);
                rootCommand.Arguments.Add(resultFile);

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

                rootCommand.SetAction(parseResult => handle1(parseResult.GetValue(localFile), parseResult.GetValue(remoteFile), parseResult.GetValue(baseFile), parseResult.GetValue(resultFile)));

                // Parse the incoming args and invoke the handler
                return rootCommand.Parse(args).Invoke();
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
