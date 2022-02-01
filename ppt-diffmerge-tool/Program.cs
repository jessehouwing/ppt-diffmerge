using ManyConsole;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using PowerPointApplication = Microsoft.Office.Interop.PowerPoint.Application;
using Microsoft.Office.Interop.PowerPoint;

namespace ppt_diffmerge_tool
{
    internal class MergeCommand : ConsoleCommand
    {
        readonly ManualResetEvent handle = new ManualResetEvent(false);

        public string @Base
        {
            get;
            set;
        }

        public string Local
        {
            get; set;
        }

        public string Remote
        {
            get; set;
        }

        public string Output
        {
            get; set;
        }

        public MergeCommand()
        {
            IsCommand("merge", "Merges 2 powerpoint files.");
            HasRequiredOption<string>("l|local=", "Local file", f => Local = f);
            HasRequiredOption<string>("r|remote=", "Remote file", f => Remote = f);
            HasOption<string>("o|output=", "(Optional) Output file", f => Output = f);
            HasOption<string>("b|base=", "(Optional) Base file", f => Base = f);
        }

        public override int Run(string[] remainingArguments)
        {
            PowerPointApplication app = null;
            Presentation presentation = null;

            try
            {
                app = new PowerPointApplication();
                app.PresentationCloseFinal += App_PresentationClose;

                if (!string.IsNullOrWhiteSpace(Output))
                {
                    File.Copy(Local, Output, true);
                    Local = Output;
                }

                presentation = app.Presentations.Open(Local);

                if (string.IsNullOrWhiteSpace(Base))
                {
                    presentation.Merge(Remote);
                }
                else
                {
                    presentation.MergeWithBaseline(Remote, Base);
                }

                handle.WaitOne();
            }
            finally
            {
                Marshal.ReleaseComObject(presentation);
                Marshal.ReleaseComObject(app);
            }
            return 0;
        }

        private void App_PresentationClose(Presentation presentation)
        {
            handle.Set();
        }
    }

    internal class DiffCommand : ConsoleCommand
    {
        readonly ManualResetEvent handle = new ManualResetEvent(false);

        public string @Base
        {
            get;
            set;
        }

        public string Local
        {
            get; set;
        }

        public string Remote
        {
            get; set;
        }

        public DiffCommand()
        {
            IsCommand("diff", "Diffs 2 powerpoint files.");
            HasRequiredOption<string>("l|local=", "Local file", f => Local = f);
            HasRequiredOption<string>("r|remote=", "Remote file", f => Remote = f);
            HasOption<string>("b|base=", "(Optional) Base file", f => Base = f);
        }

        public override int Run(string[] remainingArguments)
        {
            PowerPointApplication app = null;
            Presentation presentation = null;

            try
            {
                app = new PowerPointApplication();
                app.PresentationCloseFinal += App_PresentationClose;

                presentation = app.Presentations.Open(Local, Microsoft.Office.Core.MsoTriState.msoTrue);

                if (string.IsNullOrWhiteSpace(Base))
                {
                    presentation.Merge(Remote);
                }
                else
                {
                    presentation.MergeWithBaseline(Remote, Base);
                }

                handle.WaitOne();
            }
            finally
            {
                Marshal.ReleaseComObject(presentation);
                Marshal.ReleaseComObject(app);
            }
            return 0;
        }

        private void App_PresentationClose(Presentation presentation)
        {
            handle.Set();
        }
    }

    internal static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            int result;
            try
            {
                ConsoleCommand[] commands = { new DiffCommand(), new MergeCommand() };
                result = ConsoleCommandDispatcher.DispatchCommand(commands, args, Console.Out);
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
