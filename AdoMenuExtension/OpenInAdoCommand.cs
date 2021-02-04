using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace AdoMenuExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class OpenInAdoCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0f142ab0-b3fc-4268-bd96-33c76ddff4bf");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenInAdoCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private OpenInAdoCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static OpenInAdoCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new OpenInAdoCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var dte = GetDTE2();

            if (dte?.ActiveDocument?.Selection is TextSelection selection)
            {
                var path = dte.ActiveDocument.FullName;
                if (!File.Exists(path))
                {
                    System.Windows.Forms.MessageBox.Show($"Invalid path: {path}", "Error");
                    return;
                }

                var dir = Path.GetDirectoryName(path);

                var url = GetRepoUrl(dir);
                if (url == null)
                {
                    return;
                }
                var gitRoot = GetRepoRoot(dir);
                var pathInUrl = path.Substring(gitRoot.Length).Replace(@"\", "%2F");

                //var position = $"{selection.TopPoint.Line}:{selection.TopPoint.LineCharOffset}-{selection.BottomPoint.Line}:{selection.BottomPoint.LineCharOffset}  URL:{url}, Path:{pathInUrl}";
                //System.Windows.Forms.MessageBox.Show(path + " :: " + position);

                System.Diagnostics.Process.Start($"{url}?path={pathInUrl}&line={selection.TopPoint.Line}&lineEnd={selection.BottomPoint.Line}&lineStartColumn={selection.TopPoint.LineCharOffset}&lineEndColumn={selection.BottomPoint.LineCharOffset}&lineStyle=plain&_a=contents");
            }
        }

        private string GetRepoUrl(string dir)
        {
            var output = RunCommand(dir, "git remote get-url origin");
            if (!output.StartsWith("http"))
            {
                System.Windows.Forms.MessageBox.Show($"Git error: Unable to get remote url by 'git remote get-url origin'. Output: {output}", "Error");
                return null;
            }
            return output;
        }

        private string GetRepoRoot(string dir)
        {
            var output = RunCommand(dir, "git rev-parse --show-toplevel");
            return output;
        }

        private string RunCommand(string dir, string cmd)
        {
            var git = new System.Diagnostics.Process();

            git.StartInfo.RedirectStandardOutput = true;
            git.StartInfo.UseShellExecute = false;
            git.StartInfo.CreateNoWindow = true;
            git.StartInfo.WorkingDirectory = dir;

            git.StartInfo.FileName = @"powershell";
            git.StartInfo.Arguments = $@"-c ""{cmd}""";
            git.Start();
            var output = git.StandardOutput.ReadLine();
            git.WaitForExit();
            return output;
        }

        private static EnvDTE80.DTE2 GetDTE2()
        {
            return Package.GetGlobalService(typeof(DTE)) as EnvDTE80.DTE2;
        }
    }
}
