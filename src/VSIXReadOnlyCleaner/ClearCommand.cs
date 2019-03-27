using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;
using System.Linq;
using EnvDTE;
using System.IO;
using Microsoft.VisualStudio.Threading;

namespace VSIXReadOnlyCleaner
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ClearCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("5d723dd8-ab65-4e48-b9ac-97877897b219");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClearCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ClearCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static ClearCommand Instance
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
            // Switch to the main thread - the call to AddCommand in ClearCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ClearCommand(package, commandService);
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

            DTE dte = ThreadHelper.JoinableTaskFactory.Run<object>(new Func<Task<object>>(() => this.ServiceProvider.GetServiceAsync(typeof(DTE)))) as DTE;
            if (dte == null) return;

            // 获取解决方案的路径
            string solutionPath = Path.GetDirectoryName(dte.Solution.FileName);
            // 获取项目的所有文件
            string[] files = Directory.GetFiles(solutionPath, "*.*", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                try
                {
                    // 设置为没有任何特殊属性的标准文件
                    File.SetAttributes(file, FileAttributes.Normal);
                }
                catch
                {
                    WriteLog($"处理失败。{file}");
                }
            }

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                "处理完成",
                "",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        /// <summary>
        /// 写入日志到磁盘
        /// </summary>
        /// <param name="message"></param>
        private void WriteLog(string message)
        {
            /**
             * 运行 Visual Studio 中使用/log命令行开关在会话期间将 ActivityLog.xml 写入到磁盘。
             * 关闭 Visual Studio 后, 找到活动日志的子文件夹中的 Visual Studio 数据： %appdata%\Microsoft\VisualStudio\15.0\ActivityLog.xml。
             * **/

            ThreadHelper.ThrowIfNotOnUIThread();

            IVsActivityLog log = ThreadHelper.JoinableTaskFactory.Run<object>(new Func<Task<object>>(() => this.ServiceProvider.GetServiceAsync(typeof(SVsActivityLog)))) as IVsActivityLog;
            if (log == null) return;

            int hr = log.LogEntry((UInt32)__ACTIVITYLOG_ENTRYTYPE.ALE_INFORMATION,
                this.ToString(),
                message);
        }
    }
}
