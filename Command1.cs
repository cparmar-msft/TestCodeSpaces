using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace SubsMenus
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command1
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int MySubsBenefitsId = 0x0100;
        public const int MySubsDocsId = 0x0101;
        public const int MySubsSupportId = 0x0102;
        public const int MySubsContactAdmin = 0x0103;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("8b7194cb-ff4b-4605-8fc6-4db2f8231115");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private static IVsUIShellOpenDocument shellOpenDocument;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command1"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command1(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var benefitsCmdId = new CommandID(CommandSet, MySubsBenefitsId);
            var benefitMenuItem = new MenuCommand(this.InvokeMySubsBenefits, benefitsCmdId);
            commandService.AddCommand(benefitMenuItem);

            var supportCmdId = new CommandID(CommandSet, MySubsSupportId);
            var supportMenuItem = new MenuCommand(this.InvokeMySubsSupport, supportCmdId);
            commandService.AddCommand(supportMenuItem);

            var docsCmdId = new CommandID(CommandSet, MySubsDocsId);
            var docsMenuItem = new MenuCommand(this.InvokeMySubsDocs, docsCmdId);
            commandService.AddCommand(docsMenuItem);

            var contactAdminCmdId = new CommandID(CommandSet, MySubsContactAdmin);
            var contactAdminMenuItem = new MenuCommand(this.InvokeMySubsContactAdmin, contactAdminCmdId);
            commandService.AddCommand(contactAdminMenuItem);

        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command1 Instance
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
            // Switch to the main thread - the call to AddCommand in Command1's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new Command1(package, commandService);

            shellOpenDocument = await package.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void InvokeMySubsBenefits(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            //var shellOpenDoc = this.ServiceProvider.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            System.Diagnostics.Debug.Assert(shellOpenDocument != null, "Could not get open doc service");
            if (shellOpenDocument != null)
            {
                if (!ErrorHandler.Succeeded(shellOpenDocument.OpenStandardPreviewer(
                    (uint)(__VSOSPFLAGS.OSP_LaunchSystemBrowser | __VSOSPFLAGS.OSP_LaunchSingleBrowser),
                    @"https://my.visualstudio.com/benefits",
                    VSPREVIEWRESOLUTION.PR_Default, 0)))
                    System.Diagnostics.Debug.WriteLine("Could not navigate to benefits site");
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void InvokeMySubsSupport(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //var shellOpenDoc = this.ServiceProvider.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            System.Diagnostics.Debug.Assert(shellOpenDocument != null, "Could not get open doc service");
            if (shellOpenDocument != null)
            {
                if (!ErrorHandler.Succeeded(shellOpenDocument.OpenStandardPreviewer(
                    (uint)(__VSOSPFLAGS.OSP_LaunchSystemBrowser | __VSOSPFLAGS.OSP_LaunchSingleBrowser),
                    @"https://my.visualstudio.com/gethelp",
                    VSPREVIEWRESOLUTION.PR_Default, 0)))
                    System.Diagnostics.Debug.WriteLine("Could not navigate to benefits site");
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void InvokeMySubsDocs(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //var shellOpenDoc = this.ServiceProvider.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            System.Diagnostics.Debug.Assert(shellOpenDocument != null, "Could not get open doc service");
            if (shellOpenDocument != null)
            {
                if (!ErrorHandler.Succeeded(shellOpenDocument.OpenStandardPreviewer(
                    (uint)(__VSOSPFLAGS.OSP_LaunchSystemBrowser | __VSOSPFLAGS.OSP_LaunchSingleBrowser),
                    @"https://docs.microsoft.com/en-us/visualstudio/subscriptions/",
                    VSPREVIEWRESOLUTION.PR_Default, 0)))
                    System.Diagnostics.Debug.WriteLine("Could not navigate to benefits site");
            }
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void InvokeMySubsContactAdmin(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            //var shellOpenDoc = this.ServiceProvider.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            System.Diagnostics.Debug.Assert(shellOpenDocument != null, "Could not get open doc service");
            if (shellOpenDocument != null)
            {
                if (!ErrorHandler.Succeeded(shellOpenDocument.OpenStandardPreviewer(
                    (uint)(__VSOSPFLAGS.OSP_LaunchSystemBrowser | __VSOSPFLAGS.OSP_LaunchSingleBrowser),
                    @"https://my.visualstudio.com/benefits",
                    VSPREVIEWRESOLUTION.PR_Default, 0)))
                    System.Diagnostics.Debug.WriteLine("Could not navigate to benefits site");
            }
        }
    }
}
