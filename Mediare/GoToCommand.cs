//------------------------------------------------------------------------------
// <copyright file="GoToCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Constants = EnvDTE.Constants;

namespace Mediare
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class GoToCommand
    {
        public class ProjectItemWrapper : IEquatable<ProjectItemWrapper>
        {
            public string Filename { get; set; }
            public string Project { get; set; }
            public string Path { get; set; }
            public ProjectItem ProjItem;

            private ProjectItemWrapper()
            {

            }

            public ProjectItemWrapper(ProjectItem inItem)
            {
                ProjItem = inItem;
                Path = inItem.FileNames[1];
                Filename = System.IO.Path.GetFileName(Path);
                Project = ProjItem.ContainingProject.Name;
            }

            public bool Equals(ProjectItemWrapper other)
            {
                return Path == other.Path;
            }
        }

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("109b9932-925c-4316-97a5-dbd9ccbb4995");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        public abstract class EnvDTEProjectKinds
        {
            public const string vsProjectKindSolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
        }

        static readonly Guid ProjectFileGuid = new Guid("6BB5F8EE-4483-11D3-8BCF-00C04F8EC28C");
        static readonly Guid ProjectFolderGuid = new Guid("6BB5F8EF-4483-11D3-8BCF-00C04F8EC28C");
        static readonly Guid ProjectVirtualFolderGuid = new Guid("6BB5F8F0-4483-11D3-8BCF-00C04F8EC28C");

        static readonly List<string> FileEndingsToSkip = new List<string>()
            {
                ".vcxproj.filters",
                ".vcxproj"
            };

        /// <summary>
        /// Initializes a new instance of the <see cref="GoToCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private GoToCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static GoToCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new GoToCommand(package);
        }

        public static DTE GetActiveIDE()
        {
            // Get an instance of currently running Visual Studio IDE.
            DTE dte2 = Package.GetGlobalService(typeof(DTE)) as DTE;
            return dte2;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "GoToCommand";

            Document doc = GetActiveIDE().ActiveDocument;

            if (doc == null)
            {
                return;
            }

            var suffix = string.Empty;

            if (!doc.Name.EndsWith(".cs"))
            {
                return;
            }

            var name = doc.Name.Split('.')[0];

            if (name.EndsWith("Command") || name.EndsWith("Query") || name.EndsWith("Request"))
            {
                suffix = "Handler";
            }
            else if (name.EndsWith("ViewModel") || name.EndsWith("DataRecord"))
            {
                suffix = "Mapper";
            }

            var targetName = name + suffix + ".cs";

            var projItems = new Dictionary<string, ProjectItemWrapper>(StringComparer.Ordinal);
            foreach (var proj in GetProjects())
            {
                foreach (var item in EnumerateProjectItems(proj.ProjectItems))
                {
                    if (item.Filename == targetName)
                    {
                        if (!projItems.ContainsKey(item.Path))
                        {
                            projItems.Add(item.Path, item);
                        }
                    }
                }
            }

            //var wnd = new ListFiles(projItems.Values);
            //wnd.Owner = HwndSource.FromHwnd(new IntPtr(GetActiveIDE().MainWindow.HWnd)).RootVisual as System.Windows.Window;
            //wnd.Width = wnd.Owner.Width / 2;
            //wnd.Height = wnd.Owner.Height / 3;
            //wnd.ShowDialog();

            if (projItems.Count == 1)
            {
                var dte = GetActiveIDE();
                dte.ItemOperations.OpenFile(projItems.First().Key,
                    Constants.vsViewKindTextView);
            }
            else
            {
                this.ShowMessage(targetName + "\n" + $"Number of files: {projItems.Count}");
            }

            //dte.OpenFile(EnvDTE.Constants.vsViewKindTextView, "C:\\Dev\\TheCodeManual\\StringReplacer\\StringReplacer.cs");

            
        }

        public void ShowMessage(string message)
        {
            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                "Mediare",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }

        public static IList<Project> GetProjects()
        {
            Projects projects = GetActiveIDE().Solution.Projects;
            List<Project> list = new List<Project>();
            var item = projects.GetEnumerator();
            while (item.MoveNext())
            {
                var project = item.Current as Project;
                if (project == null)
                {
                    continue;
                }

                if (project.Kind == EnvDTEProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(project));
                }

                list.Add(project);
            }

            return list;
        }

        private static IEnumerable<Project> GetSolutionFolderProjects(Project solutionFolder)
        {
            List<Project> list = new List<Project>();
            for (var i = 1; i <= solutionFolder.ProjectItems.Count; i++)
            {
                var subProject = solutionFolder.ProjectItems.Item(i).SubProject;
                if (subProject == null)
                {
                    continue;
                }

                // If this is another solution folder, do a recursive call, otherwise add
                if (subProject.Kind == EnvDTEProjectKinds.vsProjectKindSolutionFolder)
                {
                    list.AddRange(GetSolutionFolderProjects(subProject));
                }

                list.Add(subProject);
            }
            return list;
        }

        private IEnumerable<ProjectItemWrapper> EnumerateProjectItems(ProjectItems items)
        {
            if (items != null)
            {
                for (int i = 1; i <= items.Count; i++)
                {
                    var itm = items.Item(i);

                    foreach (var res in EnumerateProjectItems(itm.ProjectItems))
                    {
                        yield return res;
                    }

                    try
                    {
                        var itmGuid = Guid.Parse(itm.Kind);
                        if (itmGuid.Equals(ProjectVirtualFolderGuid)
                            || itmGuid.Equals(ProjectFolderGuid))
                        {
                            continue;
                        }
                    }
                    catch (Exception)
                    {
                        // itm.Kind may throw an exception with certain node types like WixExtension (COMException)
                    }

                    for (short j = 0; itm != null && j < itm.FileCount; j++)
                    {
                        bool bSkip = false;
                        foreach (var ending in FileEndingsToSkip)
                        {
                            if (itm.FileNames[1] == null || itm.FileNames[1].EndsWith(ending))
                            {
                                bSkip = true;
                                break;
                            }
                        }

                        if (!bSkip)
                        {
                            yield return new ProjectItemWrapper(itm);
                        }
                    }
                }
            }
        }
    }
}
