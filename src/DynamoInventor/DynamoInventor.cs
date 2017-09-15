using Dynamo.Interfaces;
using Dynamo.PackageManager;
using Dynamo.Utilities;
using DynamoInventor.Properties;
using Inventor;
using InventorServices.Persistence;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DynamoInventor
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement.
    /// </summary>
    [GuidAttribute("476f38a1-75f3-450b-a75a-6f030bf012a9")]
    public class DynamoInventor : Inventor.ApplicationAddInServer
    {
        #region Private fields

        private enum Versions
        { ShapeManager = 221 }

        private Inventor.Application inventorApplication;
        private DynamoInventorAddinButton dynamoAddinButton;
        private UserInterfaceEvents userInterfaceEvents;
        private UserInterfaceManager userInterfaceManager;
        private RibbonPanel assemblyRibbonPanel;
        private RibbonPanel partRibbonPanel;
        private RibbonPanel drawingRibbonPanel;

        //public IPathManager DynamoPathManager { get { return pathManager; } }
        //private readonly IPathManager pathManager;

        private string commandBarInternalName = "Dynamo:InventorDynamo:DynamoCommandBar";
        private string commandBarDisplayName = "Dynamo";
        private string ribbonPanelInternalName = "Dynamo:InventorDynamo:DynamoRibbonPanel";
        private string ribbonPanelDisplayName = "Dynamo";
        private string buttonInternalName = "Dynamo:InventorDynamo:DynamoButton";
        private string buttonDisplayName = "Dynamo";
        private string commandCategoryInternalName = "Dynamo:InventorDynamo:DynamoCommandCat";
        private string commandCategoryDisplayName = "Dynamo";

        private Inventor.UserInterfaceEventsSink_OnResetCommandBarsEventHandler UserInterfaceEventsSink_OnResetCommandBarsEventDelegate;
        private Inventor.UserInterfaceEventsSink_OnResetEnvironmentsEventHandler UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate;
        private Inventor.UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler UserInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;

        private Inventor.ApplicationEvents appEvents = null;
        private AppDomain currentDomain = AppDomain.CurrentDomain;
        private static AssemblyHelper assemblyHelper;
        internal static string corePath;

        #endregion Private fields

        #region Public constructors

        //public DynamoInventor()
        //{
        //    Debugger.Break();
        //    Assembly.LoadFrom(@"C:\Program Files\Dynamo 0.9\DynamoCore.dll");
        //    // Even though this method has no dependencies on DynamoCore, resolution of DynamoCore is
        //    // failing prior to the handler to AssemblyResolve event getting registered.
        //    SubscribeAssemblyResolvingEvent();
        //}

        #endregion Public constructors

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            Debugger.Break();
            SubscribeAssemblyResolvingEvent();
            //currentDomain.AssemblyResolve += new ResolveEventHandler(MyResolveEventHandler);
            try
            {
                SetupDynamoPaths();
                inventorApplication = addInSiteObject.Application;
                PersistenceManager.InventorApplication = inventorApplication;
                userInterfaceManager = inventorApplication.UserInterfaceManager;

                //initialize event delegates
                userInterfaceEvents = inventorApplication.UserInterfaceManager.UserInterfaceEvents;

                UserInterfaceEventsSink_OnResetCommandBarsEventDelegate = new UserInterfaceEventsSink_OnResetCommandBarsEventHandler(UserInterfaceEvents_OnResetCommandBars);
                userInterfaceEvents.OnResetCommandBars += UserInterfaceEventsSink_OnResetCommandBarsEventDelegate;

                UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate = new UserInterfaceEventsSink_OnResetEnvironmentsEventHandler(UserInterfaceEvents_OnResetEnvironments);
                userInterfaceEvents.OnResetEnvironments += UserInterfaceEventsSink_OnResetEnvironmentsEventDelegate;

                UserInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate = new UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler(UserInterfaceEvents_OnResetRibbonInterface);
                userInterfaceEvents.OnResetRibbonInterface += UserInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;

                appEvents = inventorApplication.ApplicationEvents;
                appEvents.OnActivateDocument += appEvents_OnActivateDocument;
                appEvents.OnDeactivateDocument += appEvents_OnDeactivateDocument;

                Icon dynamoIcon = Resources.logo_square_32x32;

                //retrieve the GUID for this class
                GuidAttribute addInCLSID;
                addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(this.GetType(), typeof(GuidAttribute));
                string addInCLSIDString;
                addInCLSIDString = "{" + addInCLSID.Value + "}";

                dynamoAddinButton = new DynamoInventorAddinButton(
                        buttonDisplayName, buttonInternalName, CommandTypesEnum.kShapeEditCmdType,
                        addInCLSIDString, "Initialize Dynamo.",
                        "Dynamo is a visual programming environment for Inventor.", dynamoIcon, dynamoIcon, ButtonDisplayEnum.kDisplayTextInLearningMode);

                CommandCategory assemblyUtilitiesCategory = inventorApplication.CommandManager.CommandCategories.Add(commandCategoryDisplayName, commandCategoryInternalName, addInCLSID);
                assemblyUtilitiesCategory.Add(dynamoAddinButton.ButtonDefinition);

                if (firstTime == true)
                {
                    InterfaceStyleEnum interfaceStyle;
                    interfaceStyle = userInterfaceManager.InterfaceStyle;

                    if (interfaceStyle == InterfaceStyleEnum.kClassicInterface)
                    {
                        CommandBar assemblyUtilityCommandBar;

                        assemblyUtilityCommandBar = userInterfaceManager.CommandBars.Add(commandBarDisplayName,
                                                                                         commandBarInternalName,
                                                                                         CommandBarTypeEnum.kRegularCommandBar,
                                                                                         addInCLSID);
                    }
                    else
                    {
                        Inventor.Ribbons ribbons = userInterfaceManager.Ribbons;
                        Inventor.Ribbon assemblyRibbon = ribbons["Assembly"];
                        RibbonTabs ribbonTabs = assemblyRibbon.RibbonTabs;
                        RibbonTab assemblyRibbonTab = ribbonTabs["id_AddInsTab"];
                        RibbonPanels ribbonPanels = assemblyRibbonTab.RibbonPanels;
                        assemblyRibbonPanel = ribbonPanels.Add(ribbonPanelDisplayName, ribbonPanelInternalName, Guid.NewGuid().ToString(), "", false);
                        CommandControls assemblyRibbonPanelCtrls = assemblyRibbonPanel.CommandControls;
                        CommandControl assemblyCmdBtnCmdCtrl = assemblyRibbonPanelCtrls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false);

                        Inventor.Ribbon partRibbon = ribbons["Part"];
                        RibbonTabs partRibbonTabs = partRibbon.RibbonTabs;
                        RibbonTab modelRibbonTab = partRibbonTabs["id_AddInsTab"];
                        RibbonPanels partRibbonPanels = modelRibbonTab.RibbonPanels;
                        partRibbonPanel = partRibbonPanels.Add(ribbonPanelDisplayName, ribbonPanelInternalName, Guid.NewGuid().ToString(), "", false);
                        CommandControls partRibbonPanelCtrls = partRibbonPanel.CommandControls;
                        CommandControl partCmdBtnCmdCtrl = partRibbonPanelCtrls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false);

                        Inventor.Ribbon drawingRibbon = ribbons["Drawing"];
                        RibbonTabs drawingRibbonTabs = drawingRibbon.RibbonTabs;
                        RibbonTab drawingRibbonTab = drawingRibbonTabs["id_AddInsTab"];
                        RibbonPanels drawingRibbonPanels = drawingRibbonTab.RibbonPanels;
                        drawingRibbonPanel = drawingRibbonPanels.Add(ribbonPanelDisplayName, ribbonPanelInternalName, Guid.NewGuid().ToString(), "", false);
                        CommandControls drawingRibbonPanelCtrls = drawingRibbonPanel.CommandControls;
                        CommandControl drawingCmdBtnCmdCtrl = drawingRibbonPanelCtrls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private Assembly MyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            //This handler is called only when the common language runtime tries to bind to the assembly and fails.

            //Retrieve the list of referenced assemblies in an array of AssemblyName.
            Assembly MyAssembly, objExecutingAssemblies;
            string strTempAssmbPath = "";

            objExecutingAssemblies = Assembly.GetExecutingAssembly();
            AssemblyName[] arrReferencedAssmbNames = objExecutingAssemblies.GetReferencedAssemblies();

            //Loop through the array of referenced assembly names.
            foreach (AssemblyName strAssmbName in arrReferencedAssmbNames)
            {
                //Check for the assembly names that have raised the "AssemblyResolve" event.
                if (strAssmbName.FullName.Substring(0, strAssmbName.FullName.IndexOf(",")) == args.Name.Substring(0, args.Name.IndexOf(",")))
                {
                    //Build the path of the assembly from where it has to be loaded.
                    strTempAssmbPath = "C:\\Program Files\\Dynamo 0.9\\" + args.Name.Substring(0, args.Name.IndexOf(",")) + ".dll";
                    break;
                }
            }
            //Load the assembly from the specified path.
            MyAssembly = Assembly.LoadFrom(strTempAssmbPath);

            //Return the loaded assembly.
            return MyAssembly;
        }

        private void appEvents_OnDeactivateDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (beforeOrAfter == EventTimingEnum.kBefore)
            {
                //I think it is probably not necessary to register these events as the registration will happen on its own in InventorServices.
                //PersistenceManager.ResetOnDocumentDeactivate();
            }
        }

        private void appEvents_OnActivateDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (beforeOrAfter == EventTimingEnum.kAfter)
            {
                //PersistenceManager.ResetOnDocumentActivate(documentObject);
            }
        }

        public void Deactivate()
        {
            // TODO Dispose in InventorServices
            inventorApplication = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void UserInterfaceEvents_OnResetCommandBars(ObjectsEnumerator commandBars, NameValueMap context)
        {
            try
            {
                CommandBar commandBar;
                for (int commandBarCt = 1; commandBarCt <= commandBars.Count; commandBarCt++)
                {
                    commandBar = (Inventor.CommandBar)commandBars[commandBarCt];
                    if (commandBar.InternalName == commandBarInternalName)
                    {
                        commandBar.Controls.AddButton(dynamoAddinButton.ButtonDefinition, 0);
                        return;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void UserInterfaceEvents_OnResetEnvironments(ObjectsEnumerator environments, NameValueMap context)
        {
            //TODO: Fix this
            try
            {
                Inventor.Environment environment;
                for (int i = 1; i <= environments.Count; i++)
                {
                    environment = (Inventor.Environment)environments[i];
                    if (environment.InternalName == "AMxAssemblyEnvironment")
                    {
                        environment.PanelBar.CommandBarList.Add(inventorApplication.UserInterfaceManager.CommandBars[commandBarInternalName]);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        private void UserInterfaceEvents_OnResetRibbonInterface(NameValueMap context)
        {
            //TODO: Fix this
            try
            {
                //get the ribbon associated with part document
                Inventor.Ribbons ribbons = userInterfaceManager.Ribbons;
                Inventor.Ribbon assemblyRibbon = ribbons["Assembly"];

                //get the tabls associated with part ribbon
                RibbonTabs ribbonTabs = assemblyRibbon.RibbonTabs;
                RibbonTab assemblyRibbonTab = ribbonTabs["id_Assembly"];

                //create a new panel with the tab
                RibbonPanels ribbonPanels = assemblyRibbonTab.RibbonPanels;
                assemblyRibbonPanel = ribbonPanels.Add(ribbonPanelDisplayName,
                                                     ribbonPanelInternalName,
                                                     Guid.NewGuid().ToString(),
                                                     "",
                                                     false);

                CommandControls assemblyRibbonPanelCtrls = assemblyRibbonPanel.CommandControls;
                CommandControl copyUtilCmdBtnCmdCtrl = assemblyRibbonPanelCtrls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        /// <summary>
        /// Automation is part of the ApplicationAddInServer implementation.
        /// We don't need this for now.
        /// </summary>
        public object Automation
        {
            get { return null; }
        }

        /// <summary>
        /// Automation is part of the ApplicationAddInServer implementation.
        /// </summary>
        public void ExecuteCommand(int CommandID)
        {
        }

        #endregion ApplicationAddInServer Members

        public static void SetupDynamoPaths()
        {
            UpdateSystemPathForProcess();
            //this is getting the parent directory - in this case: "C:\Program Files\Dynamo 0.9"
            //var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var assemblyLocation = @"C:\Program Files\Dynamo\Dynamo Core\1.3\DynamoShapeManager.dll";
            var assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);
            var parentDirectory = Directory.GetParent(assemblyDirectory);
            corePath = assemblyDirectory; // parentDirectory.FullName;

            string assDir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string moduleRootFolder = assemblyDirectory;
            //string moduleRootFolder = System.IO.Path.GetDirectoryName(assDir);

            var resolutionPaths = new[]
            {
                System.IO.Path.Combine(moduleRootFolder, "nodes")
            };

            assemblyHelper = new AssemblyHelper(moduleRootFolder, resolutionPaths);
            AppDomain.CurrentDomain.AssemblyResolve += assemblyHelper.ResolveAssembly;
            // Add the Inventor_20xx folder for assembly resolution
            //DynamoPathManager.Instance.AddResolutionPath(assDir);
            //DynamoPathManager.AddResolutionPath(assDir);

            // Setup the core paths
            // TODO currently DynamoInventor's output path is set to the same one as DynamoCore.
            // When DynamoInventor is in the Inventor_20xx folder, the application is somehow
            // trying to resolve the Dynamo dependencies prior to registering a handler to the
            // AssemblyResolve event, so for now DynamoInventor is in Dynamo's debug path and the other
            // dlls for this project are in the Inventor_20xx folder below.
            // This doesn't hurt anything for now, but once 2016 comes out it will be a problem.
            //DynamoPathManager.Instance.InitializeCore(System.IO.Path.GetFullPath(assDir + @"\.."));
            //DynamoPathManager.Instance.InitializeCore(@"C:\Program Files\Dynamo 0.9");

            // Add Revit-specific paths for loading.
            ///DynamoPathManager.Instance.AddPreloadLibrary(System.IO.Path.Combine(assDir, "Inventor_2015\\InventorLibrary.dll"));

            // TODO: Fix this for versioning
            //DynamoPathManager.Instance.SetLibGPath("219");
        }

        /// <summary>
        /// Add the main exec path to the system PATH
        /// This is required to pickup certain dlls.
        /// </summary>
        private static void UpdateSystemPathForProcess()
        {
            var assemblyLocation = @"C:\Program Files\Dynamo\Dynamo Core\1.3\DynamoShapeManager.dll";
            //var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);
            var parentDirectory = Directory.GetParent(assemblyDirectory);
            var corePath = parentDirectory.FullName;

            var path =
                    System.Environment.GetEnvironmentVariable(
                        "Path",
                        EnvironmentVariableTarget.Process) + ";" + corePath;
            System.Environment.SetEnvironmentVariable("Path", path, EnvironmentVariableTarget.Process);
        }

        public static string GetGeometryFactoryPath(string corePath)
        {
            var dynamoAsmPath = @"C:\Program Files\Dynamo\Dynamo Core\1.3\DynamoShapeManager.dll";
            //var dynamoAsmPath = System.IO.Path.Combine(corePath, "DynamoShapeManager.dll");
            var assembly = Assembly.LoadFrom(dynamoAsmPath);
            if (assembly == null)
                throw new FileNotFoundException("File not found", dynamoAsmPath);

            var utilities = assembly.GetType("DynamoShapeManager.Utilities");
            var getGeometryFactoryPath = utilities.GetMethod("GetGeometryFactoryPath");

            return (getGeometryFactoryPath.Invoke(null,
                new object[] { corePath, Versions.ShapeManager }) as string);
        }

        private void SubscribeAssemblyResolvingEvent()
        {
            AppDomain.CurrentDomain.AssemblyResolve += ResolveAssembly;
        }

        /// <summary>
        /// Handler to the ApplicationDomain's AssemblyResolve event.
        /// If an assembly's location cannot be resolved, an exception is
        /// thrown. Failure to resolve an assembly will leave Dynamo in 
        /// a bad state, so we should throw an exception here which gets caught 
        /// by our unhandled exception handler and presents the crash dialogue.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static Assembly ResolveAssembly(object sender, ResolveEventArgs args)
        {
            var assemblyPath = string.Empty;
            var assemblyName = new AssemblyName(args.Name).Name + ".dll";

            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);

                // Try "Dynamo 0.x\Revit_20xx" folder first...
                assemblyPath = System.IO.Path.Combine(assemblyDirectory, assemblyName);
                if (!System.IO.File.Exists(assemblyPath))
                {
                    // If assembly cannot be found, try in "Dynamo 0.x" folder.
                    var parentDirectory = Directory.GetParent(assemblyDirectory);
                    assemblyPath = System.IO.Path.Combine(parentDirectory.FullName, assemblyName);
                }

                return (System.IO.File.Exists(assemblyPath) ? Assembly.LoadFrom(assemblyPath) : null);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("The location of the assembly, {0} could not be resolved for loading.", assemblyPath), ex);
            }
        }
    }
}