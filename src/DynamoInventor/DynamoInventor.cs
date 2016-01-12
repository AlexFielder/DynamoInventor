using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Inventor;
using Dynamo.PackageManager;
using DynamoInventor.Properties;
using Dynamo.Utilities;
using InventorServices.Persistence;
using Dynamo.Interfaces;
using System.IO;

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
        enum Versions { ShapeManager = 220 }

        private Inventor.Application inventorApplication;
        private DynamoInventorAddinButton dynamoAddinButton;
        private UserInterfaceEvents userInterfaceEvents;
        UserInterfaceManager userInterfaceManager;
        RibbonPanel assemblyRibbonPanel;
        RibbonPanel partRibbonPanel;
        RibbonPanel drawingRibbonPanel;
        
        public IPathManager DynamoPathManager { get { return pathManager; } }
        private readonly IPathManager pathManager;

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
        private static AssemblyHelper assemblyHelper;
        internal static string corePath;
        #endregion

        #region Public constructors
        public DynamoInventor()
        {
            Assembly.LoadFrom(@"C:\Program Files\Dynamo 0.9\DynamoCore.dll");
            // Even though this method has no dependencies on DynamoCore, resolution of DynamoCore is 
            // failing prior to the handler to AssemblyResolve event getting registered.  
            SubscribeAssemblyResolvingEvent();
        }
        #endregion

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {        
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
                        assemblyRibbonPanel = ribbonPanels.Add(ribbonPanelDisplayName, ribbonPanelInternalName, "{DB59D9A7-EE4C-434A-BB5A-F93E8866E872}", "", false);
                        CommandControls assemblyRibbonPanelCtrls = assemblyRibbonPanel.CommandControls;
                        CommandControl assemblyCmdBtnCmdCtrl = assemblyRibbonPanelCtrls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false); 

                        Inventor.Ribbon partRibbon = ribbons["Part"];
                        RibbonTabs partRibbonTabs = partRibbon.RibbonTabs;
                        RibbonTab modelRibbonTab = partRibbonTabs["id_AddInsTab"];
                        RibbonPanels partRibbonPanels = modelRibbonTab.RibbonPanels;
                        partRibbonPanel = partRibbonPanels.Add(ribbonPanelDisplayName, ribbonPanelInternalName, "{DB59D9A7-EE4C-434A-BB5A-F93E8866E872}", "", false);
                        CommandControls partRibbonPanelCtrls = partRibbonPanel.CommandControls;
                        CommandControl partCmdBtnCmdCtrl = partRibbonPanelCtrls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false);

                        Inventor.Ribbon drawingRibbon = ribbons["Drawing"];
                        RibbonTabs drawingRibbonTabs = drawingRibbon.RibbonTabs;
                        RibbonTab drawingRibbonTab = drawingRibbonTabs["id_AddInsTab"];
                        RibbonPanels drawingRibbonPanels = drawingRibbonTab.RibbonPanels;
                        drawingRibbonPanel = drawingRibbonPanels.Add(ribbonPanelDisplayName, ribbonPanelInternalName, "{DB59D9A7-EE4C-434A-BB5A-F93E8866E872}", "", false);
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

        void appEvents_OnDeactivateDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
        {
            handlingCode = HandlingCodeEnum.kEventNotHandled;
            if (beforeOrAfter == EventTimingEnum.kBefore)
            {
                //I think it is probably not necessary to register these events as the registration will happen on its own in InventorServices.
                //PersistenceManager.ResetOnDocumentDeactivate(); 
            }       
        }

        void appEvents_OnActivateDocument(_Document documentObject, EventTimingEnum beforeOrAfter, NameValueMap context, out HandlingCodeEnum handlingCode)
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
                                                     "{DB59D9A7-EE4C-434A-BB5A-F93E8866E872}", 
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
        #endregion

        public static void SetupDynamoPaths()
        {
            UpdateSystemPathForProcess();
            //this is getting the parent directory - in this case: "C:\Program Files\Dynamo 0.9"
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
            var assemblyDirectory = System.IO.Path.GetDirectoryName(assemblyLocation);
            var parentDirectory = Directory.GetParent(assemblyDirectory);
            corePath = parentDirectory.FullName;

            string assDir = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string moduleRootFolder = System.IO.Path.GetDirectoryName(assDir);

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
            var assemblyLocation = Assembly.GetExecutingAssembly().Location;
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
            var dynamoAsmPath = System.IO.Path.Combine(corePath, "DynamoShapeManager.dll");
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
            AppDomain.CurrentDomain.AssemblyResolve += assemblyHelper.ResolveAssembly;
        }
        
    }
}
