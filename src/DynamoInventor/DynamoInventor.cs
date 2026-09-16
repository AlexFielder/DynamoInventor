using System;
using System.Drawing;
using System.Runtime.InteropServices;
using DynamoInventor.Properties;
using Inventor;
using InventorServices.Persistence;

namespace DynamoInventor
{
    /// <summary>
    /// The Inventor add-in entry point (ApplicationAddInServer). Registers the Dynamo ribbon button;
    /// Dynamo itself is only loaded when the button is first pressed (see <see cref="DynamoSession"/>).
    /// </summary>
    [ComVisible(true)]
    [Guid("476f38a1-75f3-450b-a75a-6f030bf012a9")]
    public class DynamoInventor : ApplicationAddInServer
    {
        private const string RibbonPanelInternalName = "Dynamo:InventorDynamo:DynamoRibbonPanel";
        private const string RibbonPanelDisplayName = "Dynamo";
        private const string ButtonInternalName = "Dynamo:InventorDynamo:DynamoButton";
        private const string ButtonDisplayName = "Dynamo";
        private const string CommandCategoryInternalName = "Dynamo:InventorDynamo:DynamoCommandCat";
        private const string CommandCategoryDisplayName = "Dynamo";

        private static readonly string[] RibbonsWithButton = { "Assembly", "Part", "Drawing", "ZeroDoc" };

        private Inventor.Application inventorApplication;
        private UserInterfaceManager userInterfaceManager;
        private UserInterfaceEvents userInterfaceEvents;
        private DynamoInventorAddinButton dynamoAddinButton;
        private string addInClsid;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                inventorApplication = addInSiteObject.Application;
                PersistenceManager.InventorApplication = inventorApplication;
                userInterfaceManager = inventorApplication.UserInterfaceManager;

                DynamoRuntime.Log("Activate: Inventor " + inventorApplication.SoftwareVersion.DisplayVersion +
                                  " build " + inventorApplication.SoftwareVersion.BuildIdentifier);

                var clsid = (GuidAttribute)System.Attribute.GetCustomAttribute(GetType(), typeof(GuidAttribute));
                addInClsid = "{" + clsid.Value + "}";

                Icon dynamoIcon = Resources.logo_square_32x32;
                dynamoAddinButton = new DynamoInventorAddinButton(
                    ButtonDisplayName, ButtonInternalName, CommandTypesEnum.kShapeEditCmdType,
                    addInClsid, "Open Dynamo.",
                    "Dynamo is a visual programming environment for Inventor.",
                    dynamoIcon, dynamoIcon, ButtonDisplayEnum.kDisplayTextInLearningMode);

                var category = inventorApplication.CommandManager.CommandCategories.Add(
                    CommandCategoryDisplayName, CommandCategoryInternalName, addInClsid);
                category.Add(dynamoAddinButton.ButtonDefinition);

                if (firstTime)
                {
                    AddRibbonPanels();
                }

                userInterfaceEvents = userInterfaceManager.UserInterfaceEvents;
                userInterfaceEvents.OnResetRibbonInterface += UserInterfaceEvents_OnResetRibbonInterface;
            }
            catch (Exception e)
            {
                DynamoRuntime.Log("Activate failed: " + e);
                System.Windows.Forms.MessageBox.Show(e.ToString(), "Dynamo for Inventor");
            }
        }

        public void Deactivate()
        {
            try
            {
                DynamoSession.Close();
            }
            catch (Exception e)
            {
                DynamoRuntime.Log("Deactivate: " + e);
            }

            if (userInterfaceEvents != null)
            {
                userInterfaceEvents.OnResetRibbonInterface -= UserInterfaceEvents_OnResetRibbonInterface;
                userInterfaceEvents = null;
            }
            dynamoAddinButton = null;
            userInterfaceManager = null;
            PersistenceManager.InventorApplication = null;
            inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
        }

        public object Automation => null;

        private void UserInterfaceEvents_OnResetRibbonInterface(NameValueMap context)
        {
            try
            {
                AddRibbonPanels();
            }
            catch (Exception e)
            {
                DynamoRuntime.Log("OnResetRibbonInterface: " + e);
            }
        }

        /// <summary>Puts a "Dynamo" panel with the button on the Add-Ins tab of each document ribbon.</summary>
        private void AddRibbonPanels()
        {
            foreach (var ribbonName in RibbonsWithButton)
            {
                try
                {
                    Ribbon ribbon = userInterfaceManager.Ribbons[ribbonName];
                    RibbonTab addInsTab = ribbon.RibbonTabs["id_AddInsTab"];
                    RibbonPanel panel = addInsTab.RibbonPanels.Add(
                        RibbonPanelDisplayName, RibbonPanelInternalName, addInClsid, "", false);
                    panel.CommandControls.AddButton(dynamoAddinButton.ButtonDefinition, true, true, "", false);
                }
                catch (Exception e)
                {
                    // Not every ribbon has an Add-Ins tab in every Inventor configuration; keep going.
                    DynamoRuntime.Log("Ribbon '" + ribbonName + "': " + e.Message);
                }
            }
        }
    }
}
