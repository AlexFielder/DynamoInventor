using System;
using System.Drawing;
using Inventor;

namespace DynamoInventor
{
    /// <summary>
    /// Ribbon button that opens (or brings forward) the out-of-process Dynamo host.
    /// </summary>
    internal class DynamoInventorAddinButton : Button
    {
        public DynamoInventorAddinButton(Inventor.Application inventorApplication,
                                         string displayName,
                                         string internalName,
                                         CommandTypesEnum commandType,
                                         string clientId,
                                         string description,
                                         string tooltip,
                                         Icon standardIcon,
                                         Icon largeIcon,
                                         ButtonDisplayEnum buttonDisplayType)
            : base(inventorApplication, displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            try
            {
                DynamoAppLauncher.OpenOrFocus(InventorApplication);
            }
            catch (Exception e)
            {
                AddinLog.Log("Failed to open Dynamo: " + e);
                System.Windows.Forms.MessageBox.Show(
                    e.Message + System.Environment.NewLine + System.Environment.NewLine + "Details: " + AddinLog.LogPath,
                    "Dynamo for Inventor",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
