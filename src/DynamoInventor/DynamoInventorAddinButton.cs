using System;
using System.Drawing;
using Inventor;
using InventorServices.Persistence;

namespace DynamoInventor
{
    /// <summary>
    /// Ribbon button that opens (or brings forward) the Dynamo window.
    /// </summary>
    internal class DynamoInventorAddinButton : Button
    {
        public DynamoInventorAddinButton(string displayName,
                                         string internalName,
                                         CommandTypesEnum commandType,
                                         string clientId,
                                         string description,
                                         string tooltip,
                                         Icon standardIcon,
                                         Icon largeIcon,
                                         ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {
        }

        protected override void ButtonDefinition_OnExecute(NameValueMap context)
        {
            try
            {
                DynamoSession.Open(PersistenceManager.InventorApplication);
            }
            catch (Exception e)
            {
                DynamoRuntime.Log("Failed to open Dynamo: " + e);
                System.Windows.Forms.MessageBox.Show(
                    e.Message + System.Environment.NewLine + System.Environment.NewLine + "Details: " + DynamoRuntime.LogPath,
                    "Dynamo for Inventor",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
