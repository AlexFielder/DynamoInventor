using System;
using System.Drawing;
using Inventor;

namespace DynamoInventor
{
    /// <summary>Base class for a ribbon button definition with an OnExecute handler.</summary>
    internal abstract class Button
    {
        private readonly ButtonDefinition buttonDefinition;
        private readonly ButtonDefinitionSink_OnExecuteEventHandler onExecute;

        public ButtonDefinition ButtonDefinition => buttonDefinition;

        protected Inventor.Application InventorApplication { get; }

        protected Button(Inventor.Application inventorApplication,
                         string displayName,
                         string internalName,
                         CommandTypesEnum commandType,
                         string clientId,
                         string description,
                         string tooltip,
                         Icon standardIcon,
                         Icon largeIcon,
                         ButtonDisplayEnum buttonDisplayType)
        {
            InventorApplication = inventorApplication;
            try
            {
                var standardIconIPictureDisp = PictureDispConverter.ToIPictureDisp(standardIcon);
                var largeIconIPictureDisp = PictureDispConverter.ToIPictureDisp(largeIcon);
                buttonDefinition = inventorApplication.CommandManager.ControlDefinitions.AddButtonDefinition(
                    displayName, internalName, commandType, clientId, description, tooltip,
                    standardIconIPictureDisp, largeIconIPictureDisp, buttonDisplayType);

                buttonDefinition.Enabled = true;
                onExecute = new ButtonDefinitionSink_OnExecuteEventHandler(ButtonDefinition_OnExecute);
                buttonDefinition.OnExecute += onExecute;
            }
            catch (Exception e)
            {
                throw new ApplicationException(e.ToString());
            }
        }

        protected abstract void ButtonDefinition_OnExecute(NameValueMap context);
    }
}
