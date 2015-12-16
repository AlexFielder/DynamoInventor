using Dynamo.Controls;
using Dynamo.ViewModels;
using DynamoUnits;
using Inventor;
using InventorServices.Persistence;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows;
using System.Windows.Interop;

namespace DynamoInventor
{
    internal class DynamoInventorAddinButton : Button
    {
        private static bool isRunning = false;
        public static double? dynamoViewX = null;
        public static double? dynamoViewY = null;
        public static double? dynamoViewWidth = null;
        public static double? dynamoViewHeight = null;
        private bool handledCrash = false;
        //private ButtonDefinition m_buttonDefinition;
        //private ButtonDefinitionSink_OnExecuteEventHandler ButtonDefinition_OnExecuteEventDelegate;

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

		public DynamoInventorAddinButton(string displayName, 
                                         string internalName, 
                                         CommandTypesEnum commandType, 
                                         string clientId, 
                                         string description, 
                                         string tooltip, 
                                         ButtonDisplayEnum buttonDisplayType)
			: base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
		{		
        }

		override protected void ButtonDefinition_OnExecute(NameValueMap context)
		{
			try
			{
                if (isRunning == false)
                {
                    //Start Dynamo!  
                    //DynamoInventor.SetupDynamoPaths();
                    string inventorContext = "Inventor " + PersistenceManager.InventorApplication.SoftwareVersion.DisplayVersion;

                    //Setup base units.  Need to double check what to do.  The ui default for me is inches, but API always must take cm.
                    BaseUnit.HostApplicationInternalAreaUnit = AreaUnit.SquareCentimeter;
                    BaseUnit.HostApplicationInternalLengthUnit = LengthUnit.Centimeter;
                    BaseUnit.HostApplicationInternalVolumeUnit = VolumeUnit.CubicCentimeter;

                    //Setup DocumentManager...this is all taken care of on its own.  Reference to active application will happen
                    //when first call to binder.InventorApplication happens

                    Dynamo.Models.DynamoModel inventorDynamoModel = Dynamo.Models.DynamoModel.Start();

                    DynamoViewModel dynamoViewModel = DynamoViewModel.Start(
                    new DynamoViewModel.StartConfiguration()
                    {
                        DynamoModel = inventorDynamoModel
                    });


                    IntPtr mwHandle = Process.GetCurrentProcess().MainWindowHandle;
                    var dynamoView = new DynamoView(dynamoViewModel);
                    new WindowInteropHelper(dynamoView).Owner = mwHandle;

                    handledCrash = false;
                    dynamoView.Show();
                    isRunning = true;
                }

                else if (isRunning == true)
                {
                    System.Windows.Forms.MessageBox.Show("Dynamo is already running.");
                }

                else
                {
                    System.Windows.Forms.MessageBox.Show("Something terrible happened.");
                }		
			}

			catch(Exception e)
			{
                System.Windows.Forms.MessageBox.Show(e.ToString());
			}
		}
	}
}
