using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using System.Text;

using Dynamo;
using Dynamo.Core;
using Dynamo.Models;
using Dynamo.Interfaces;
using Dynamo.Nodes;
using Dynamo.Utilities;
using Dynamo.ViewModels;
using Dynamo.Wpf.ViewModels.Watch3D;
using Dynamo.Services;
using Dynamo.Core.Threading;
using Dynamo.Extensions;
using Dynamo.UpdateManager;
using Greg;

namespace DynamoInventor.Models
{
    public class InventorDynamoModel : DynamoModel
    {
        public interface IInventorStartConfiguration: IStartConfiguration
        {

        }

        public struct InventorStartConfiguration : IInventorStartConfiguration
        {
            public IAuthProvider AuthProvider { get; set; }

            public string Context { get; set; }

            public string DynamoCorePath { get; set; }

            public IEnumerable<IExtension> Extensions { get; set; }
            
            public string GeometryFactoryPath { get; set; }
            
            public IPathResolver PathResolver { get; set; }

            public IPreferences Preferences { get; set; }

            public TaskProcessMode ProcessMode { get; set; }

            public ISchedulerThread SchedulerThread { get; set; }

            public bool StartInTestMode { get; set; }

            public IUpdateManager UpdateManager { get; set; }
        }
        public new static InventorDynamoModel Start()
        {
            return Start();
        }

        public new static InventorDynamoModel Start(IStartConfiguration configuration)
        {
            if (string.IsNullOrEmpty(configuration.Context))
                configuration.Context = Dynamo.Core.Context.REVIT_2015;
            if (string.IsNullOrEmpty(configuration.DynamoCorePath))
            {
                var asmLocation = Assembly.GetExecutingAssembly().Location;
                configuration.DynamoCorePath = Path.GetDirectoryName(asmLocation);
            }

            if (configuration.Preferences == null)
                configuration.Preferences = new PreferenceSettings();

            return new InventorDynamoModel(configuration);
        }
        private InventorDynamoModel(IStartConfiguration configuration) :
            base(configuration)
        {
            string context = configuration.Context;
            IPreferences preferences = configuration.Preferences;
            string corePath = configuration.DynamoCorePath;
            bool isTestMode = configuration.StartInTestMode;
        }
    }
}
