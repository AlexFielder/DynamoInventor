using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Dynamo.Models;
using Dynamo.Configuration;
using Dynamo.Interfaces;
using Dynamo.Extensions;
using Greg;
using Dynamo.Scheduler;
using Dynamo.Updates;
using System;
using Dynamo.Configuration;

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

            public string DynamoHostPath { get; set; }

            public IEnumerable<IExtension> Extensions { get; set; }
            
            public string GeometryFactoryPath { get; set; }
            
            public IPathResolver PathResolver { get; set; }

            public IPreferences Preferences { get; set; }

            public TaskProcessMode ProcessMode { get; set; }

            public ISchedulerThread SchedulerThread { get; set; }

            public bool StartInTestMode { get; set; }

            public IUpdateManager UpdateManager { get; set; }

            TaskProcessMode IStartConfiguration.ProcessMode { get; set; }

            ISchedulerThread IStartConfiguration.SchedulerThread { get; set; }

            IUpdateManager IStartConfiguration.UpdateManager { get; set; }
        }
        public new static InventorDynamoModel Start()
        {
            return Start();
        }

        public new static InventorDynamoModel Start(IStartConfiguration configuration)
        {
            if (string.IsNullOrEmpty(configuration.Context))
                configuration.Context = Dynamo.Configuration.Context.NONE;
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
