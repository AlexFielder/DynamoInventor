namespace DynamoInventor.Models
{
    //public class DynamoModel : DynamoModel
    //{
    //    //private DynamoModel.IStartConfiguration configuration;

    //    //public new static InventorDynamoModel Start()
    //    //{
    //    //    return InventorDynamoModel.Start(new DynamoModel.IStartConfiguration());
    //    //}

    //    //public static new InventorDynamoModel Start(DynamoModel.IStartConfiguration configuration)
    //    //{
    //    //    if (string.IsNullOrEmpty(configuration.Context))
    //    //        configuration.Context = Dynamo.Core.Context.NONE;
    //    //    if (string.IsNullOrEmpty(configuration.DynamoCorePath))
    //    //    {
    //    //        var asmLocation = Assembly.GetExecutingAssembly().Location;
    //    //        configuration.DynamoCorePath = Path.GetDirectoryName(asmLocation);
    //    //    }

    //    //    if (configuration.Preferences == null)
    //    //        configuration.Preferences = new PreferenceSettings();

    //    //    return new InventorDynamoModel(configuration);
    //    //    //return new InventorDynamoModel(configuration);
    //    //}

    //    //private InventorDynamoModel(DynamoModel.IStartConfiguration configuration) :
    //    //    base(configuration)
    //    //{
    //    //    string context = configuration.Context;
    //    //    IPreferences preferences = configuration.Preferences;
    //    //    string corePath = configuration.DynamoCorePath;
    //    //    bool isTestMode = configuration.StartInTestMode;
    //    //}

    //    //public InventorDynamoModel(DynamoModel.IStartConfiguration configuration)
    //    //{
    //    //    this.configuration = configuration;
    //    //}
    //}
}
