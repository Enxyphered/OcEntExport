using Newtonsoft.Json;
using System.Configuration;
using System.Data;
using System.IO;
using System.Windows;

namespace OcEntExport;

public partial class App : Application
{
    public static AppSettings? GetAppSettings()
    {
        var path = "AppSettings.json";
        if (!File.Exists(path))
            return null;

        var json = File.ReadAllText(path);
        try
        {
            return JsonConvert.DeserializeObject<AppSettings>(json);
        }
        catch
        {
            return null;
        }
    }
}

public struct AppSettings
{
    public string Query { get; set; }
    public string HostServer { get; set; }
    public Dictionary<string, string> TypeMappings { get; set; }
}