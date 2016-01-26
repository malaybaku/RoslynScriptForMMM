using System;
using System.Linq;
using System.IO;
using System.Windows.Forms;

using MikuMikuPlugin;

using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;

namespace RoslynScriptForMMM
{
    public class ScriptRunner : ICommandPlugin
    {
        public Scene Scene { get; set; }
        public IWin32Window ApplicationForm { get; set; }
        public System.Drawing.Image Image => null;
        public System.Drawing.Image SmallImage => null;

        public string Text => "C# Script";
        public string EnglishText => "C# Script";
        public string Description => "Run C# Script";

        public Guid GUID => new Guid("9a222735-51c3-4179-9e83-3928ed735825");

        public void Dispose()
        {
            //不要: 別にスクリプトも保持するわけじゃないし。
        }

        public async void Run(CommandArgs e)
        {
            if(!File.Exists(filePath))
            {
                MessageBox.Show(
                    $"Script File was not found. Please check if file exists: {Path.GetFullPath(filePath)}"
                    );
                return;
            }

            var script = CSharpScript.Create(
                File.ReadAllText(filePath),
                _scriptOption,
                typeof(ApiHolder)
                );

            if(!CheckScriptValidity(script))
            {
                MessageBox.Show(
                    $"Detected compile error. Please check log file: {Path.GetFullPath(errorLogPath)}"
                    );
                return;
            }


            var mmm = new ApiHolder(new ScriptPluginApi(Scene, ApplicationForm));
            await script.RunAsync(mmm);
        }

        private bool CheckScriptValidity(Script script)
        {
            var diagnostics = script.Compile();
            if(!diagnostics.Any(d => d.WarningLevel == 0))
            {
                return true;
            }

            using (var sw = new StreamWriter(errorLogPath, true))
            {
                sw.WriteLine($"{DateTime.Now.ToLongDateString()}, {DateTime.Now.ToLongTimeString()}");
                foreach (var d in diagnostics)
                {
                    sw.WriteLine(d.ToString());
                }
                sw.WriteLine();
            }
            return false;
        }

        private readonly ScriptOptions _scriptOption = ScriptOptions
            .Default
            .WithFilePath(filePath)
            .WithReferences(
                typeof(Scene).Assembly,
                typeof(MessageBox).Assembly,
                typeof(System.Drawing.Image).Assembly
                );

        private static readonly string filePath = "Script\\MMMPlugin.csx";
        private static readonly string errorLogPath = "Script\\CompileError.log";

    }

    public class ApiHolder
    {
        public ApiHolder(ScriptPluginApi api)
        {
            MMM = api;
        }

        public ScriptPluginApi MMM { get; }
    }

    public class ScriptPluginApi
    {
        public ScriptPluginApi(Scene scene, IWin32Window applicationForm)
        {
            Scene = scene;
            ApplicationForm = applicationForm;
        }

        public Scene Scene { get; }
        public IWin32Window ApplicationForm { get; }
    }

}
