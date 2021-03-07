﻿using System;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace EntityieldsAnalyser
{
    // Do not forget to update version number and author (company attribute) in AssemblyInfo.cs class
    // To generate Base64 string for Images below, you can use https://www.base64-image.de/
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Entity Fields Analysis"),
        //ExportMetadata("Description", "A tool to get an overview about the fields in your entity, and also a field calculator you can use to get an idea if you entity is able to support the fields you want to create."),
        ExportMetadata("Description", "A tool to get an overview about the fields in your entity, and an entity field calculator."),
        // Please specify the base64 content of a 32x32 pixels image
        ExportMetadata("SmallImageBase64", "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"),
        // Please specify the base64 content of a 80x80 pixels image
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "Lavender"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray")]
    public class MyPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MyPluginControl();
        }

        /// <summary>
        /// Constructor 
        /// </summary>
        public MyPlugin()
        {
            // If you have external assemblies that you need to load, uncomment the following to 
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        /// <summary>
        /// Event fired by CLR when an assembly reference fails to load
        /// Assumes that related assemblies will be loaded from a subfolder named the same as the Plugin
        /// For example, a folder named Sample.XrmToolBox.MyPlugin 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}