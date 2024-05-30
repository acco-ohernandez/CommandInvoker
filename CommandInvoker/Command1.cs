#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

#endregion

namespace CommandInvoker
{
    [Transaction(TransactionMode.Manual)]
    public class CommandInvoker : IExternalCommand
    {
        public const string __title__ = "Rehost Hanger";
        public const string __doc__ = "";
        public const string __author__ = "Justice Pitts";
        public const string __min_revit_ver__ = "2020";
        public const bool __beta__ = true;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //const string basePath = @"C:\pyRevit-Master\extensions\ACCO_VC.extension\ACCO_VC.tab\Hydronic Supports.panel\HangersStack.stack\Hanger Hosting.pulldown\";
            //const string dllPath = basePath + "bin\\RehostHanger.dll";

            const string basePath = @"C:\Users\ohernandez\AppData\Roaming\Autodesk\Revit\Addins\2021";

            const string dllPath = basePath + @"\BuildRibbonFromXML.dll";
            const string commandName = "RibbonBuilder";

            // Check if the DLL file exists before attempting to load it
            if (!File.Exists(dllPath))
            {
                message = $"Error: DLL not found at path {dllPath}";
                return Result.Failed;
            }

            try
            {
                var assemblyBytes = File.ReadAllBytes(dllPath);
                var objAssembly = Assembly.Load(assemblyBytes);
                var myIEnumerableType = GetTypesSafely(objAssembly);

                bool commandFound = false;
                foreach (var objType in myIEnumerableType)
                {
                    if (objType.IsClass && string.Equals(objType.Name, commandName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        var ibaseObject = Activator.CreateInstance(objType);
                        var arguments = new object[] { commandData, basePath + "bin", elements };
                        objType.InvokeMember("Execute", BindingFlags.Default | BindingFlags.InvokeMethod, null, ibaseObject, arguments);
                        commandFound = true;
                        break;
                    }
                }

                if (!commandFound)
                {
                    message = $"Error: Command '{commandName}' not found in assembly.";
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                message = $"Error executing command: {ex.Message}";
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private static IEnumerable<Type> GetTypesSafely(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(x => x != null);
            }
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommandInvoker";
            string buttonTitle = "Command Invoker";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
