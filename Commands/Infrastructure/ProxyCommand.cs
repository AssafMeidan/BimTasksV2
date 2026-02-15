using System;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BimTasksV2.Commands.Infrastructure
{
    /// <summary>
    /// Abstract base class for all proxy commands.
    /// Runs in the default ALC and must NOT reference Prism/Unity/Serilog.
    /// Loads the real handler from the isolated BimTasksLoadContext via reflection.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public abstract class ProxyCommand : IExternalCommand
    {
        /// <summary>
        /// The simple class name of the handler (e.g., "TestInfrastructureRunner").
        /// Must exist in BimTasksV2.Commands.Handlers namespace.
        /// </summary>
        protected abstract string HandlerTypeName { get; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var loadContext = BimTasksApp.LoadContext;
                if (loadContext == null)
                {
                    TaskDialog.Show("BimTasksV2", "Plugin not initialized. Please restart Revit.");
                    return Result.Failed;
                }

                string assemblyPath = Assembly.GetExecutingAssembly().Location;
                var isolatedAssembly = loadContext.LoadFromAssemblyPath(assemblyPath);

                string fullTypeName = $"BimTasksV2.Commands.Handlers.{HandlerTypeName}";
                var handlerType = isolatedAssembly.GetType(fullTypeName);
                if (handlerType == null)
                {
                    TaskDialog.Show("BimTasksV2", $"Handler not found: {fullTypeName}");
                    return Result.Failed;
                }

                var handler = Activator.CreateInstance(handlerType);
                var executeMethod = handlerType.GetMethod("Execute", BindingFlags.Public | BindingFlags.Instance);
                if (executeMethod == null)
                {
                    TaskDialog.Show("BimTasksV2", $"Execute method not found on {fullTypeName}");
                    return Result.Failed;
                }

                executeMethod.Invoke(handler, new object[] { commandData.Application });
                return Result.Succeeded;
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                message = ex.InnerException.Message;
                TaskDialog.Show("BimTasksV2 Error", $"{HandlerTypeName} failed:\n{ex.InnerException.Message}");
                return Result.Failed;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("BimTasksV2 Error", $"{HandlerTypeName} failed:\n{ex.Message}");
                return Result.Failed;
            }
        }
    }
}
