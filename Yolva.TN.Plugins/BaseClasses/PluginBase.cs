using Microsoft.Xrm.Sdk;
using System;

namespace Yolva.TN.Plugins.BaseClasses
{
    public abstract class PluginBase : IPlugin
    {
        /// <summary>
        /// Plugin unsecure configuration
        /// </summary>
        protected string UnsecureConfiguration { get; private set; }
        /// <summary>
        /// Plugin secure configuration
        /// </summary>
        protected string SecureConfiguration { get; private set; }

        public PluginBase(string unsecure, string secure)
        {
            UnsecureConfiguration = unsecure;
            SecureConfiguration = secure;
        }
        /// <summary>
        /// Main business logic happens here
        /// </summary>
        protected abstract void ExecuteBusinessLogic(PluginContext pluginContext);

        public void Execute(IServiceProvider serviceProvider)
        {
            PluginContext pluginContext = new PluginContext(serviceProvider);
            if (CheckPluginAttributes(pluginContext))
            {
                ExecuteBusinessLogic(pluginContext);
            }

        }

        private bool CheckPluginAttributes(PluginContext pluginContext)
        {
            bool isValid = true;
            Type currentType = this.GetType();
            Type attributeType = typeof(PluginValidationAttribute);
            if (Attribute.IsDefined(currentType, attributeType))
            {
                PluginValidationAttribute attributeValue = Attribute.GetCustomAttribute(currentType, attributeType) as PluginValidationAttribute;
                if (attributeValue != null)
                {
                    if (!string.IsNullOrEmpty(attributeValue.EntityName) && !string.Equals(attributeValue.EntityName, pluginContext.PrimaryEntityName, StringComparison.InvariantCultureIgnoreCase)
                        || (!string.IsNullOrEmpty(attributeValue.Message) && !string.Equals(attributeValue.Message, pluginContext.MessageName, StringComparison.InvariantCultureIgnoreCase))
                        || (attributeValue.Stage != null && attributeValue.Stage != pluginContext.Stage))
                    {
                        isValid = false;
                    }
                }
            }
            return isValid;
        }
    }

    public class PluginValidationAttribute : Attribute
    {
        public string EntityName { get; set; }
        public string Message { get; set; }
        public PluginStage? Stage { get; set; }

        public PluginValidationAttribute(string entityName, string message, PluginStage stage)
        {
            EntityName = entityName;
            Message = message;
            Stage = stage;
        }
    }
}
