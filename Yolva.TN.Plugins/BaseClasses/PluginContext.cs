using Microsoft.Xrm.Sdk;
using System;


namespace Yolva.TN.Plugins.BaseClasses
{
    public class PluginContext
    {
        private IServiceProvider serviceProvider;

        private IPluginExecutionContext pluginExecutionContext;
        /// <summary>
        /// Current PluginExecutionContext
        /// </summary>
        public IPluginExecutionContext PluginExecutionContext
        {
            get
            {
                if (pluginExecutionContext == null)
                {
                    pluginExecutionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                }
                return pluginExecutionContext;
            }
        }

        private IOrganizationService adminOrganizationService;
        /// <summary>
        /// OrranizationService under system credentials
        /// </summary>
        public IOrganizationService AdminOrganizationService
        {
            get
            {
                if (adminOrganizationService == null)
                {
                    IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    adminOrganizationService = factory.CreateOrganizationService(null);
                }
                return adminOrganizationService;
            }
        }

        private IOrganizationService organizationService;
        /// <summary>
        /// Organization service under current user
        /// </summary>
        public IOrganizationService OrganizationService
        {
            get
            {
                if (organizationService == null)
                {
                    IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    organizationService = factory.CreateOrganizationService(Guid.Empty);
                }
                return organizationService;
            }
        }

        /// <summary>
        /// Gets orgatization service under specific user credentials
        /// </summary>
        /// <param name="userId">User id</param>
        /// <returns>Organization service</returns>
        public IOrganizationService GetUserService(Guid userId)
        {
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            return factory.CreateOrganizationService(userId);
        }

        /// <summary>
        /// Current plugin stage
        /// </summary>
        public PluginStage Stage
        {
            get
            {
                return (PluginStage)PluginExecutionContext.Stage;
            }
        }

        /// <summary>
        /// Current user info 
        /// </summary>
        public Guid UserId
        {
            get
            {
                return pluginExecutionContext.InitiatingUserId;
            }
        }

        /// <summary>
        /// Plugin message name
        /// </summary>
        public string MessageName
        {
            get
            {
                return PluginExecutionContext.MessageName;
            }
        }

        /// <summary>
        /// Current primary entity name
        /// </summary>
        public string PrimaryEntityName
        {
            get
            {
                return PluginExecutionContext.PrimaryEntityName;
            }
        }

        /// <summary>
        /// Target image entity
        /// </summary>
        public Entity TargetImageEntity
        {
            get
            {
                Entity targetImageEntity = (PluginExecutionContext.InputParameters != null && PluginExecutionContext.InputParameters.Contains("Target")) ? PluginExecutionContext.InputParameters["Target"] as Entity : null;
                if (targetImageEntity == null)
                {
                    throw new Exception("Target image entity doesn't exist in current context.");
                }
                return targetImageEntity;
            }
            private set { }
        }

        /// <summary>
        /// Target image entity reference
        /// </summary>
        public EntityReference TargetImageEntityReference
        {
            get
            {
                EntityReference targetImageEntityReference = (PluginExecutionContext.InputParameters != null && PluginExecutionContext.InputParameters.Contains("Target")) ?
                    PluginExecutionContext.InputParameters["Target"] as EntityReference : null;
                if (targetImageEntityReference == null)
                {
                    throw new Exception("Target image entity reference doesn't exist.");
                }
                return targetImageEntityReference;
            }
            private set { }
        }
        /// <summary>
        /// Pre image entity
        /// </summary>
        public Entity PreImageEntity
        {
            get
            {
                Entity preImageEntity = (PluginExecutionContext.PreEntityImages != null && PluginExecutionContext.PreEntityImages.Contains("preImage")) ?
                    PluginExecutionContext.PreEntityImages["preImage"] as Entity : null;
                if (preImageEntity == null)
                {
                    throw new Exception("Pre image entity doesn't exist.");
                }
                return preImageEntity;
            }
            private set { }
        }
        /// <summary>
        /// Post image entity
        /// </summary>
        public Entity PostImageEntity
        {
            get
            {
                Entity postImageEntity = (PluginExecutionContext.PostEntityImages != null && PluginExecutionContext.PostEntityImages.Contains("postImage")) ?
                    PluginExecutionContext.PostEntityImages["postImage"] as Entity : null;
                if (postImageEntity == null)
                {
                    throw new Exception("Post image entity doesn't exist.");
                }
                return postImageEntity;
            }
            private set { }
        }

        public ITracingService TracingService
        {
            get
            {
                return (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            }
            private set { }
        }

        internal PluginContext(IServiceProvider serviceProvider)
        {
            this.serviceProvider = serviceProvider;
        }
    }
}
