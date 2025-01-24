using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace CasePlugin
{
    /// <summary>
    /// Plugin to stop new case creation for a customer if that customer has an existing active case.
    /// </summary>
    public class CasePlugin : IPlugin
    {
        /// <summary>
        /// Executes plugin logic.
        /// </summary>
        /// <param name="serviceProvider">The service provider.</param>
        /// <exception cref="InvalidPluginExecutionException">Thrown when plugin execution encounters an invalid operation.</exception>
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));
            var tracingService = (ITracingService) serviceProvider.GetService(typeof(ITracingService));
            var serviceFactory = (IOrganizationServiceFactory) serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            int stageNumber = 0;
            try
            {
                tracingService.Trace($"Stage {++stageNumber}: Plugin execution started");

                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
                {
                    // Retrieve target entity (current case) from context.
                    var entity = (Entity)context.InputParameters["Target"];
                    tracingService.Trace($"Stage {++stageNumber}: Retrieved target entity");

                    if (entity.Attributes.Contains("customerid") && entity["customerid"] is EntityReference)
                    {
                        // Retrieve customer ID from case.
                        var customer = (EntityReference)entity["customerid"];
                        Guid customerId = customer.Id;
                        tracingService.Trace($"Stage {++stageNumber}: Retrieved customer ID");

                        // Build query to check for existing cases linked to that customer ID.
                        QueryExpression queryExpression = BuildQuery(customerId);
                        tracingService.Trace($"Stage {++stageNumber}: Created query to check for existing incidents");

                        // Execute query & handle.
                        bool existingCases = CheckForExistingCases(service, queryExpression);
                        if (existingCases)
                        {
                            // If cases exist for this customer, throw exception to stop case creation.
                            tracingService.Trace($"Stage {++stageNumber}: Found existing case for customer ID: {customerId}. Case creation aborted.");
                            throw new InvalidPluginExecutionException($"Case already exists for customer ID: {customerId}.");
                        }
                        else
                        {
                            // Else, allow case creation and add to trace log.
                            tracingService.Trace($"Stage {++stageNumber}: No existing case for customer. New case created successfully.");
                        }
                    }
                    else
                    {
                        tracingService.Trace($"Stage {++stageNumber}: Unable to retrieve customer ID.");
                        throw new InvalidPluginExecutionException("Cannot retrieve customer ID from entity.");
                    }

                }
                else
                {
                    tracingService.Trace($"Stage {++stageNumber}: Unable to retrieve target entity from context.");
                    throw new InvalidPluginExecutionException("Cannot retrieve target entity from context.");
                }
            } 
            catch (InvalidPluginExecutionException ex)
            {
                // Catch and log any InvalidPluginExecutionExceptions occurred during plugin execution.
                tracingService.Trace($"Plugin execution failed: {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                // Catch and log any unexpected errors occurred during plugin execution.
                tracingService.Trace($"Unexpected error: {ex.Message}");
                throw new InvalidPluginExecutionException("An unexpected error has occurred.", ex);
            }         
        }

        /// <summary>
        /// Builds a query to check for existing active cases for specified customer.
        /// </summary>
        /// <param name="customerId">The customer ID.</param>
        /// <returns>A QueryExpression to check for existing active cases for specific customer ID.</returns>
        private static QueryExpression BuildQuery(Guid customerId)
        {
            QueryExpression queryExpression = new QueryExpression("incident")
            {
                ColumnSet = new ColumnSet("incidentid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("customerid", ConditionOperator.Equal, customerId),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 0)
                    }
                },
                TopCount = 1
            };
            return queryExpression;
        }

        /// <summary>
        /// Executes a query to check for existing active cases for a customer.
        /// </summary>
        /// <param name="service">The organisation service.</param>
        /// <param name="queryExpression">The query expression.</param>
        /// <returns>True if existing active cases are found, else false.</returns>
        private static bool CheckForExistingCases(IOrganizationService service, QueryExpression queryExpression)
        {
            EntityCollection results = service.RetrieveMultiple(queryExpression);
            return results.Entities.Count > 0;
        }
    }
}
