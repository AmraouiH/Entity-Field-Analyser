using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntityieldsAnalyser.Model
{
    public class EntityFieldAnalyserModel
    {

        public string EntityName { get; set; }
        public string EntitySchemaName { get; set; }
        public int NumberOfCustomAttributes { get; set; }
        public int RecordCount { get; set; }
        public bool HasModificationDates { get; set; }
        public DateTime LastCreated { get; set; }
        public DateTime LastModified { get; set; }
        public string ErrorMessage { get; set; }
    }
    public static class EntityUsageExtensions
    {
        public static void Count(this EntityFieldAnalyserModel entityUsage, IOrganizationService service)
        {
            try
            {
                int totalCount = 0;
                DateTime lastCreated = DateTime.MinValue;
                DateTime lastModified = DateTime.MinValue;

                QueryExpression query = new QueryExpression(entityUsage.EntitySchemaName);
                query.Distinct = true;
                if (totalCount > 0)
                {
                    query.ColumnSet = new ColumnSet("createdon", "modifiedon");
                }
                else
                {
                    query.ColumnSet = new ColumnSet(false);
                }
                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                if (entityUsage.HasModificationDates)
                {
                    query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
                    query.Orders.Add(new OrderExpression("modifiedon", OrderType.Descending));
                }
                EntityCollection entityCollection = service.RetrieveMultiple(query);
                totalCount = entityCollection.Entities.Count;
                if (totalCount > 0 && entityUsage.HasModificationDates)
                {
                    lastCreated = entityCollection.Entities.First().GetAttributeValue<DateTime>("createdon");
                    lastModified = entityCollection.Entities.First().GetAttributeValue<DateTime>("modifiedon");
                }

                while (entityCollection.MoreRecords)
                {
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                    entityCollection = service.RetrieveMultiple(query);
                    totalCount = totalCount + entityCollection.Entities.Count;
                    if (entityCollection.Entities.Count > 0 && entityUsage.HasModificationDates)
                    {
                        lastCreated = entityCollection.Entities.First().GetAttributeValue<DateTime>("createdon");
                        lastModified = entityCollection.Entities.First().GetAttributeValue<DateTime>("modifiedon");
                    }
                }

                entityUsage.RecordCount = totalCount;
                entityUsage.LastCreated = lastCreated;
                entityUsage.LastModified = lastModified;
            }
            catch (Exception ex)
            {
                entityUsage.ErrorMessage = ex.Message;
            }


        }
    }
}
