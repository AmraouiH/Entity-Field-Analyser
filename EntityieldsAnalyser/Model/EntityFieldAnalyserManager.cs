using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using XrmToolBox.Extensibility;
using Label = System.Windows.Forms.Label;

namespace EntityieldsAnalyser
{
    public enum EntityType { All, Custom, Standard, Filter }

    class EntityFieldAnalyserManager
    {
        #region Variables
        private IOrganizationService _service;
        private IEnumerable<EntityMetadata> _metadataList;
        public static EntityInfo entityInfo          = null;

        #endregion
        #region Get EntityFields
        public static Dictionary<AttributeTypeCode, List<entityParam>> getEntityFields(IOrganizationService service, String entityTechnicalName, String entityName, BackgroundWorker worker)
        {
            entityInfo = new EntityInfo();
            Dictionary<AttributeTypeCode, List<entityParam>> _data = new Dictionary<AttributeTypeCode, List<entityParam>>();
            RetrieveEntityRequest retrieveBankAccountEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entityTechnicalName,
                RetrieveAsIfPublished = true
            };

            worker.ReportProgress(0,"Retrieving Entity MetaData...");
            RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveBankAccountEntityRequest);
            _data = formatList(retrieveEntityResponse, _data);
            EntityCollection _entityRecords               = getEntityRecords(service, worker, entityTechnicalName);
            worker.ReportProgress(0, "Analysing ...");
            entityInfo.entityRecordsCount                 = _entityRecords.Entities.Count;
            entityInfo.entityName                         = entityName;
            entityInfo.entityTechnicalName                = entityTechnicalName;
            entityInfo.entityDateOfCreation               = retrieveEntityResponse.EntityMetadata.CreatedOn != null ? (DateTime)retrieveEntityResponse.EntityMetadata.CreatedOn.Value.Date : DateTime.MinValue.Date;

            return setDictionaryCount(_entityRecords, _data); ;
        }
        #endregion
        #region formatList function
        // ordering result data to a dictionary
        private static Dictionary<AttributeTypeCode, List<entityParam>> formatList(RetrieveEntityResponse _metadata, Dictionary<AttributeTypeCode, List<entityParam>> _data)
        {
            var attributes = FilterAttributes(_metadata.EntityMetadata.Attributes);
            entityInfo.entityFieldsCount = attributes.Count();
            foreach (var field in attributes)
            {
                    if (!_data.ContainsKey(field.AttributeType.Value))
                    {
                        _data.Add(field.AttributeType.Value, new List<entityParam>() { setObject(field) });
                    }
                    else
                    {
                        _data[field.AttributeType.Value].Add(setObject(field));
                    }
            }

            return _data;
        }
        #endregion
        #region setObject Function for Data Dictionary
        private static entityParam setObject(AttributeMetadata field)
        {
            return new entityParam()
            {
                displayName       = field.DisplayName.UserLocalizedLabel != null ? field.DisplayName.UserLocalizedLabel.Label : @"N/A",
                fieldName         = field.LogicalName,
                isManaged         = field.IsManaged == true ? "Managed" : "Unmanaged",
                target            = (field.AttributeType.Value == AttributeTypeCode.Lookup || field.AttributeType.Value == AttributeTypeCode.Owner) && ((LookupAttributeMetadata)field).Targets.Length > 0 ? ((LookupAttributeMetadata)field).Targets[0] : String.Empty,
                dateOfCreation    = field.CreatedOn != null ? field.CreatedOn.Value.Date : DateTime.MinValue,
                introducedVersion = field.IntroducedVersion,
                isAuditable       = field.IsAuditEnabled.Value,
                requiredLevel     = ((AttributeRequiredLevel)field.RequiredLevel.Value).ToString(),
                isSearchable      = (bool)field.IsSearchable,
                isCustom          = (bool)field.IsCustomAttribute,
                totalFiledRecords = 0
            };
        }
        #endregion
        #region Get Entity Records
        private static EntityCollection getEntityRecords(IOrganizationService service, BackgroundWorker worker, String entityName)
        {
            EntityCollection entities = new EntityCollection();
            QueryExpression query = new QueryExpression(entityName);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            var PageCookie = String.Empty;
            var PageNumber = 1;
            var PageSize = 5000;

            EntityCollection result;
            do
            {
                query.PageInfo = new PagingInfo() { PageNumber = 1, Count = PageSize };

                if (PageNumber != 1)
                {
                    query.PageInfo.PageNumber   = PageNumber;
                    query.PageInfo.PagingCookie = PageCookie;
                }
                result = service.RetrieveMultiple(query);

                if (result.MoreRecords)
                {
                    entities.Entities.AddRange(result.Entities);
                    PageNumber++;
                    PageCookie = result.PagingCookie;
                }
                worker.ReportProgress(0, $"Retrieving Entity Records ...{Environment.NewLine}"+ entities.Entities.Count+" Records");
            }
            while (result.MoreRecords);
            entities.Entities.AddRange(result.Entities);

            return entities;
        }
        #endregion
        #region Calculate how much record is containing the field
        private static Dictionary<AttributeTypeCode, List<entityParam>> setDictionaryCount(EntityCollection entityCollection, Dictionary<AttributeTypeCode, List<entityParam>> _data)
        {
            var _totalRecords = entityCollection.Entities.Count;
            foreach (var entityParams in _data.Values)
            {
                foreach (var entityParam in entityParams)
                {
                    foreach (var record in entityCollection.Entities)
                    {
                        if (record.Contains(entityParam.fieldName))
                        {
                            entityParam.totalFiledRecords++;
                        }
                    }
                    //for Managed/Unmanaged Chart
                    if (entityParam.isManaged == "Managed")
                        entityInfo.managedFieldsCount++;
                    else
                        entityInfo.unmanagedFieldsCount++;

                    //For Custom Standard Chart
                    if (entityParam.isCustom)
                        entityInfo.entityCustomFieldsCount++;
                    else
                        entityInfo.entityStandardFieldsCount++;

                }
            }
            return CalculatePercentageOfUse(_totalRecords, _data);
        }


        #endregion
        #region calculate The Percentage of use of each field
        private static Dictionary<AttributeTypeCode, List<entityParam>> CalculatePercentageOfUse(int recordsCount, Dictionary<AttributeTypeCode, List<entityParam>> _data)
        {
            foreach (var elem in _data.Values)
            {
                foreach (var field in elem)
                {
                    if (field.totalFiledRecords > 0)
                        field.percentageOfUse = ((float)(field.totalFiledRecords * 100.0 / recordsCount)).ToString("0.##\\%");
                    else
                        field.percentageOfUse = 0.ToString("0.##\\%");
                }
            }
            return _data;
        }
        #endregion
        #region Calculate The Total of Columns used in the ENtity(Database Table)
        private static int DbUsedColumns(IDictionary<AttributeTypeCode, List<entityParam>> dict)
        {
            var currentSize = 0;
            foreach (var element in dict)
            {
                if (element.Key == AttributeTypeCode.Owner || element.Key == AttributeTypeCode.Lookup)
                {
                    currentSize += dict[element.Key].Count * 3;
                }
                else if (element.Key == AttributeTypeCode.Status || element.Key == AttributeTypeCode.Boolean || element.Key == AttributeTypeCode.Picklist || element.Key == AttributeTypeCode.Money)
                {
                    currentSize += dict[element.Key].Count * 2;
                }
                else
                {
                    currentSize += dict[element.Key].Count;
                }
            }

            return currentSize;
        }
        #endregion
        #region ChartFieldsType
        public static void SetChartFieldsType(Dictionary<AttributeTypeCode, List<entityParam>> dict, Chart ChartFieldTypes, bool displayValues = true)
        {
            var fieldsCount = 0;
            foreach (var e in dict.Values)
            {
                fieldsCount += e.Count;
            }

            //Disable formating when display the count
            if (displayValues)
                ChartFieldTypes.Series["fieldsReport"].LabelFormat = String.Empty;
            else//enable formating to percentage when check the checkbox
                ChartFieldTypes.Series["fieldsReport"].LabelFormat = "0.#%";

            //refrech chart data
            if (ChartFieldTypes.Series["fieldsReport"].Points.Count > 0) {
                ChartFieldTypes.Series["fieldsReport"].Points.Clear();
            }
            //Add Data to chart
            foreach (var element in dict)
            {
                ChartFieldTypes.Series["fieldsReport"].Points.AddXY(element.Key.ToString(), displayValues ? element.Value.Count : (element.Value.Count * 1.0 / fieldsCount));
            }
            //Add Sorting to Chart
            ChartFieldTypes.Series["fieldsReport"].IsValueShownAsLabel = true;
            ChartFieldTypes.DataManipulator.Sort(PointSortOrder.Descending, ChartFieldTypes.Series["fieldsReport"]);
            ChartFieldTypes.Dock = DockStyle.Fill;
        }
        #endregion
        #region ChartFieldAvailable
        public static void SetChartFieldAvailable( Dictionary<AttributeTypeCode, List<entityParam>> dict, Chart ChartFieldAvailabe, bool displayPercentage = false)
        {
            entityInfo.entityTotalUseOfColumns = DbUsedColumns(dict);
            var availableFieldToCreate = entityInfo.entityDefaultColumnSize - entityInfo.entityTotalUseOfColumns;

            if (!displayPercentage)
            {
                ChartFieldAvailabe.Series["AvailableField"].LabelFormat         = String.Empty;
                ChartFieldAvailabe.Series["AvailableField"].IsValueShownAsLabel = false;
            }
            else//enable formating to percentage when check the checkbox
            {
                ChartFieldAvailabe.Series["AvailableField"].LabelFormat         = "0.#%";
                ChartFieldAvailabe.Series["AvailableField"].IsValueShownAsLabel = true;
            }

            #region Clear Chart Data
            if (ChartFieldAvailabe.Series["AvailableField"].Points.Count > 0) {
                ChartFieldAvailabe.Series["AvailableField"].Points.Clear();
            }
            #endregion
            #region set chart Data
            ChartFieldAvailabe.Series["AvailableField"].Points.AddXY("Available Fields To Create ", displayPercentage ? availableFieldToCreate * 1.0 / entityInfo.entityDefaultColumnSize : availableFieldToCreate);
            ChartFieldAvailabe.Series["AvailableField"].Points.AddXY("Created Fields", displayPercentage ? entityInfo.entityTotalUseOfColumns * 1.0 / entityInfo.entityDefaultColumnSize : entityInfo.entityTotalUseOfColumns);
            ChartFieldAvailabe.Dock = DockStyle.Fill;
            #endregion
        }
        #endregion
        #region SetFieldDataGridViewHeaders
        public static void SetFieldDataGridViewHeaders(DataTable dt, DataGridView fieldPropretiesView)
        {
            fieldPropretiesView.DataSource = dt;
            fieldPropretiesView.Sort(fieldPropretiesView.Columns[2], ListSortDirection.Ascending);
            fieldPropretiesView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Regular);
            fieldPropretiesView.Columns[0].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[0].HeaderText              = "Display Name";
            fieldPropretiesView.Columns[1].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[1].HeaderText              = "Schema Name";
            fieldPropretiesView.Columns[2].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[2].HeaderText              = "Managed/Unmanaged";
            fieldPropretiesView.Columns[3].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[3].HeaderText              = "IsAuditable";
            fieldPropretiesView.Columns[4].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[4].HeaderText              = "IsSearchable";
            fieldPropretiesView.Columns[5].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[5].HeaderText              = "Required Level";
            fieldPropretiesView.Columns[6].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[6].HeaderText              = "Introduced Version";
            fieldPropretiesView.Columns[7].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[7].HeaderText              = "CreatedOn";
            fieldPropretiesView.Columns[8].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[8].HeaderText              = "Target";
            fieldPropretiesView.Columns[9].AutoSizeMode            = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[9].HeaderText              = "Percentage Of Use";

            fieldPropretiesView.ReadOnly                           = true;
            fieldPropretiesView.Columns[8].Visible                 = false;
        }
        #endregion
        #region SetEntitiesGridViewHeaders
        public static void SetEntitiesGridViewHeaders(DataTable dt, DataGridView entityGridView)
        {
            entityGridView.DataSource = dt;
            entityGridView.Sort(entityGridView.Columns[1], ListSortDirection.Ascending);
            entityGridView.ColumnHeadersDefaultCellStyle.Font    = new Font("Segoe UI Semibold", 9.75F, FontStyle.Regular);
            entityGridView.Columns[0].AutoSizeMode               = DataGridViewAutoSizeColumnMode.Fill;
            entityGridView.Columns[0].HeaderText                 = "Display Name";
            entityGridView.Columns[1].AutoSizeMode               = DataGridViewAutoSizeColumnMode.Fill;
            entityGridView.Columns[1].HeaderText                 = "Schema Name";
            entityGridView.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            entityGridView.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
        }
        #endregion
        #region GetEntities
        public static RetrieveMetadataChangesResponse GetEntitiesMetadat(IOrganizationService service, string selectedTypeOfEntities) {

            MetadataFilterExpression entityFilter = new MetadataFilterExpression(LogicalOperator.And);

            switch (selectedTypeOfEntities)
            {
                case "All Entities":
                    break;
                case "Custom Entities":
                    entityFilter.Conditions.Add(new MetadataConditionExpression("IsCustomEntity", MetadataConditionOperator.Equals, true));
                    break;
                case "CRM Entities":
                    entityFilter.Conditions.Add(new MetadataConditionExpression("IsCustomEntity", MetadataConditionOperator.Equals, false));
                    break;
                case "System Entities":
                    entityFilter.Conditions.Add(new MetadataConditionExpression("IsManaged", MetadataConditionOperator.Equals, true));
                    break;
                default:
                    MessageBox.Show("Error: Entities Drop Down was changed, please restart the tool and try again.");
                    break;
            }

            EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter,
                Properties = new MetadataPropertiesExpression("LogicalName", "DisplayName")
            };
            RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };

            return (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);
        }
        #endregion
        public static void SetChartManagedUnmanagedFields(Chart managedUnmanagedFieldsChart, bool displayValues=true) {
            if (managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].Points.Count > 0)
                managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].Points.Clear();


            //Disable formating when display the count
            if (displayValues)
                managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].LabelFormat = String.Empty;
            else//enable formating to percentage when check the checkbox
                managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].LabelFormat = "0.#%";

            managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].Points.AddXY("Managed Fields", displayValues ? entityInfo.managedFieldsCount : (entityInfo.managedFieldsCount * 1.0 / (entityInfo.managedFieldsCount + entityInfo.unmanagedFieldsCount)));
            managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].Points.AddXY("Unmanaged Fields", displayValues ? entityInfo.unmanagedFieldsCount : (entityInfo.unmanagedFieldsCount * 1.0 / (entityInfo.managedFieldsCount + entityInfo.unmanagedFieldsCount)));
            managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].IsValueShownAsLabel = true;
            managedUnmanagedFieldsChart.Dock = DockStyle.Fill;
        }

        public static bool CanICreateThisNumberOfFields(String lookupsType, String pickListTypes, String othersTypes)
        {
            int askedCreatedFieldsColumnsSize = (int.Parse(lookupsType) * 3) + (int.Parse(pickListTypes) * 2) + (int.Parse(othersTypes) * 1);
            int availableFieldsToCreate = entityInfo.entityDefaultColumnSize - entityInfo.entityTotalUseOfColumns;
            if (availableFieldsToCreate > askedCreatedFieldsColumnsSize)
                return true;
            else
                return false;
        }

        public static void SetFieldDataGridViewContent(List<entityParam> entities, DataTable dtFields) {

            foreach (var item in entities)
            {
                DataRow row = dtFields.NewRow();

                row["Display Name"]         = item.displayName;
                row["Schema Name"]          = item.fieldName;
                row["Managed/Unmanaged"]    = item.isManaged;
                row["IsAuditable"]          = item.isAuditable;
                row["IsSearchable"]         = item.isSearchable;
                row["Required Level"]       = item.requiredLevel;
                row["Introduced Version"]   = item.introducedVersion;
                row["CreatedOn"]            = item.dateOfCreation;
                row["Percentage Of Use"]    = item.percentageOfUse;
                row["Target"]               = item.target;

                dtFields.Rows.Add(row);
            }
        }

        public static void setStatisticsFieldText(Label statisticsText) {
            if (entityInfo.entityRecordsCount > 0)
                statisticsText.Text = "Statistics Based On " + entityInfo.entityRecordsCount + " Records";
            else
                statisticsText.Text = String.Empty;
        }

        public static void  CallExportFunction(Dictionary<AttributeTypeCode, List<entityParam>> entityParams) {
            FileManaged.ExportFile(entityParams, entityInfo);
        }

        public static string[] SelectedEntity(DataGridViewRowCollection rows)
        {
            List<String> selectedEntity = new List<string>();
            foreach (DataGridViewRow row in rows)
            {
                var cell = row.Cells[2];
                if (cell != null)
                {
                    if ((bool)cell.Value != null && (bool)cell.Value == true)
                    {
                        selectedEntity.Add(row.Cells[0].Value.ToString());
                        selectedEntity.Add(row.Cells[1].Value.ToString());
                    }
                }
            }
            return selectedEntity.ToArray();
        }

        public static IEnumerable<AttributeMetadata> FilterAttributes(AttributeMetadata[] attributes)
        {
            return attributes.Where(a =>
                            a.AttributeOf == null
                            && a.AttributeType.Value != AttributeTypeCode.Virtual
                            && a.AttributeType.Value != AttributeTypeCode.PartyList
                            && a.IsValidForRead.Value
                            && a.LogicalName.IndexOf("composite") < 0
                            ).OrderBy(a => a.LogicalName);
        }
    }
}