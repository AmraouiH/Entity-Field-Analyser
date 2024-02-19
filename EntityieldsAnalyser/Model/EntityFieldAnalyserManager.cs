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
        public static Dictionary<AttributeTypeCode, List<entityParam>> _data = null;

        #endregion
        #region Get EntityFields
        public static Dictionary<AttributeTypeCode, List<entityParam>> getEntityFields(IOrganizationService service, String entityTechnicalName, String entityName, BackgroundWorker worker, String analyseType)
        {
            entityInfo = new EntityInfo();
            _data = new Dictionary<AttributeTypeCode, List<entityParam>>();
            RetrieveEntityRequest retrieveBankAccountEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entityTechnicalName,
                RetrieveAsIfPublished = true
            };

            worker.ReportProgress(0,"Retrieving Entity MetaData...");
            RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveBankAccountEntityRequest);
            formatList(retrieveEntityResponse, _data);

            if(analyseType == "Metadata + Data usage")
                getEntityRecords(service, worker, _data, entityTechnicalName, entityInfo);

            CalculateDataForChartManagedUnmanaged(_data);

            worker.ReportProgress(0, "Analysing ...");
            entityInfo.entityName                         = entityName;
            entityInfo.entityTechnicalName                = entityTechnicalName;
            entityInfo.entityDateOfCreation               = retrieveEntityResponse.EntityMetadata.CreatedOn != null ? (DateTime)retrieveEntityResponse.EntityMetadata.CreatedOn.Value.Date : DateTime.MinValue.Date;

            return CalculatePercentageOfUse(entityInfo.entityRecordsCount, _data);
        }
        #endregion
        #region formatList function
        // ordering result data to a dictionary
        private static void formatList(RetrieveEntityResponse _metadata, Dictionary<AttributeTypeCode, List<entityParam>> _data)
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
        }
        #endregion
        #region setObject Function for Data Dictionary
        private static entityParam setObject(AttributeMetadata field)
        {
            entityParam e = new entityParam();
            e.displayName = field?.DisplayName?.UserLocalizedLabel != null ? field?.DisplayName?.UserLocalizedLabel?.Label : @"N/A";
            e.fieldName = field?.LogicalName;
            e.attributeType = field.AttributeType.Value.ToString();
            e.isManaged = field?.IsManaged == true ? "Managed" : "Unmanaged";
            e.target = (field.AttributeType.Value == AttributeTypeCode.Lookup || field.AttributeType.Value == AttributeTypeCode.Owner) && ((LookupAttributeMetadata)field).Targets.Length > 0 ? ((LookupAttributeMetadata)field).Targets[0] : String.Empty;
            e.dateOfCreation = field.CreatedOn != null ? field.CreatedOn.Value.Date : DateTime.MinValue;
            e.introducedVersion = field?.IntroducedVersion;
            e.isAuditable = field?.IsAuditEnabled?.Value;
            e.requiredLevel = ((AttributeRequiredLevel)field?.RequiredLevel?.Value).ToString();
            e.isSearchable = (bool)field?.IsValidForAdvancedFind?.Value;
            e.isCustom = (bool)field?.IsCustomAttribute;
            e.totalFiledRecords = 0;
            e.AttributeOf = field?.AttributeOf;
            e.AutoNumberFormat = field?.AutoNumberFormat;
            e.CanBeSecuredForCreate = (bool)field?.CanBeSecuredForCreate;
            e.CanBeSecuredForRead = (bool)field?.CanBeSecuredForRead;
            e.CanBeSecuredForUpdate = (bool)field?.CanBeSecuredForUpdate;
            e.CanModifyAdditionalSettings = (bool)field?.CanModifyAdditionalSettings?.Value;
            e.ColumnNumber = (int)field?.ColumnNumber;
            e.DeprecatedVersion = field?.DeprecatedVersion != null ? field?.DeprecatedVersion : "";
            e.Description = field?.Description?.UserLocalizedLabel?.Label;
            e.EntityLogicalName = field?.EntityLogicalName;
            e.ExternalName = field?.ExternalName;
            e.InheritsFrom = field?.InheritsFrom;
            e.IsCustomizable = (bool)field?.IsCustomizable?.Value;
            e.IsDataSourceSecret = (bool)field?.IsDataSourceSecret;
            e.IsFilterable = (bool)field?.IsFilterable;
            e.IsGlobalFilterEnabled = (bool)field?.IsGlobalFilterEnabled?.Value;
            e.IsLogical = (bool)field?.IsLogical;
            e.IsPrimaryId = (bool)field?.IsPrimaryId;
            e.IsPrimaryName = (bool)field?.IsPrimaryName;
            e.IsRenameable = (bool)field?.IsRenameable?.Value;
            e.IsRequiredForForm = (bool)field?.IsRequiredForForm;
            e.IsRetrievable = (bool)field?.IsRetrievable;
            e.IsSecured = (bool)field?.IsSecured;
            e.IsSortableEnabled = (bool)field?.IsSortableEnabled?.Value;
            e.IsValidForAdvancedFind = (bool)field?.IsValidForAdvancedFind?.Value;
            e.IsValidForCreate = (bool)field?.IsValidForCreate;
            e.IsValidForForm = (bool)field?.IsValidForForm;
            e.IsValidForGrid = (bool)field?.IsValidForGrid;
            e.IsValidForRead = (bool)field?.IsValidForRead;
            e.IsValidForUpdate = (bool)field?.IsValidForUpdate;
            e.IsValidODataAttribute = field?.IsValidODataAttribute;
            e.LinkedAttributeId = field?.LinkedAttributeId?.ToString();
            e.ModifiedOn = field?.ModifiedOn != null ? field.ModifiedOn.Value.Date : DateTime.MinValue;
            e.SourceType = field?.SourceType != null ? field?.SourceType?.ToString() : "";

            return e;
        }
        #endregion
        #region Data For Chat Managed/Unmanaged
        private static void CalculateDataForChartManagedUnmanaged(Dictionary<AttributeTypeCode, List<entityParam>> _data) {
            foreach (var entityParams in _data.Values)
            {
                foreach (var entityParam in entityParams)
                {
                    //for Managed/Unmanaged Chart
                    if (entityParam.isManaged == "Managed")
                        entityInfo.managedFieldsCount++;
                    else
                        entityInfo.unmanagedFieldsCount++;

                    //For Custom Standard Chart
                    if (entityParam.isCustom != null && entityParam.isCustom == true)
                        entityInfo.entityCustomFieldsCount++;
                    else
                        entityInfo.entityStandardFieldsCount++;     
                }
            }
        }

        #endregion


        #region Get Entity Records
        private static void getEntityRecords(IOrganizationService service, BackgroundWorker worker, Dictionary<AttributeTypeCode, List<entityParam>> _data, String entityName, EntityInfo entityInfo)
        {
            EntityCollection entities = new EntityCollection();
            QueryExpression query = new QueryExpression(entityName);
            query.ColumnSet = new ColumnSet(true);
            query.NoLock = true;
            var PageCookie = String.Empty;
            var PageNumber = 1;
            var PageSize = 5000;
            var totalrecords = 0;

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
                totalrecords += result.Entities.Count();
                worker.ReportProgress(0, $"Retrieving Entity Records ...{Environment.NewLine}"+ totalrecords + " Records");
                setDictionaryCount(entities, _data, (PageNumber > 2 ? false : true));
            }
            while (result.MoreRecords);
            entities.Entities.AddRange(result.Entities);
            setDictionaryCount(entities, _data, false);

            entityInfo.entityRecordsCount = totalrecords;
        }
        #endregion
        #region Calculate how much record is containing the field
        private static void setDictionaryCount(EntityCollection entityCollection, Dictionary<AttributeTypeCode, List<entityParam>> _data, bool shouldCalculateForReporting)
        {
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
                }
            }
            entityCollection.Entities.Clear();
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
                        field.percentageOfUse = (field.totalFiledRecords * 100.0 / recordsCount);
                    else
                        field.percentageOfUse = 0;
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
        public static void SetFieldDataGridViewHeaders(DataTable dt, DataGridView fieldPropretiesView, String analyseType, bool shouldDisplayAllColumns, string selectedType)
        {
            fieldPropretiesView.DataSource = dt;
            fieldPropretiesView.RowHeadersVisible = false;
            fieldPropretiesView.Sort(fieldPropretiesView.Columns[4], ListSortDirection.Ascending);
            fieldPropretiesView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Regular);
            fieldPropretiesView.Columns[0].HeaderText              = "Display Name";
            fieldPropretiesView.Columns[1].HeaderText              = "Schema Name";
            fieldPropretiesView.Columns[2].HeaderText              = "Type";
            fieldPropretiesView.Columns[3].HeaderText              = "Description";
            fieldPropretiesView.Columns[4].HeaderText              = "Target";
            fieldPropretiesView.Columns[5].HeaderText              = "Managed/Unmanaged";
            fieldPropretiesView.Columns[6].HeaderText              = "IsCustom";
            fieldPropretiesView.Columns[7].HeaderText              = "IsAuditable";
            fieldPropretiesView.Columns[8].HeaderText              = "IsSearchable";
            fieldPropretiesView.Columns[9].HeaderText              = "Required Level";
            fieldPropretiesView.Columns[10].HeaderText              = "Introduced Version";
            fieldPropretiesView.Columns[11].HeaderText             = "CreatedOn";
            fieldPropretiesView.Columns[12].HeaderText             = "ModifiedOn";
            fieldPropretiesView.Columns[13].HeaderText             = "AttributeOf";
            fieldPropretiesView.Columns[14].HeaderText             = "AutoNumberFormat";
            fieldPropretiesView.Columns[15].HeaderText             = "CanBeSecuredForCreate";
            fieldPropretiesView.Columns[16].HeaderText             = "CanBeSecuredForRead";
            fieldPropretiesView.Columns[17].HeaderText             = "CanBeSecuredForUpdate";
            fieldPropretiesView.Columns[18].HeaderText             = "CanModifyAdditionalSettings";
            fieldPropretiesView.Columns[19].HeaderText             = "ColumnNumber";
            fieldPropretiesView.Columns[20].HeaderText             = "DeprecatedVersion";
            fieldPropretiesView.Columns[21].HeaderText             = "ExternalName";
            fieldPropretiesView.Columns[22].HeaderText             = "InheritsFrom";
            fieldPropretiesView.Columns[23].HeaderText             = "IsCustomizable";
            fieldPropretiesView.Columns[24].HeaderText             = "IsDataSourceSecret";
            fieldPropretiesView.Columns[25].HeaderText             = "IsFilterable";
            fieldPropretiesView.Columns[26].HeaderText             = "IsGlobalFilterEnabled";
            fieldPropretiesView.Columns[27].HeaderText             = "IsLogical";
            fieldPropretiesView.Columns[28].HeaderText             = "IsPrimaryId";
            fieldPropretiesView.Columns[29].HeaderText             = "IsPrimaryName";
            fieldPropretiesView.Columns[30].HeaderText             = "IsRenameable";
            fieldPropretiesView.Columns[31].HeaderText             = "IsRequiredForForm";
            fieldPropretiesView.Columns[32].HeaderText             = "IsRetrievable";
            fieldPropretiesView.Columns[33].HeaderText             = "IsSecured";
            fieldPropretiesView.Columns[34].HeaderText             = "IsSortableEnabled";
            fieldPropretiesView.Columns[35].HeaderText             = "IsValidForAdvancedFind";
            fieldPropretiesView.Columns[36].HeaderText             = "IsValidForCreate";
            fieldPropretiesView.Columns[37].HeaderText             = "IsValidForForm";
            fieldPropretiesView.Columns[38].HeaderText             = "IsValidForGrid";      
            fieldPropretiesView.Columns[39].HeaderText             = "IsValidForRead";
            fieldPropretiesView.Columns[40].HeaderText             = "IsValidForUpdate";
            fieldPropretiesView.Columns[41].HeaderText             = "IsValidODataAttribute";
            fieldPropretiesView.Columns[42].HeaderText             = "LinkedAttributeId";
            fieldPropretiesView.Columns[43].HeaderText             = "EntityLogicalName";
            fieldPropretiesView.Columns[44].HeaderText             = "SourceType";

            if (analyseType == "Metadata + Data usage")
            {
                fieldPropretiesView.Columns[45].HeaderText = "Percentage Of Use %";
                fieldPropretiesView.Columns[45].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                fieldPropretiesView.Columns[45].DefaultCellStyle.Format = "N2";
                fieldPropretiesView.Columns[45].Frozen = false;
                foreach (DataGridViewRow item in fieldPropretiesView.Rows)
                {
                    if (item.Cells[45].Value != null &&  item.Cells[45].Value.ToString() == "0")
                    {
                        foreach (DataGridViewCell t in item.Cells)
                        {
                            t.Style.BackColor = Color.LightGray;
                        }
                    }
                }
            }

            for (int i = 0; i <= 44; i++)
            {
                fieldPropretiesView.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                fieldPropretiesView.Columns[i].Frozen = false;
            }
            

            fieldPropretiesView.Columns[2].Visible = false;
            fieldPropretiesView.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            fieldPropretiesView.Columns[3].Width = 300;

            fieldPropretiesView.Columns[4].Visible                 = false;

            if (!shouldDisplayAllColumns)
            {
                var columnsCount = analyseType == "Metadata + Data usage" ? fieldPropretiesView.Columns.Count - 1 : fieldPropretiesView.Columns.Count;

                for (int i = 13; i < columnsCount; i++)
                {
                    fieldPropretiesView.Columns[i].Visible = false;
                }
            }

            if (selectedType == AttributeTypeCode.Lookup.ToString() || selectedType == AttributeTypeCode.Owner.ToString() || selectedType == "ALL")
                fieldPropretiesView.Columns[4].Visible = true;
            if (selectedType == "ALL")
                fieldPropretiesView.Columns[2].Visible = true;
        }

        #endregion
        #region SetEntitiesGridViewHeaders
        public static void SetEntitiesGridViewHeaders(DataTable dt, DataGridView entityGridView)
        {
            entityGridView.DataSource = dt;
            entityGridView.RowHeadersVisible = false;
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

        public static void SetFieldDataGridViewContent(List<entityParam> entities, DataTable dtFields, String analyseType) {

            foreach (var item in entities)
            {
                DataRow row = dtFields.NewRow();

                row["DisplayName"]         = item.displayName;
                row["SchemaName"]          = item.fieldName;
                row["Type"]                 = item.attributeType;
                row["Description"]          = item.Description;
                row["Target"]               = item.target;
                row["Managed/Unmanaged"]    = item.isManaged;
                row["IsCustom"]              = item.isCustom;
                row["IsAuditable"]          = item.isAuditable;
                row["IsSearchable"]         = item.isSearchable;
                row["Required Level"]       = item.requiredLevel;
                row["Introduced Version"]   = item.introducedVersion;
                row["CreatedOn"]            = item.dateOfCreation;
                row["ModifiedOn"]           = item.ModifiedOn;

                if (analyseType == "Metadata + Data usage")
                    row["Percentage Of Use"]    = item.percentageOfUse;

                row["AttributeOf"]          = item.AttributeOf;
                row["AutoNumberFormat"]      = item.AutoNumberFormat;
                row["CanBeSecuredForCreate"] = item.CanBeSecuredForCreate;
                row["CanBeSecuredForRead"]      = item.CanBeSecuredForRead;
                row["CanBeSecuredForUpdate"]        = item.CanBeSecuredForUpdate;
                row["CanModifyAdditionalSettings"] = item.CanModifyAdditionalSettings;
                row["ColumnNumber"] = item.ColumnNumber;
                row["DeprecatedVersion"] = item.DeprecatedVersion;
                row["ExternalName"] = item.ExternalName;
                row["InheritsFrom"] = item.InheritsFrom;
                row["IsCustomizable"] = item.IsCustomizable;
                row["IsDataSourceSecret"] = item.IsDataSourceSecret;
                row["IsFilterable"] = item.IsFilterable;
                row["IsGlobalFilterEnabled"] = item.IsGlobalFilterEnabled;
                row["IsLogical"] = item.IsLogical;
                row["IsPrimaryId"] = item.IsPrimaryId;
                row["IsPrimaryName"] = item.IsPrimaryName;
                row["IsRenameable"] = item.IsRenameable;
                row["IsRequiredForForm"] = item.IsRequiredForForm;
                row["IsRetrievable"] = item.IsRetrievable;
                row["IsSecured"] = item.IsSecured;
                row["IsSortableEnabled"] = item.IsSortableEnabled;
                row["IsValidForAdvancedFind"] = item.IsValidForAdvancedFind;
                row["IsValidForCreate"] = item.IsValidForCreate;
                row["IsValidForForm"] = item.IsValidForForm;
                row["IsValidForGrid"] = item.IsValidForGrid;
                row["IsValidForRead"] = item.IsValidForRead;
                row["IsValidForUpdate"] = item.IsValidForUpdate;
                row["IsValidODataAttribute"] = item.IsValidODataAttribute;
                row["LinkedAttributeId"] = item.LinkedAttributeId;
                row["EntityLogicalName"] = item.EntityLogicalName;
                row["SourceType"] = item.SourceType;

                dtFields.Rows.Add(row);
            }
        }

        public static void setStatisticsFieldText(Label statisticsText) {
            if (entityInfo.entityRecordsCount > 0)
                statisticsText.Text = "Statistics Based On " + entityInfo.entityRecordsCount + " Records";
            else
                statisticsText.Text = String.Empty;
        }

        public static void  CallExportFunction(Dictionary<AttributeTypeCode, List<entityParam>> entityParams, bool analyseType) {
            FileManaged.ExportFile(entityParams, entityInfo, analyseType);
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