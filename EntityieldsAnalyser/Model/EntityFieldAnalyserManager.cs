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
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Label = System.Windows.Forms.Label;

namespace EntityieldsAnalyser
{
    public enum EntityType { All, Custom, Standard, Filter }

    class EntityFieldAnalyserManager
    {
        #region Variables
        private IOrganizationService _service;
        private IEnumerable<EntityMetadata> _metadataList;
        public static int _managedFieldsCount = 0;
        public static int _unmanagedFieldsCount = 0;
        public static int _customField = 0;
        public static int _standardField = 0;
        public static int _currentUseOfColumns = 0;
        public static int _entityDefaultColumnSize = 1024;
        public static int _totalRecordsEntity = 0;
        public static EntityInfo entityInfo = null;

        #endregion
        #region Get EntityFields
        public static Dictionary<AttributeTypeCode, List<entityParam>> getEntityFields(IOrganizationService service, String entityName)
        {
            entityInfo = new EntityInfo();
            Dictionary<AttributeTypeCode, List<entityParam>> _data = new Dictionary<AttributeTypeCode, List<entityParam>>();
            RetrieveEntityRequest retrieveBankAccountEntityRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = entityName
            };
            RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.Execute(retrieveBankAccountEntityRequest);
            _data = formatList(retrieveEntityResponse, _data);
            EntityCollection _entityRecords = getEntityRecords(service, entityName);
            _totalRecordsEntity = _entityRecords.Entities.Count;
            entityInfo.entityName = entityName;
            entityInfo.numberOfFields = retrieveEntityResponse.EntityMetadata.Attributes.Count();
            entityInfo.numberOfRecords = _totalRecordsEntity;
            _unmanagedFieldsCount = _managedFieldsCount = _customField= _standardField= _currentUseOfColumns = 0;

            return setDictionaryCount(_entityRecords, _data); ;
        }
        #endregion
        #region formatList function
        // ordering result data to a dictionary
        private static Dictionary<AttributeTypeCode, List<entityParam>> formatList(RetrieveEntityResponse _metadata, Dictionary<AttributeTypeCode, List<entityParam>> _data)
        {
            foreach (var field in _metadata.EntityMetadata.Attributes)
            {
                if (!_data.ContainsKey(field.AttributeType.Value))
                {
                    _data.Add(field.AttributeType.Value, new List<entityParam>() { setObject(field)});
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
                fieldName = field.LogicalName,
                isManaged = (bool)field.IsManaged,
                target = field.AttributeType.Value == AttributeTypeCode.Lookup && ((LookupAttributeMetadata)field).Targets.Length > 0 ? ((LookupAttributeMetadata)field).Targets[0] : String.Empty,
                dateOfCreation = field.CreatedOn.Value.Date,
                introducedVersion = field.IntroducedVersion,
                isAuditable = field.IsAuditEnabled.Value,
                requiredLevel = ((AttributeRequiredLevel)field.RequiredLevel.Value).ToString(),
                isSearchable = (bool)field.IsSearchable,
                isCustom = (bool)field.IsCustomAttribute,
                totalFiledRecords = 0
            };
        }
        #endregion
        #region Get Entity Records
        private static EntityCollection getEntityRecords(IOrganizationService service, String entityName)
        {
            EntityCollection entities = new EntityCollection();
            QueryExpression query = new QueryExpression(entityName);
            query.ColumnSet = new ColumnSet(true);
            var PageCookie = String.Empty;
            var PageNumber = 1;
            var PageSize = 5000;

            EntityCollection result;
            do
            {
                query.PageInfo = new PagingInfo() { PageNumber = 1, Count = PageSize };

                if (PageNumber != 1)
                {
                    query.PageInfo.PageNumber = PageNumber;
                    query.PageInfo.PagingCookie = PageCookie;
                }
                result = service.RetrieveMultiple(query);

                if (result.MoreRecords)
                {
                    entities.Entities.AddRange(result.Entities);
                    PageNumber++;
                    PageCookie = result.PagingCookie;
                }
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
                    if (entityParam.isManaged)
                        _managedFieldsCount++;
                    else
                        _unmanagedFieldsCount++;

                   //For Custom Standard Chart
                    if (entityParam.isCustom)
                        _customField++;
                    else
                        _standardField++;

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
                        field.percentageOfUse = String.Format("{0:P2}.", ((double)(field.totalFiledRecords * 1.0 / recordsCount)));//String.Format("Value: {0:P2}.", 0.8526) 
                    else
                        field.percentageOfUse = 0 + "%";
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
                ChartFieldTypes.Series["fieldsReport"].LabelFormat = "";
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
            _currentUseOfColumns = DbUsedColumns(dict);
            var availableFieldToCreate = _entityDefaultColumnSize - _currentUseOfColumns;

            if (!displayPercentage)
            {
                ChartFieldAvailabe.Series["AvailableField"].LabelFormat = "";
                ChartFieldAvailabe.Series["AvailableField"].IsValueShownAsLabel = false;
            }
            else//enable formating to percentage when check the checkbox
            {
                ChartFieldAvailabe.Series["AvailableField"].LabelFormat = "0.#%";
                ChartFieldAvailabe.Series["AvailableField"].IsValueShownAsLabel = true;
            }

            #region Clear Chart Data
            if (ChartFieldAvailabe.Series["AvailableField"].Points.Count > 0) {
                ChartFieldAvailabe.Series["AvailableField"].Points.Clear();
            }
            #endregion
            #region set chart Data
            ChartFieldAvailabe.Series["AvailableField"].Points.AddXY("Aailable Fields To Create ", displayPercentage ? availableFieldToCreate * 1.0 / _entityDefaultColumnSize : availableFieldToCreate);
            ChartFieldAvailabe.Series["AvailableField"].Points.AddXY("Created Fields", displayPercentage ? _currentUseOfColumns * 1.0 / _entityDefaultColumnSize : _currentUseOfColumns);
            ChartFieldAvailabe.Dock = DockStyle.Fill;
            #endregion
        }
        #endregion
        #region SetFieldDataGridViewHeaders
        public static void SetFieldDataGridViewHeaders(DataTable dt, DataGridView fieldPropretiesView)
        {
            fieldPropretiesView.DataSource = dt;
            fieldPropretiesView.Sort(fieldPropretiesView.Columns[1], ListSortDirection.Ascending);
            fieldPropretiesView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Regular);
            fieldPropretiesView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[0].HeaderText = "DisplayName";
            fieldPropretiesView.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[1].HeaderText = "Managed/Unmanaged";
            fieldPropretiesView.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[2].HeaderText = "IsAuditable";
            fieldPropretiesView.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[3].HeaderText = "IsSearchable";
            fieldPropretiesView.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[4].HeaderText = "Required Level";
            fieldPropretiesView.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[5].HeaderText = "Introduced Version";
            fieldPropretiesView.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[6].HeaderText = "CreatedOn";
            fieldPropretiesView.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[7].HeaderText = "Target";
            fieldPropretiesView.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            fieldPropretiesView.Columns[8].HeaderText = "Percentage Of Use";
            fieldPropretiesView.Columns[8].DefaultCellStyle.Format = "0\\%";
            fieldPropretiesView.ReadOnly = true;
            fieldPropretiesView.Columns[7].Visible = false;
        }
        #endregion
        #region SetEntitiesGridViewHeaders
        public static void SetEntitiesGridViewHeaders(DataTable dt, DataGridView entityGridView)
        {
            entityGridView.DataSource = dt;
            entityGridView.Sort(entityGridView.Columns[1], ListSortDirection.Ascending);
            entityGridView.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9.75F, FontStyle.Regular);
            entityGridView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            entityGridView.Columns[0].HeaderText = "Display Name";
            entityGridView.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            entityGridView.Columns[1].HeaderText = "Schema Name";
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
                Criteria = entityFilter
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
                managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].LabelFormat = "";
            else//enable formating to percentage when check the checkbox
                managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].LabelFormat = "0.#%";

            managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].Points.AddXY("Managed Fields", displayValues ? _managedFieldsCount : (_managedFieldsCount * 1.0 / (_managedFieldsCount+_unmanagedFieldsCount)));
            managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].Points.AddXY("Unmanaged Fields", displayValues ? _unmanagedFieldsCount : (_unmanagedFieldsCount * 1.0 / (_managedFieldsCount + _unmanagedFieldsCount)));
            managedUnmanagedFieldsChart.Series["managedUnmanagedFields"].IsValueShownAsLabel = true;
            managedUnmanagedFieldsChart.Dock = DockStyle.Fill;
        }

        public static bool CanICreateThisNumberOfFields(String lookupsType, String pickListTypes, String othersTypes)
        {
            int askedCreatedFieldsColumnsSize = (int.Parse(lookupsType) * 3) + (int.Parse(pickListTypes) * 2) + (int.Parse(othersTypes) * 1);
            int availableFieldsToCreate = _entityDefaultColumnSize - _currentUseOfColumns;
            if (availableFieldsToCreate > askedCreatedFieldsColumnsSize)
                return true;
            else
                return false;
        }

        public static void SetFieldDataGridViewContent(List<entityParam> entities, DataTable dtFields) {

            foreach (var item in entities)
            {
                DataRow row = dtFields.NewRow();

                row["DisplayName"] = item.fieldName;
                row["Managed/Unmanaged"] = item.isManaged == true ? "Managed" : "Unmanaged";
                row["IsAuditable"] = item.isAuditable;
                row["IsSearchable"] = item.isSearchable;
                row["Required Level"] = item.requiredLevel;
                row["Introduced Version"] = item.introducedVersion;
                row["CreatedOn"] = item.dateOfCreation;
                row["Percentage Of Use"] = item.percentageOfUse;
                row["Target"] = item.target;

                dtFields.Rows.Add(row);
            }

        }

        public static void setStatisticsFieldText(Label statisticsText) {
            if (_totalRecordsEntity > 0)
                statisticsText.Text = "Statistics Based On " + _totalRecordsEntity + " Records";
            else
                statisticsText.Text = "";
            _totalRecordsEntity = 0;
        }

        public static SaveFileDialog  CallExportFunction(Dictionary<AttributeTypeCode, List<entityParam>> entityParams) {
            return FileManaged.ExportFile(entityParams, entityInfo, new int[] { _managedFieldsCount, _unmanagedFieldsCount }, new int[] { _customField, _standardField }, new int[] { _currentUseOfColumns , _entityDefaultColumnSize });
        }
    }
    #region Class EntityParam
    public class entityParam
    {
        public string fieldName;
        public DateTime dateOfCreation;
        public string introducedVersion;
        public bool isManaged;
        public bool isAuditable;
        public string requiredLevel;
        public bool isSearchable;
        public bool isOnForm;
        public bool isCustom;
        public string target;
        public int totalFiledRecords;
        public string percentageOfUse;
    }
    #endregion
    public class EntityInfo
    {
        public string entityName;
        public int numberOfFields;
        public int numberOfRecords;
    }
}