using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using System.Diagnostics;

namespace EntityieldsAnalyser
{
    public partial class MyPluginControl : PluginControlBase
    {
        #region Variables
        private Settings mySettings;
        DataTable dtEntities;
        DataTable dtFields;
        Dictionary<AttributeTypeCode, List<entityParam>> entityFields = null;
        #endregion

        public MyPluginControl()
        {
            InitializeComponent();
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            #region ManageComponenetVisibility
            InitComponents();
            entityTypeComboBox.SelectedIndex = 0;
            #endregion
            //Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }
        }

        /// <summary>
        /// This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        /// <summary>
        /// This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
        }
        #region Load Entities Button
        private void getEntitiesButton_Click(object sender, EventArgs e)
        {
            ExecuteMethod(loadEntities);
        }
        #endregion
        #region Load Entities Function
        private void loadEntities()
        {
            InitComponents();
            string selectedTypeOfEntities = entityTypeComboBox.SelectedItem.ToString();
            searchEntity.Enabled = false;
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Retrieving Entities...",
                Work = (worker, args) =>
                {
                    #region Variables
                    dtEntities = new DataTable();
                    #endregion
                    #region getEntitiesMetadata
                    RetrieveMetadataChangesResponse _allEntitiesResp = EntityFieldAnalyserManager.GetEntitiesMetadat(Service, selectedTypeOfEntities);
                    #endregion

                    worker.ReportProgress(0, string.Format("Metadata has been retrieved!"));

                    #region Entities Data Table Set
                    dtEntities.Columns.Add("DisplayName", typeof(string));
                    dtEntities.Columns.Add("SchemaName", typeof(string));
                    dtEntities.Columns.Add("Analyse", typeof(bool));

                    foreach (var item in _allEntitiesResp.EntityMetadata)
                    {
                        DataRow row        = dtEntities.NewRow();
                        row["DisplayName"] = item.DisplayName.LocalizedLabels.Count > 0 ? item.DisplayName.UserLocalizedLabel.Label.ToString() : null;
                        row["SchemaName"]  = item.LogicalName;
                        row["Analyse"]     = false;

                        dtEntities.Rows.Add(row);
                    }
                    #endregion
                    args.Result = dtEntities;
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result;
                    if (result != null)
                    {
                        #region Set Retrieved Data in the Data Grid View
                        EntityFieldAnalyserManager.SetEntitiesGridViewHeaders((DataTable)args.Result, EntityGridView);
                        #endregion
                        #region ManageComponenetVisibility
                        searchEntity.Enabled  = true;
                        analyseButton.Enabled = true;
                        #endregion
                    }
                }
            });
        }
        #endregion
        #region Search For Entity
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (searchEntity.Text == "" && dtEntities.Rows.Count > 0 && EntityGridView.Rows.Count != dtEntities.Rows.Count)
            {
                EntityFieldAnalyserManager.SetEntitiesGridViewHeaders(dtEntities, EntityGridView);
            }
            else if (searchEntity.Text != "Search" && searchEntity.Text != "" && dtEntities.Rows.Count > 0)
            {
                string searchValue = searchEntity.Text.ToLower();
                try
                {
                    DataRow[] filtered = dtEntities.Select("DisplayName LIKE '%" + searchValue + "%' OR SchemaName LIKE '%" + searchValue + "%'");
                    if (filtered.Count() > 0)
                    {
                        EntityFieldAnalyserManager.SetEntitiesGridViewHeaders(filtered.CopyToDataTable(), EntityGridView);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Invalid Search Character. Please do not use ' [ ] within searches.");
                }
            }
        }
        #endregion

        #region Delete Text When First Click
        private void textBox1_Click(object sender, EventArgs e)
        {
            if (searchEntity.Text == "Search")
            {
                searchEntity.Clear();
            }
        }
        #endregion
        #region Analyse Button
        private void analyseButton_Click(object sender, EventArgs e)
        {
            ExecuteMethod(AnalyseEntity);
        }
        #endregion
        #region Analyse Entity Function
        private void AnalyseEntity()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Analyse Entity ...",
                Work = (worker, args) =>
                {
                    try
                    {
                        #region Variables
                        dtFields                       = new DataTable();
                        DataGridViewRowCollection rows = EntityGridView.Rows;
                        bool isEntitySelected          = false;
                        #endregion
                        #region Analyse The Selected Entity
                        foreach (DataGridViewRow row in rows)
                        {
                            var cell = row.Cells[2];
                            if (cell != null)
                            {
                                if ((bool)cell.Value != null && (bool)cell.Value == true)
                                {
                                    isEntitySelected            = true;
                                    string entityName           = row.Cells[0].Value.ToString();
                                    string entityTechnicalName  = row.Cells[1].Value.ToString();
                                    entityFields                = EntityFieldAnalyserManager.getEntityFields(Service, entityTechnicalName, entityName);
                                }
                            }
                        }

                        if (!isEntitySelected)
                        {
                            MessageBox.Show("No Entity selected! Please Select an Entity", "Warning");
                            return;
                        }
                        #endregion

                        #region Entity Fiels Metadata  Set
                        dtFields.Columns.Add("Display Name", typeof(string));
                        dtFields.Columns.Add("Schema Name", typeof(string));
                        dtFields.Columns.Add("Managed/Unmanaged", typeof(string));
                        dtFields.Columns.Add("IsAuditable", typeof(bool));
                        dtFields.Columns.Add("IsSearchable", typeof(bool));
                        dtFields.Columns.Add("Required Level", typeof(string));
                        dtFields.Columns.Add("Introduced Version", typeof(String));
                        dtFields.Columns.Add("CreatedOn", typeof(DateTime));
                        dtFields.Columns.Add("Target", typeof(string));
                        dtFields.Columns.Add("Percentage Of Use", typeof(string));
                        #endregion

                        args.Result = entityFields;
                    }
                    catch (Exception e) {
                        MessageBox.Show(e.Message, "Warning");
                    }
                },
                ProgressChanged = e =>
                {
                    // If progress has to be notified to user, use the following method:
                    SetWorkingMessage(e.UserState.ToString());
                },
                PostWorkCallBack = (args) =>
                {
                    fieldTypeCombobox.Items.Clear();
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = (Dictionary<AttributeTypeCode, List<entityParam>>)args.Result;
                    if (result != null)
                    {
                        foreach (var type in result.Keys)
                        {
                            fieldTypeCombobox.Items.Add(type);
                        }
                        #region Set Charts Data
                        EntityFieldAnalyserManager.setStatisticsFieldText(label7);
                        EntityFieldAnalyserManager.SetChartFieldsType(result, ChartFieldTypes);
                        EntityFieldAnalyserManager.SetChartFieldAvailable(result, ChartFieldAvailable);
                        EntityFieldAnalyserManager.SetChartManagedUnmanagedFields(chartManagedUnmanagedFields);
                        #endregion
                        #region Manage Componenet Visibilty
                        searchEntity.Enabled                   = true;
                        DisplayPercentageCheckbox.Visible      = true;
                        fieldTypeCombobox.Enabled              = true;
                        DisplayPercentageCheckbox.Checked      = false;
                        EntityFieldsCreatedGroupBox.Visible    = true;
                        ManagedUnmanagedFieldsgroupBox.Visible = true;
                        EntityFieldsTypesGroupBox.Visible      = true;
                        fieldCalculatorGroupBox.Visible        = true;
                        buttonExport.Enabled                   = true;
                        #endregion
                        fieldTypeCombobox.SelectedIndex        = 1;//Select the second type in the fields type combobox
                        #region Chart ToolTip
                        ToolTip toolTip = new ToolTip();
                        toolTip.SetToolTip(this.EntityFieldsCreatedGroupBox, "This Chart Display The Percentage of Use of the Entity Total Fields Volume");
                        toolTip.SetToolTip(this.ManagedUnmanagedFieldsgroupBox, "This Chart Display The Count Managned/Unmanaged Fields in Your Entity");
                        toolTip.SetToolTip(this.EntityFieldsTypesGroupBox, "This Chart Display The Count of Fields Per Type");
                        #endregion
                    }
                }
            });
        }
        #endregion
        #region Display Field based On Selected Type
        private void fieldTypeCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region clear grid view
            DataTable DT = (DataTable)fieldPropretiesView.DataSource;
            if (DT != null)
                DT.Clear();
            #endregion
            #region Set Fields DataGridView Data After Change The Type
            EntityFieldAnalyserManager.SetFieldDataGridViewContent(entityFields[(AttributeTypeCode)fieldTypeCombobox.SelectedItem], dtFields);
            EntityFieldAnalyserManager.SetFieldDataGridViewHeaders((DataTable)dtFields, fieldPropretiesView);
            //Display the Column Target Just in case of Lookup Type
            if (fieldTypeCombobox.SelectedItem.ToString() == AttributeTypeCode.Lookup.ToString())
            {
                fieldPropretiesView.Columns[8].Visible = true;
            }
            #endregion
        }
        #endregion

        #region Allow Select of One Entity
        private void EntityGridView_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                foreach (DataGridViewRow row in this.EntityGridView.Rows)
                    row.Cells["Analyse"].Value = false;

                this.EntityGridView.Rows[e.RowIndex].Cells["Analyse"].Value = true;
            }
        }
        #endregion
        #region Change Chart Content (Percentage => Record Count)
        private void DisplayPercentageCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if(DisplayPercentageCheckbox.Checked == true)
            {
                EntityFieldAnalyserManager.SetChartFieldsType(entityFields, ChartFieldTypes, false);
                EntityFieldAnalyserManager.SetChartFieldAvailable(entityFields, ChartFieldAvailable, true);
                EntityFieldAnalyserManager.SetChartManagedUnmanagedFields(chartManagedUnmanagedFields, false);
            }
            else
            {
                EntityFieldAnalyserManager.SetChartFieldsType(entityFields, ChartFieldTypes, true);
                EntityFieldAnalyserManager.SetChartFieldAvailable(entityFields, ChartFieldAvailable, false);
                EntityFieldAnalyserManager.SetChartManagedUnmanagedFields(chartManagedUnmanagedFields, true);
            }
        }

        private void InitComponents() {
            if(fieldPropretiesView.DataSource != null)
                fieldPropretiesView.DataSource = null;
            if (fieldTypeCombobox.Items.Count > 0)
                fieldTypeCombobox.Items.Clear();
            if(chartManagedUnmanagedFields.Series["managedUnmanagedFields"].Points.Count > 0)
                chartManagedUnmanagedFields.Series["managedUnmanagedFields"].Points.Clear();
            if (ChartFieldAvailable.Series["AvailableField"].Points.Count > 0)
                ChartFieldAvailable.Series["AvailableField"].Points.Clear();
            if (ChartFieldTypes.Series["fieldsReport"].Points.Count > 0)
                ChartFieldTypes.Series["fieldsReport"].Points.Clear();

            fieldTypeCombobox.Enabled                = false;
            EntityFieldsCreatedGroupBox.Visible      = false;
            ManagedUnmanagedFieldsgroupBox.Visible   = false;
            EntityFieldsTypesGroupBox.Visible        = false;
            DisplayPercentageCheckbox.Visible        = false;
            searchEntity.Enabled                     = false;
            analyseButton.Enabled                    = false;
            fieldTypeCombobox.Enabled                = false;
            fieldCalculatorGroupBox.Visible          = false;
            buttonExport.Enabled                     = false;
        }
        #endregion

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool isItPossibleToCreateThisFields = EntityFieldAnalyserManager.CanICreateThisNumberOfFields(textBox1.Text, textBox2.Text, textBox3.Text);
            if (isItPossibleToCreateThisFields)
                MessageBox.Show("You Can Create This Fields", "Success");
            else
                MessageBox.Show("You Can't Create This Number of Fields", "Warning");
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            EntityFieldAnalyserManager.CallExportFunction(entityFields);
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void label6_Click(object sender, EventArgs e)
        {
            Process.Start("https://hamzaamraoui.medium.com/field-limits-in-dynamics-365-how-many-fields-is-too-many-fields-ab39c699336e");
        }

        private void toolStripMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            string message = "";

            message += "1. Select an Entity Filter and click Load Entities";
            message += Environment.NewLine;
            message += "2. Click the checkbox on any entity that you would like to Analyse";
            message += Environment.NewLine;
            message += "3. Search is wildcard already so you only need to type in the text you want to search for";
            message += Environment.NewLine;
            message += "4. Click Analyse when Ready";
            message += Environment.NewLine;
            message += "5. Click on Export to Save the Results in Excel File";
            message += Environment.NewLine;
            message += Environment.NewLine;
            message += "If you have any issues please log them via GitHub and/or contact me at hamzamraoui11@gmail.com";

            MessageBox.Show(message);
        }

        private void byButton_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.linkedin.com/in/hamza-amraoui/");
        }
    }
}