using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using XrmToolBox.Extensibility;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using System.Diagnostics;
using EntityieldsAnalyser.UI;

namespace EntityieldsAnalyser
{
    public partial class MyPluginControl : PluginControlBase
    {
        #region Variables
        private Settings mySettings;
        DataTable dtEntities;
        DataTable dtFields;
        Dictionary<AttributeTypeCode, List<entityParam>> entityFields = null;
        String entitySelectedName = String.Empty;
        String entitySelectedSchemaName = String.Empty;
        #endregion

        public MyPluginControl()
        {
            InitializeComponent();
            BuildInsightsUi();
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            #region ManageComponenetVisibility
            ApplyInterfaceTheme();
            InitComponents();
            entityTypeComboBox.SelectedIndex = 0;
            AnalyseType.SelectedIndex = 0;
            UpdateGuide();
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

        #region Interface Theme
        private void ApplyInterfaceTheme()
        {
            BackColor = Theme.Canvas;
            ForeColor = Theme.TextPrimary;

            #region Command bar
            toolStripMenu.Renderer = new FluentToolStripRenderer();
            toolStripMenu.BackColor = Theme.Surface;
            foreach (ToolStripItem item in toolStripMenu.Items)
            {
                item.Font = Theme.Body;
                item.ForeColor = Theme.TextPrimary;
            }
            analyseButton.Font = Theme.BodyStrong;

            tsbClose.Image = Theme.Icon(Theme.Glyph.Close, Theme.TextSecondary);
            toolStripButton.Image = Theme.Icon(Theme.Glyph.List, Theme.Brand);
            analyseButton.Image = Theme.Icon(Theme.Glyph.Play, Theme.TextOnBrand);
            buttonExport.Image = Theme.Icon(Theme.Glyph.Export, Color.FromArgb(16, 124, 65));
            byButton.Image = Theme.Icon(Theme.Glyph.Contact, Theme.TextSecondary);
            helpButton.Image = Theme.Icon(Theme.Glyph.Help, Theme.TextSecondary);
            buymeacoffeeicon.Image = Theme.Icon(Theme.Glyph.Heart, Color.FromArgb(196, 49, 75));
            #endregion

            #region Labels
            foreach (var label in new[] { label2, label3, label4 })
            {
                label.ForeColor = Theme.TextSecondary;
                label.Font = Theme.BodyStrong;
            }
            foreach (var label in new[] { label5, fieldsResultLabel, guideHint, featureOverviewText, featureFieldsText, featureUsageText, featureCapacityText, usageEmptyText })
                label.ForeColor = Theme.TextSecondary;
            welcomeTitle.Font = featuresTitle.Font = Theme.Subtitle;
            welcomeTitle.ForeColor = featuresTitle.ForeColor = Theme.TextPrimary;
            budgetHeadline.Font = new Font("Segoe UI Semibold", 13F);
            MoreDetailsLink.ForeColor = Theme.Brand;
            MoreDetailsLink.MouseEnter += (s, e) => MoreDetailsLink.Font = new Font(Theme.Body, FontStyle.Underline);
            MoreDetailsLink.MouseLeave += (s, e) => MoreDetailsLink.Font = Theme.Body;
            calculatorResult.Font = Theme.BodyStrong;
            #endregion

            #region Grids
            GridStyler.Apply(EntityGridView,
                () => String.IsNullOrEmpty(searchEntity.Text) ? "Choose a filter and click “Get Entities” to list your entities." : "No entity matches your search.",
                PaintEntityCell);
            EntityGridView.RowTemplate.Height = Theme.Scale(46);
            EntityGridView.DataBindingComplete += (s, e) => EntityGridView.ClearSelection();

            GridStyler.Apply(fieldPropretiesView,
                () => _insights == null ? "Select an entity, then click “Analyse” to inspect its fields." : "No field matches these filters. Remove a filter above to see more fields.");
            #endregion
        }

        /// <summary>Entities grid: radio style selector and two line "display name / schema name" cell.</summary>
        private bool PaintEntityCell(DataGridViewCellPaintingEventArgs e)
        {
            var column = EntityGridView.Columns[e.ColumnIndex];
            var selected = (e.State & DataGridViewElementStates.Selected) != 0;
            var bounds = e.CellBounds;

            if (column.DataPropertyName == "Analyse")
            {
                e.PaintBackground(bounds, selected);
                var isChecked = e.Value is bool && (bool)e.Value;
                var size = Theme.Scale(16);
                var r = new RectangleF(bounds.X + (bounds.Width - size) / 2f, bounds.Y + (bounds.Height - size) / 2f, size, size);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                if (isChecked)
                {
                    using (var brush = new SolidBrush(Theme.Brand))
                        e.Graphics.FillEllipse(brush, r);
                    var dot = Theme.Scale(6);
                    using (var brush = new SolidBrush(Theme.Surface))
                        e.Graphics.FillEllipse(brush, r.X + (size - dot) / 2f, r.Y + (size - dot) / 2f, dot, dot);
                }
                else
                {
                    using (var pen = new Pen(Theme.TextSecondary))
                        e.Graphics.DrawEllipse(pen, r);
                }
                e.Handled = true;
                return true;
            }

            if (column.DataPropertyName == "DisplayName" && EntityGridView.Columns.Contains("SchemaName"))
            {
                e.PaintBackground(bounds, selected);
                var schema = Convert.ToString(EntityGridView.Rows[e.RowIndex].Cells["SchemaName"].Value);
                var textBounds = new Rectangle(bounds.X + Theme.Scale(4), bounds.Y, bounds.Width - Theme.Scale(8), bounds.Height);
                var half = bounds.Height / 2;
                TextRenderer.DrawText(e.Graphics, Convert.ToString(e.Value), Theme.BodyStrong,
                    new Rectangle(textBounds.X, textBounds.Y + Theme.Scale(4), textBounds.Width, half - Theme.Scale(3)), Theme.TextPrimary,
                    Draw.LeftMiddle);
                TextRenderer.DrawText(e.Graphics, schema, Theme.Caption,
                    new Rectangle(textBounds.X, textBounds.Y + half - Theme.Scale(1), textBounds.Width, half - Theme.Scale(4)), Theme.TextTertiary,
                    Draw.LeftMiddle);
                e.Handled = true;
                return true;
            }

            return false;
        }

        private static string Percent(int part, int whole) => whole <= 0 ? "0%" : ((double)part / whole).ToString("0%");
        #endregion

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
            UpdateHeader();
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
                    var entities = new DataTable();
                    #endregion
                    #region getEntitiesMetadata
                    RetrieveMetadataChangesResponse _allEntitiesResp = EntityFieldAnalyserManager.GetEntitiesMetadat(Service, selectedTypeOfEntities);
                    #endregion

                    worker.ReportProgress(0, string.Format("Metadata has been retrieved!"));

                    #region Entities Data Table Set
                    entities.Columns.Add("DisplayName", typeof(string));
                    entities.Columns.Add("SchemaName", typeof(string));
                    entities.Columns.Add("Analyse", typeof(bool));

                    foreach (var item in _allEntitiesResp.EntityMetadata)
                    {
                        DataRow row        = entities.NewRow();
                        row["DisplayName"] = item.DisplayName.LocalizedLabels.Count > 0 ? item.DisplayName.UserLocalizedLabel.Label.ToString() : "N/A";
                        row["SchemaName"]  = item.LogicalName;
                        row["Analyse"]     = false;

                        entities.Rows.Add(row);
                    }
                    #endregion
                    args.Result = entities;
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
                    var result = args.Result as DataTable;
                    if (result != null)
                    {
                        dtEntities = result;
                        #region Set Retrieved Data in the Data Grid View
                        EntityFieldAnalyserManager.SetEntitiesGridViewHeaders(dtEntities, EntityGridView);
                        groupBox1.Subtitle = $"{dtEntities.Rows.Count:N0} · {selectedTypeOfEntities} · double-click to analyse";
                        #endregion
                        #region ManageComponenetVisibility
                        searchEntity.Enabled  = true;
                        analyseButton.Enabled = true;
                        AnalyseType.Enabled   = true;
                        #endregion
                    }
                    UpdateGuide();
                }
            });
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
            String[] entitySelected = EntityFieldAnalyserManager.SelectedEntity(EntityGridView.Rows);
            if (entitySelected.Length > 0) {
                searchEntity.Enabled          = false;
                analyseButton.Enabled         = false;
                AnalyseType.Enabled           = false;
                buttonExport.Enabled          = false;
                entityTypeComboBox.Enabled    = false;
                toolStripButton.Enabled       = false;
                entitySelectedName = entitySelected[0];
                entitySelectedSchemaName = entitySelected[1];
                ExecuteMethod(AnalyseEntity);
            }
            else
            {
                MessageBox.Show("No Entity Was Selected! Please Select an Entity", "Warning");
                return;
            }
        }
        #endregion
        #region Analyse Entity Function
        private void AnalyseEntity()
        {
            fieldTypeCombobox.Enabled = false;
            searchField.Enabled = false;
            buttonExport.Enabled = false;
            analyseButton.Enabled = false;
            displayAllColumns.Enabled = false;
            var analyseType = AnalyseType.SelectedItem.ToString();
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Analysing ...",
                Work = (worker, args) =>
                {
                    try
                    {
                        dtFields      = new DataTable();
                        entityFields  = EntityFieldAnalyserManager.getEntityFields(Service, entitySelectedSchemaName, entitySelectedName, worker, analyseType);

                        #region Entity Fiels Metadata  Set
                        dtFields.Columns.Add("DisplayName", typeof(string));
                        dtFields.Columns.Add("SchemaName", typeof(string));
                        dtFields.Columns.Add("Type", typeof(string));
                        dtFields.Columns.Add("Description", typeof(string));
                        dtFields.Columns.Add("Target", typeof(string));
                        dtFields.Columns.Add("Managed/Unmanaged", typeof(string));
                        dtFields.Columns.Add("IsCustom", typeof(bool));
                        dtFields.Columns.Add("IsAuditable", typeof(bool));
                        dtFields.Columns.Add("IsSearchable", typeof(bool));
                        dtFields.Columns.Add("Required Level", typeof(string));
                        dtFields.Columns.Add("Introduced Version", typeof(String));
                        dtFields.Columns.Add("CreatedOn", typeof(DateTime));
                        dtFields.Columns.Add("ModifiedOn", typeof(DateTime));
                        dtFields.Columns.Add("AttributeOf", typeof(string));
                        dtFields.Columns.Add("AutoNumberFormat", typeof(string));
                        dtFields.Columns.Add("CanBeSecuredForCreate", typeof(bool));
                        dtFields.Columns.Add("CanBeSecuredForRead", typeof(bool));
                        dtFields.Columns.Add("CanBeSecuredForUpdate", typeof(bool));
                        dtFields.Columns.Add("CanModifyAdditionalSettings", typeof(bool));
                        dtFields.Columns.Add("ColumnNumber", typeof(int));
                        dtFields.Columns.Add("DeprecatedVersion", typeof(string));
                        dtFields.Columns.Add("ExternalName", typeof(string));
                        dtFields.Columns.Add("InheritsFrom", typeof(string));
                        dtFields.Columns.Add("IsCustomizable", typeof(bool));
                        dtFields.Columns.Add("IsDataSourceSecret", typeof(bool));
                        dtFields.Columns.Add("IsFilterable", typeof(bool));
                        dtFields.Columns.Add("IsGlobalFilterEnabled", typeof(bool));
                        dtFields.Columns.Add("IsLogical", typeof(bool));
                        dtFields.Columns.Add("IsPrimaryId", typeof(bool));
                        dtFields.Columns.Add("IsPrimaryName", typeof(bool));
                        dtFields.Columns.Add("IsRenameable", typeof(bool));
                        dtFields.Columns.Add("IsRequiredForForm", typeof(bool));
                        dtFields.Columns.Add("IsRetrievable", typeof(bool));
                        dtFields.Columns.Add("IsSecured", typeof(bool));
                        dtFields.Columns.Add("IsSortableEnabled", typeof(bool));
                        dtFields.Columns.Add("IsValidForAdvancedFind", typeof(bool));
                        dtFields.Columns.Add("IsValidForCreate", typeof(bool));
                        dtFields.Columns.Add("IsValidForForm", typeof(bool));
                        dtFields.Columns.Add("IsValidForGrid", typeof(bool));
                        dtFields.Columns.Add("IsValidForRead", typeof(bool));
                        dtFields.Columns.Add("IsValidForUpdate", typeof(bool));
                        dtFields.Columns.Add("IsValidODataAttribute", typeof(bool));
                        dtFields.Columns.Add("LinkedAttributeId", typeof(string));
                        dtFields.Columns.Add("EntityLogicalName", typeof(string));
                        dtFields.Columns.Add("SourceType", typeof(string));
                        if (analyseType == DataUsageMode)
                            dtFields.Columns.Add("Percentage Of Use", typeof(double));

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
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result as Dictionary<AttributeTypeCode, List<entityParam>>;
                    if (result != null)
                    {
                        _analysedType = analyseType;
                        _insights = new FieldInsights(result, analyseType == DataUsageMode);
                        EntityFieldAnalyserManager.entityInfo.entityTotalUseOfColumns = _insights.ColumnsUsed;

                        #region Fields page filters
                        _suppressFieldFilter = true;
                        fieldTypeCombobox.Items.Clear();
                        fieldTypeCombobox.Items.Add("ALL");
                        foreach (var type in result.Keys.OrderBy(k => k.ToString()))
                            fieldTypeCombobox.Items.Add(type);
                        fieldTypeCombobox.SelectedIndex = 0;
                        searchField.Text = String.Empty;
                        if (_insights.HasDataUsage)
                            fieldsChips.SetChips("All", "Custom", "Unmanaged", "Required", "Never used");
                        else
                            fieldsChips.SetChips("All", "Custom", "Unmanaged", "Required");
                        fieldsChips.DismissibleText = null;
                        _drillFilter = null;
                        _suppressFieldFilter = false;
                        displayAllColumns.Checked = false;
                        ApplyFieldFilters();
                        #endregion

                        RenderInsights();
                        for (var p = 0; p < pivotTabs.Items.Count; p++)
                            pivotTabs.SetEnabled(p, true);
                        UpdateHeader();
                        ShowPage(Page.Overview);

                        fieldTypeCombobox.Enabled = true;
                        searchField.Enabled       = true;
                        displayAllColumns.Enabled = true;
                        buttonExport.Enabled      = true;
                    }

                    #region Manage Componenet Visibilty
                    searchEntity.Enabled       = true;
                    analyseButton.Enabled      = true;
                    AnalyseType.Enabled        = true;
                    entityTypeComboBox.Enabled = true;
                    toolStripButton.Enabled    = true;
                    #endregion
                    UpdateGuide();
                }
            });
        }
        #endregion
        #region Display Field based On Selected Type
        private void fieldTypeCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFieldFilters();
        }
        #endregion

        #region Allow Select of One Entity
        private void EntityGridView_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                foreach (DataGridViewRow row in this.EntityGridView.Rows)
                {
                    if ((bool)EntityGridView.Rows[e.RowIndex].Cells["Analyse"].Value)
                        continue;
                    row.Cells["Analyse"].Value = false;
                }

                this.EntityGridView.Rows[e.RowIndex].Cells["Analyse"].Value = !(bool)EntityGridView.Rows[e.RowIndex].Cells["Analyse"].Value;
                UpdateGuide();
            }
        }

        /// <summary>Double-click: select this entity and analyse it right away.</summary>
        private void EntityGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            foreach (DataGridViewRow row in EntityGridView.Rows)
                row.Cells["Analyse"].Value = row.Index == e.RowIndex;
            EntityGridView.InvalidateColumn(EntityGridView.Columns["Analyse"].Index);
            UpdateGuide();

            if (analyseButton.Enabled)
                analyseButton_Click(sender, e);
        }
        #endregion
        #region InitComponents
        private void InitComponents() {
            if(fieldPropretiesView.DataSource != null)
                fieldPropretiesView.DataSource = null;
            if (fieldTypeCombobox.Items.Count > 0)
                fieldTypeCombobox.Items.Clear();

            _insights                                = null;
            _drillFilter                             = null;
            fieldsChips.DismissibleText              = null;
            fieldTypeCombobox.Enabled                = false;
            searchEntity.Enabled                     = false;
            analyseButton.Enabled                    = false;
            searchField.Enabled                      = false;
            buttonExport.Enabled                     = false;
            AnalyseType.Enabled                      = false;
            fieldsResultLabel.Text                   = String.Empty;
            calculatorResult.Text                    = String.Empty;

            foreach (var tile in new[] { kpiFields, kpiCustom, kpiUnmanaged, kpiCapacity, kpiRecords, kpiUnused, kpiLowUsage, kpiAverage })
                tile.Reset();

            UpdateHeader();
            ShowWelcome();
            UpdateGuide();
        }
        #endregion
        #region Export
        private void buttonExport_Click(object sender, EventArgs e)
        {
            EntityFieldAnalyserManager.CallExportFunction(entityFields, _analysedType == DataUsageMode);
        }
        #endregion
        #region CloseTool
        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }
        #endregion
        #region MoreDetails
        private void label6_Click(object sender, EventArgs e)
        {
            Process.Start("https://hamzaamraoui.medium.com/field-limits-in-dynamics-365-how-many-fields-is-too-many-fields-ab39c699336e");
        }
        #endregion
        #region Help
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            string message = "";
            message += "We recommend to use the tool on big screen for best visibility, Below the steps to follow : ";
            message += Environment.NewLine;
            message += Environment.NewLine;
            message += "1. Select an Entity Filter and click Get Entities";
            message += Environment.NewLine;
            message += "2. Click the entity that you would like to Analyse (double-click analyses it right away)";
            message += Environment.NewLine;
            message += "3. Search is wildcard already so you only need to type in the text you want to search for";
            message += Environment.NewLine;
            message += "4. Select the type of analyse, MetadataOnly : if you want to get only the details of entity attributes. Metadata + Data usage : is for getting metadata details + data usage of each records on all database(it takes much time depends on your entity data size)";
            message += Environment.NewLine;
            message += "5. Click Analyse when Ready, then use the tabs: Overview, Fields, Data usage and Capacity";
            message += Environment.NewLine;
            message += "6. Click any bar, slice or tile to open the matching fields in the Fields tab. The blue chip shows the active filter, click its × to remove it";
            message += Environment.NewLine;
            message += "7. Rows highlighted in light orange have a percentage of use = 0%, that mean that the field is not contain data on all database";
            message += Environment.NewLine;
            message += "8. Click on Export to Save the Results in Excel File";
            message += Environment.NewLine;
            message += Environment.NewLine;
            message += "If you have any issues please log them via GitHub and/or contact me at hamzamraoui11@gmail.com";
            MessageBox.Show(message, "Help, Issue !!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }
        #endregion
        #region DevelopedBy
        private void byButton_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.linkedin.com/in/hamza-amraoui/");
        }
        #endregion
        #region SearchForEntity
        private void searchEntity_TextChaneged(object sender, EventArgs e)
        {
            if (dtEntities == null)
                return;

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
                    else
                    {
                        EntityFieldAnalyserManager.SetEntitiesGridViewHeaders(dtEntities.Clone(), EntityGridView);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Invalid Search Character. Please do not use ' [ ] within searches.");
                }
            }
            UpdateGuide();
        }

        private void searchField_TextChaneged(object sender, EventArgs e)
        {
            ApplyFieldFilters();
        }

        #endregion

        private void displayAllColumns_CheckedChanged(object sender, EventArgs e)
        {
            if (fieldPropretiesView.DataSource == null)
                return;
            var columnsCount = _analysedType == DataUsageMode ? fieldPropretiesView.Columns.Count-1 : fieldPropretiesView.Columns.Count;
            for (int i = 13; i < columnsCount; i++)
            {
                fieldPropretiesView.Columns[i].Visible = displayAllColumns.Checked;
            }
        }

        private void buymeacoffeeiconClick(object sender, EventArgs e)
        {
            Process.Start("https://www.paypal.me/EntityFieldsAnalyser");
        }

        private void AnalyseType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(AnalyseType.SelectedItem.ToString() == DataUsageMode) {
            string message = "Metadata + Data usage option could take much time, depends on your entity volum";

            MessageBox.Show(message, "Execution Time",MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            UpdateGuide();
        }
    }
}
