using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntityieldsAnalyser
{
    public class DataColumns
    {
        public string fieldType;
        public string displayName;
        public string fieldName;
        public string Description;
        public DateTime dateOfCreation;
        public string introducedVersion;
        public string isManaged;
        public bool? isAuditable;
        public string requiredLevel;
        public bool? isSearchable;
        public bool? isOnForm;
        public bool? isCustom;
        public string target;
        public int totalFiledRecords;
        public double percentageOfUse;
        public string AttributeOf;
        public string AutoNumberFormat;
        public bool? CanBeSecuredForCreate;
        public bool? CanBeSecuredForRead;
        public bool? CanBeSecuredForUpdate;
        public bool? CanModifyAdditionalSettings;
        public int ColumnNumber;
        public string DeprecatedVersion;
        public string ExternalName;
        public string InheritsFrom;
        public bool? IsCustomizable;
        public bool? IsDataSourceSecret;
        public bool? IsFilterable;
        public bool? IsGlobalFilterEnabled;
        public bool? IsLogical;
        public bool? IsPrimaryId;
        public bool? IsPrimaryName;
        public bool? IsRenameable;
        public bool? IsRequiredForForm;
        public bool? IsRetrievable;
        public bool? IsSecured;
        public bool? IsSortableEnabled;
        public bool? IsValidForAdvancedFind;
        public bool? IsValidForCreate;
        public bool? IsValidForForm;
        public bool? IsValidForGrid;
        public bool? IsValidForRead;
        public bool? IsValidForUpdate;
        public bool? IsValidODataAttribute;
        public string LinkedAttributeId;
        public string EntityLogicalName;
        public DateTime ModifiedOn;
        public string SourceType;
    }

    public class entityParam
    {
        public string displayName;
        public string fieldName;
        public string attributeType;
        public string Description;
        public DateTime dateOfCreation;
        public string introducedVersion;
        public string isManaged;
        public bool? isAuditable;
        public string requiredLevel;
        public bool? isSearchable;
        public bool? isOnForm;
        public bool? isCustom;
        public string target;
        public int totalFiledRecords;
        public double percentageOfUse;
        public string AttributeOf;
        public string AutoNumberFormat;
        public bool? CanBeSecuredForCreate;
        public bool? CanBeSecuredForRead;
        public bool? CanBeSecuredForUpdate;
        public bool? CanModifyAdditionalSettings;
        public int ColumnNumber;
        public string DeprecatedVersion;
        public string ExternalName;
        public string InheritsFrom;
        public bool? IsCustomizable;
        public bool? IsDataSourceSecret;
        public bool? IsFilterable;
        public bool? IsGlobalFilterEnabled;
        public bool? IsLogical;
        public bool? IsPrimaryId;
        public bool? IsPrimaryName;
        public bool? IsRenameable;
        public bool? IsRequiredForForm;
        public bool? IsRetrievable;
        public bool? IsSecured;
        public bool? IsSortableEnabled;
        public bool? IsValidForAdvancedFind;
        public bool? IsValidForCreate;
        public bool? IsValidForForm;
        public bool? IsValidForGrid;
        public bool? IsValidForRead;
        public bool? IsValidForUpdate;
        public bool? IsValidODataAttribute;
        public string LinkedAttributeId;
        public string EntityLogicalName;
        public DateTime ModifiedOn;
        public string SourceType;
    }

    public class EntityInfo
    {
        public string entityName;
        public string entityTechnicalName;
        public DateTime entityDateOfCreation;
        public int entityFieldsCount = 0;
        public int entityRecordsCount = 0;
        public int managedFieldsCount = 0;
        public int unmanagedFieldsCount = 0;
        public int entityCustomFieldsCount = 0;
        public int entityStandardFieldsCount = 0;
        public int entityTotalUseOfColumns = 0;
        public int entityDefaultColumnSize = 1024;
    }
}
