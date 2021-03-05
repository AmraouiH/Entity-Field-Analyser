using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EntityieldsAnalyser
{
    public class DataColumns
    {
        public string fieldName;
        public string fieldType;
        public DateTime dateOfCreation;
        public string introducedVersion;
        public bool isManaged;
        public bool isAuditable;
        public string requiredLevel;
        public bool isSearchable;
        public bool isOnForm;
        public string target;
        public string percentageOfUse;
    }

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

    public class EntityInfo
    {
        public string entityName;
        public string entityTechnicalName;
        public int numberOfFields;
        public int numberOfRecords;
    }
}
