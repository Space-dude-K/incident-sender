using System.Configuration;

namespace Incidents
{
    public class ExcelDataSection : ConfigurationSection
    {
        public const string SectionName = "ExcelDataSection";

        private const string RegionCollectionName = "RegionPaths";

        private const string TemplateCollectionName = "Template";

        private const string ArchiveCollectionName = "Archive";

        private const string EmailCollectionName = "Email";

        [ConfigurationProperty(RegionCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerPathsCollection), AddItemName = "add")]
        public ConnectionManagerPathsCollection RegionPaths { get { return (ConnectionManagerPathsCollection)base[RegionCollectionName]; } }

        [ConfigurationProperty(TemplateCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerTemplateCollection), AddItemName = "add")]
        public ConnectionManagerTemplateCollection Templates { get { return (ConnectionManagerTemplateCollection)base[TemplateCollectionName]; } }

        [ConfigurationProperty(ArchiveCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerArchiveCollection), AddItemName = "add")]
        public ConnectionManagerArchiveCollection Archives { get { return (ConnectionManagerArchiveCollection)base[ArchiveCollectionName]; } }

        [ConfigurationProperty(EmailCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerEmailCollection), AddItemName = "add")]
        public ConnectionManagerEmailCollection Emails { get { return (ConnectionManagerEmailCollection)base[EmailCollectionName]; } }
    }
    public class ConnectionManagerPathsCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ConnectionManagerPathElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ConnectionManagerPathElement)element).Name;
        }
    }
    public class ConnectionManagerTemplateCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ConnectionManagerTemplateElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ConnectionManagerTemplateElement)element).Name;
        }
    }
    public class ConnectionManagerArchiveCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ConnectionManagerArchiveElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ConnectionManagerArchiveElement)element).Name;
        }
    }
    public class ConnectionManagerPathElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get { return (string)this["name"]; }
            set { this["name"] = value; }
        }

        [ConfigurationProperty("path", IsRequired = true)]
        public string Path
        {
            get { return (string)this["path"]; }
            set { this["path"] = value; }
        }
    }
    public class ConnectionManagerTemplateElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get { return (string)this["name"]; }
            set { this["name"] = value; }
        }

        [ConfigurationProperty("path", IsRequired = true)]
        public string Path
        {
            get { return (string)this["path"]; }
            set { this["path"] = value; }
        }
    }
    public class ConnectionManagerArchiveElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get { return (string)this["name"]; }
            set { this["name"] = value; }
        }

        [ConfigurationProperty("path", IsRequired = true)]
        public string Path
        {
            get { return (string)this["path"]; }
            set { this["path"] = value; }
        }
    }
}
