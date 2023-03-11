using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents
{
    class DateDataSection : ConfigurationSection
    {
        public const string SectionName = "DateDataSection";

        private const string NotWorkingDaysCollectionName = "NotWorkingDays";

        private const string WorkingDaysCollectionName = "WorkingDays";

        private const string EmailCollectionName = "Email";

        [ConfigurationProperty(NotWorkingDaysCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerNotWorkingDaysCollection), AddItemName = "add")]
        public ConnectionManagerNotWorkingDaysCollection NotWorkingDays { get { return (ConnectionManagerNotWorkingDaysCollection)base[NotWorkingDaysCollectionName]; } }

        [ConfigurationProperty(WorkingDaysCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerWorkingDaysCollection), AddItemName = "add")]
        public ConnectionManagerWorkingDaysCollection WorkingDays { get { return (ConnectionManagerWorkingDaysCollection)base[WorkingDaysCollectionName]; } }

        [ConfigurationProperty(EmailCollectionName)]
        [ConfigurationCollection(typeof(ConnectionManagerEmailCollection), AddItemName = "add")]
        public ConnectionManagerEmailCollection Emails { get { return (ConnectionManagerEmailCollection)base[EmailCollectionName]; } }
    }
    public class ConnectionManagerNotWorkingDaysCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ConnectionManagerNotWorkingDaysElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ConnectionManagerNotWorkingDaysElement)element).Name;
        }
    }
    public class ConnectionManagerWorkingDaysCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ConnectionManagerWorkingDaysElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ConnectionManagerWorkingDaysElement)element).Name;
        }
    }
    public class ConnectionManagerEmailCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ConnectionManagerEmailElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ConnectionManagerEmailElement)element).Name;
        }
    }
    public class ConnectionManagerNotWorkingDaysElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get { return (string)this["name"]; }
            set { this["name"] = value; }
        }

        [ConfigurationProperty("date", IsRequired = true)]
        public string Date
        {
            get { return (string)this["date"]; }
            set { this["date"] = value; }
        }
    }
    public class ConnectionManagerWorkingDaysElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get { return (string)this["name"]; }
            set { this["name"] = value; }
        }

        [ConfigurationProperty("date", IsRequired = true)]
        public string Date
        {
            get { return (string)this["date"]; }
            set { this["date"] = value; }
        }
    }
    public class ConnectionManagerEmailElement : ConfigurationElement
    {
        [ConfigurationProperty("name", IsRequired = true)]
        public string Name
        {
            get { return (string)this["name"]; }
            set { this["name"] = value; }
        }

        [ConfigurationProperty("email", IsRequired = true)]
        public string Email
        {
            get { return (string)this["email"]; }
            set { this["email"] = value; }
        }
    }
}
