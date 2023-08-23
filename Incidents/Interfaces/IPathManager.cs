using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents.Interfaces
{
    public interface IPathManager
    {
        ConnectionManagerArchiveCollection ArchivesCollection { get; }
        ConnectionManagerTemplateCollection TemplatesCollection { get; }
        ConnectionManagerPathsCollection RegionPathCollection { get; }
        ConnectionManagerEmailCollection EmailsCollection { get; }
        string Email1 { get; }
        string Email2 { get; }
        string Template1Path { get; }
        string Template2Path { get; }
        string ArchivePath { get; }
        string LoggerPath { get; }
        string AggregateDailyArchiveLogFileName { get; }
        string AggregateDailyArchiveFileName { get; }
        string MainLoggerPath { get; }
        string AdditionalLoggerPath { get; }
        string RawArchivePath { get; }
        string Email { get; }
        string EmailSubject { get; }
    }
}
