using System;
using Microsoft.SharePoint.Administration;

namespace PortalTI.SLM.TJ.IndicadoresHDS
{
    public class MonitoraJobConfig: SPPersistedObject
    {
        public static string ConfigNome = "MonitoraJobConfig";
        public MonitoraJobConfig() { }
        public MonitoraJobConfig(SPPersistedObject parent, Guid id) : base(ConfigNome, parent, id) { }
    }
}