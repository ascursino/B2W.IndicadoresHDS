using System.Runtime.InteropServices;
using Microsoft.SharePoint;
using System.Linq; 

namespace PortalTI.SLM.TJ.IndicadoresHDS.Features.MonitoraJobFeature
{
    [Guid("611d489f-82d6-4453-b61e-a9c3ec6de6ba")]
    public class MonitoraJobFeatureEventReceiver : SPFeatureReceiver
    {
        private const string List_JOB_NAME = "B2W-Indicadores-HDS";

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPSite site = properties.Feature.Parent as SPSite;

                // make sure the job isn't already registered 
                site.WebApplication.JobDefinitions.Where(t => t.Name.Equals(List_JOB_NAME)).ToList().ForEach(j => j.Delete());

                //job a cada 30min 
                MonitoraJob listLoggerJob = new MonitoraJob(List_JOB_NAME, site.WebApplication);
                SPMinuteSchedule schedule = new SPMinuteSchedule();
                schedule.Interval = 30;

                listLoggerJob.Schedule = schedule;
                listLoggerJob.Update();
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                SPSite site = properties.Feature.Parent as SPSite;

                // delete the job 
                site.WebApplication.JobDefinitions.Where(t => t.Name.Equals(List_JOB_NAME)).ToList().ForEach(j => j.Delete());
            });
        }
    }
}
