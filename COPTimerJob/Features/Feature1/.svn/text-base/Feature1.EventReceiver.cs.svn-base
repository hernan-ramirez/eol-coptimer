using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Administration;

namespace COPTimerJob.Features.Feature1
{
    /// <summary>
    /// Esta clase controla los eventos generados durante la activación, desactivación, instalación, desinstalación y actualización de características.
    /// </summary>
    /// <remarks>
    /// El GUID asociado a esta clase se puede usar durante el empaquetado y no se debe modificar.
    /// </remarks>

    [Guid("2a659855-90ab-4f27-a38f-322f0bd41787")]
    public class Feature1EventReceiver : SPFeatureReceiver
    {
        const string List_JOB_NAME = "SincronizadorCOP";
        // Uncomment the method below to handle the event raised after a feature has been activated.

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSite site = properties.Feature.Parent as SPSite;

            // make sure the job isn't already registered

            foreach (SPJobDefinition job in site.WebApplication.JobDefinitions)
            {

                if (job.Name == List_JOB_NAME)

                    job.Delete();

            }

            // install the job

            ImportadorTimerJob listLoggerJob = new ImportadorTimerJob(List_JOB_NAME, site.WebApplication);

            SPDailySchedule dailySchedule = new SPDailySchedule();
            dailySchedule.BeginHour = 00;
            dailySchedule.BeginMinute = 50;
            dailySchedule.BeginSecond = 0;
            dailySchedule.EndHour = 00;
            dailySchedule.EndMinute = 51;
            dailySchedule.EndSecond = 59;

            listLoggerJob.Schedule = dailySchedule;

            //SPMinuteSchedule schedule = new SPMinuteSchedule();

            //schedule.BeginSecond = 0;

            //schedule.EndSecond = 59;

            //schedule.Interval = 5;            

            listLoggerJob.Schedule = dailySchedule;

            listLoggerJob.Update();

        }

        //Uncomment the method below to handle the event raised before a feature is deactivated

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSite site = properties.Feature.Parent as SPSite;

            // delete the job

            foreach (SPJobDefinition job in site.WebApplication.JobDefinitions)
            {

                if (job.Name == List_JOB_NAME)

                    job.Delete();

            }
        }
    }
}
