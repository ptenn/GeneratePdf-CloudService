using System;
using System.IO;
using GeneratePdf_WCFService.Util;
using log4net;

namespace GeneratePdf_WCFService.Jobs
{
    public class FileCleanupJob
    {
        /// <summary>
        /// Log4Net logger set up.
        /// </summary>
        private static readonly ILog logger = LogManager.GetLogger(typeof(FileCleanupJob));

        /// <summary>
        /// Implementation of the Job Execute method.
        /// </summary>
        /// <param name="context">Reference to the JobExecutionContext.</param>
        public void Cleanup(string folder, int hours)
        {
            logger.Info("Executing FileCleanupJob");
            try
            {
                    
                logger.Info($"Checking Folder {folder},  for any files older than {hours} hours.");


                if (Directory.Exists(folder))
                {
                    // Handle files
                    String[] fileNames = Directory.GetFiles(folder);
                    foreach (string fileName in fileNames)
                    {
                        if (DateTime.Compare(DateTime.Now,
                                File.GetCreationTime(fileName).AddHours(hours)) > 0)
                        {
                            logger.Debug($"Found file '{fileName}' older than {hours} hours, checking file permissions.");
                            UserFileAccessRights userFileAccessRights = new UserFileAccessRights(folder);

                            if (userFileAccessRights.canDelete())
                            {
                                logger.Debug($"Verified permissions to delete file '{fileName}', attempting delete.");
                                try
                                {
                                    File.Delete(fileName);
                                    logger.Debug($"Successfully deleted file '{fileName}'.");
                                }
                                catch (Exception fileEx)
                                {
                                    logger.Debug($"Exception occurred attempting to delete file '{fileName}'.", fileEx);
                                }

                            }
                            else
                            {
                                logger.Debug($"Do not have permissions to delete file {fileName} in {folder}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("Attempt to run FileCleanupJob failed with exception.", ex);
            }
        }
    }
}