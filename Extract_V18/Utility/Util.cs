using System;
using System.Diagnostics;
using System.IO;

using Siemens.Engineering;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.SW.Blocks;

namespace TIA_Extract.Utility
{
    internal static class Util
    {
        public static bool ExportBlock(PlcBlock block, string filePath, FeedbackContext feedbackContext = null)
        {
            try
            {
                if (File.Exists(filePath))
                    File.Delete(filePath);
                
                block.Export(new FileInfo(filePath), ExportOptions.None, DocumentInfoOptions.None);
            }
            catch (Exception ex)
            {
                Trace.TraceError("Exception during export:" + Environment.NewLine + ex);

                if(feedbackContext != null) 
                    feedbackContext.Log(NotificationIcon.Error, ex.Message);

                return false;
            }
            return true;
        }
    }
}
