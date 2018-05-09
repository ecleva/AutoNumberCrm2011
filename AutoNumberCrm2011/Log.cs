using System;
using System.Diagnostics;

//***************************************************
//© 2008  All rights reserved.
//***************************************************
// Date		    Who			    Description
// 15-Aug-2008	Denny A1oor		AutoNumber Creation
//***************************************************

namespace AutoNumberCrm2011
{
	/// <summary>
	/// Summary description for logging.
	/// </summary>
	public class log
	{
		/// <summary>
		/// Writes the input message to the server event log
		/// </summary>
		/// <param name="strMessage">the message</param>
		public static void logError( string strMessage )
		{
			try
			{
				EventLog objLog;
				objLog = new EventLog();
				objLog.Source = "MSCRM";
				objLog.WriteEntry( strMessage, EventLogEntryType.Error );
			}
			catch {}
		}


        /// <summary>
        /// Writes the input message to the server event log
        /// </summary>
        /// <param name="strMessage">the message</param>
        public static void logInfo(string strMessage)
        {
            try
            {
                EventLog objLog;
                objLog = new EventLog();
                objLog.Source = "MSCRM";
                objLog.WriteEntry(strMessage, EventLogEntryType.Information);
            }
            catch { }
        }
	}
}
