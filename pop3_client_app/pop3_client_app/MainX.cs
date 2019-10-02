using System;
using System.Windows.Forms;

namespace Pop3ClinetApp
{
	/// <summary>
	/// Application main calss.
	/// </summary>
	public class MainX
    {
        #region static method Main

        /// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		public static void Main() 
		{
            try{
                Application.EnableVisualStyles();
			    Application.Run(new wfrm_Main());
            }
            catch{
            }
        }

        #endregion
    }
}
