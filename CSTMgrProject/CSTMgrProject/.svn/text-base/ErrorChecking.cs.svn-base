using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.Collections;
using System.DirectoryServices.ActiveDirectory;
namespace CSTMgrProject
{
    /// <summary>
    /// Class for an error checking message.
    /// Can use this class to display a formal
    /// message for the errors that occurred.
    /// </summary>
    public class ErrorChecking
    {
        #region Attributes
        //Define variable for the error message
        private string ErrorMessage;
        #endregion



        #region Functionality

        /// <summary>
        /// Add a string to our error message
        /// </summary>
        /// <param name="msg">Message to add</param>
        public void AddMsg(string msg)
        {
            ErrorMessage += msg + "\n";
        }
        /// <summary>
        /// Get the error message
        /// </summary>
        /// <returns>Error message</returns>
        public string getMsg()
        {
            return ErrorMessage;
        }
        /// <summary>
        /// Reset the error message
        /// </summary>
        public void resetMsg()
        {
            ErrorMessage = "";
        }

        #endregion
    }
}
