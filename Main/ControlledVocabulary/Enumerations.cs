//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Enumerations.cs">(c) Controlled Vocabulary on CodePlex, 2010. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    public enum MessageType
    {
        /// <summary>
        /// Info Message
        /// </summary>
        Info = 0,

        /// <summary>
        /// Warning Message
        /// </summary>
        Warning = 1,

        /// <summary>
        /// Error Message
        /// </summary>
        Error = 2
    }

    public enum ClientType
    {
        /// <summary>
        /// Outlook2003 Client
        /// </summary>
        Outlook2003 = 0,

        /// <summary>
        /// Outlook2007 Client
        /// </summary>
        Outlook2007 = 1,

        /// <summary>
        /// Outlook2010 Client
        /// </summary>
        Outlook2010 = 2,

        /// <summary>
        /// WindowsDesktop Client
        /// </summary>
        WindowsDesktop = 3
    }
}
