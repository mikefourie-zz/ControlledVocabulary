//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="Enumerations.cs">(c) Controlled Vocabulary on GitHub, 2015. All other rights reserved.</copyright>
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
        /// Outlook2010 Client
        /// </summary>
        Outlook2010,

        /// <summary>
        /// WindowsDesktop Client
        /// </summary>
        WindowsDesktop
    }
}
