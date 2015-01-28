//--------------------------------------------------------------------------------------------------------------------------------
// <copyright file="CheckedListBoxItem.cs">(c) Controlled Vocabulary on GitHub, 2015. All other rights reserved.</copyright>
//--------------------------------------------------------------------------------------------------------------------------------
namespace ControlledVocabulary
{
    /// <summary>
    /// CheckedListBoxItem helper class
    /// </summary>
    public class CheckedListBoxItem
    {
        /// <summary>
        /// Id of the ListBox item
        /// </summary>
        public int Id { get; set; }
        
        /// <summary>
        /// Name of the ListBox item
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Sets the sourcePath
        /// </summary>
        public string SourcePath { get; set; }

        /// <summary>
        /// Whether the ListBox item is checked
        /// </summary>
        public bool IsChecked { get; set; }
    }
}
