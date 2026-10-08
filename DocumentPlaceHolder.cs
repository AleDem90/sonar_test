using System.Data;

namespace it.sealink.DocumentProcessing
{
    /// <summary>
    /// classe per la definizione dei placeholder
    /// nel documento
    /// </summary>
    public class DocumentPlaceHolder
    {
        
        private Dictionary<string, object> _SinglePlaceHolders = new Dictionary<string, object>();
        /// <summary>
        /// Dizionario per la sostiztuzione dei placeholder
        /// </summary>
        public Dictionary<string, object> SinglePlaceHolders
        {
            get { return _SinglePlaceHolders; }
            set { _SinglePlaceHolders = value; }
        }

        private List<DataTable> _TablePlaceHolders = new List<DataTable>();
        /// <summary>
        /// Lista di datatable per la sostituzione di tabelle all'interno del documento
        /// </summary>
        public List<DataTable> TablePlaceHolders
        {
            get { return _TablePlaceHolders; }
            set { _TablePlaceHolders = value; }
        }

        internal Dictionary<string, ImageDimension> _imagePlaceHoldersDimension = new Dictionary<string, ImageDimension>();

        public void SetImagePlaceHolderDimension(string key, int width, int height)
        {
            if (!_imagePlaceHoldersDimension.ContainsKey(key))
            {
                _imagePlaceHoldersDimension.Add(key, new ImageDimension()
                {
                    Width = width,
                    Height = height
                });
            }
        }
    }

    internal class ImageDimension
    {
        public int Width { get; set; }
        public int Height { get; set; }
    }
}
