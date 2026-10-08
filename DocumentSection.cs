using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace it.sealink.DocumentProcessing
{
    /// <summary>
    /// Classe per la definizione della sezione 
    /// </summary>
    public class DocumentSection
    {
        /// <summary>
        /// Nome univoco identificativo della sezione
        /// </summary>
        public string SectionName { get; set; } = string.Empty;
        /// <summary>
        /// Classe DocumentPlaceHolder per la definizione dei placeholder da sostituire
        /// </summary>
        public DocumentPlaceHolder PlaceHolder { get; set; } = new DocumentPlaceHolder();
        /// <summary>
        /// Definizione della visibilità della sezione 
        /// </summary>
        public bool IsVisible { get; set; } = true;
    }
}
