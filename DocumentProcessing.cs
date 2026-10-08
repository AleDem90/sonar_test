using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using it.sealink.DocumentProcessing.utils;
using log4net;
using log4net.Repository.Hierarchy;
using SkiaSharp;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DW = DocumentFormat.OpenXml.Drawing;

namespace it.sealink.DocumentProcessing
{
    public class DocumentProcessing
    {
        private readonly ILog logger = LogExtensions.GetClassLogger();
        private Dictionary<string, ImageDimension> _imagedimension = new Dictionary<string, ImageDimension>();
        /// <summary>
        /// Crea il documento word
        /// </summary>
        /// <param name="savePath">
        /// Percorso di salvataggio
        /// </param>
        /// <param name="templatePath">
        /// Percorso del documento template
        /// </param>
        /// <param name="sections">
        /// lista di oggetti DocumentSection per la sostituzione dei placeholder
        /// </param>
        /// <returns>ritorna True se la creazione è avvenuta con successo</returns>
        public bool CreateWordDocument(string savePath, string templatePath, List<DocumentSection> sections)
        {
            logger.MethodEnter();
            bool res = true;

            try
            {
                //se non è definita la lista di sezioni ritorna
                if (sections?.Count <= 0)
                {
                    return false;
                }
                //crea la copia del template
                File.Copy(templatePath, savePath, true);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(savePath, true))
                {


                    Dictionary<DocumentSection, List<OpenXmlElement>> sectionMap = CheckSections(wordDoc, sections);
                    if (sectionMap is not null)
                    {
                        foreach (KeyValuePair<DocumentSection, List<OpenXmlElement>> item in sectionMap)
                        {
                            ReplacePlaceHolders(wordDoc, item.Key, item.Value);
                        }
                        
                        wordDoc.Save();
                    }
                    else
                    {
                        File.Delete(savePath);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error("CreateWordDocument: " + ex.Message, ex);
                File.Delete(savePath);
                return false;
            }

            return res;
        }
        /// <summary>
        /// Genera un file txt in memoria (MemoryStream) inserendo una stringa di intestazione 
        /// in prima riga e successivamente le righe del DataTable.
        /// </summary>
        /// <param name="listTxt">La lista di oggetti da esportare.</param>
        /// <param name="intestazione">La stringa ricevuta contenente l'intestazione.</param>
        /// <returns>L'array di byte del file .txt generato.</returns>
        public static byte[] CreateTxt<T>(IEnumerable<T> listTxt, string header)
        {
            using var stream = new MemoryStream();
            // UTF8Encoding(false) genera il file UTF-8 senza BOM (Byte Order Mark)
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                // Scrittura dell'intestazione (se presente)
                if (!string.IsNullOrWhiteSpace(header))
                {
                    writer.WriteLine(header);
                }

                // Scrittura delle righe
                if (listTxt.Count().Equals(0))
                    throw new Exception("Nessuna riga trovata per il file");
                foreach (T row in listTxt)
                {
                    writer.WriteLine(row);
                }

                writer.Flush();
            }

            return stream.ToArray();
        }
        public static byte[] CreateZip<T>(IEnumerable<T> listTxt, string header, int maxRows)
        {
            if (listTxt == null || !listTxt.Any())
                throw new Exception("Nessuna riga trovata per il file");

            if (maxRows <= 0)
                throw new ArgumentException("maxRows deve essere maggiore di 0", nameof(maxRows));

            using var memoryStream = new MemoryStream();
            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
            {
                int fileIndex = 1;

                foreach (var chunk in listTxt.Chunk(maxRows))
                {
                    
                    string fileName = $"QRY_{DateTime.Now:yyyyMMdd}_{fileIndex}_NUM_TELEFONO.txt";
                    ZipArchiveEntry zipEntry = archive.CreateEntry(fileName, CompressionLevel.Optimal);

                    using (var entryStream = zipEntry.Open())
                    using (var writer = new StreamWriter(entryStream, new UTF8Encoding(false)))
                    {
                        if (!string.IsNullOrWhiteSpace(header))
                        {
                            writer.WriteLine(header);
                        }

                        foreach (T row in chunk)
                        {
                            writer.WriteLine(row);
                        }

                        writer.Flush();
                    }

                    fileIndex++;
                }
            }

            return memoryStream.ToArray();
        }
        /// <summary>
        /// Genera un file Excel in memoria (MemoryStream) inserendo una stringa di intestazione 
        /// in prima riga e successivamente le righe del DataTable.
        /// </summary>
        /// <param name="dictSheet">Il Dictionary con chiave nome sheet e object i dati da esportare.</param>
        /// <param name="intestazione">La stringa ricevuta contenente l'intestazione.</param>
        /// <returns>L'array di byte del file .xlsx generato.</returns>
        public static byte[] CreateExcel<T>(Dictionary<string, IEnumerable<T>> dictSheet, string intestazione)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                foreach (string key in dictSheet.Keys)
                { 
                    if (String.IsNullOrEmpty(key) || dictSheet[key] is null)
                        throw new Exception("Nessun dataTable per la chiave: " + key);
                    var worksheet = workbook.Worksheets.Add(key);
                    if (!string.IsNullOrEmpty(intestazione))
                        worksheet.Cell(1, 1).Value = intestazione;

                    worksheet.Cell((string.IsNullOrEmpty(intestazione) ? 1 : 2), 1).InsertData(dictSheet[key]);
                    
                }
                using var stream = new MemoryStream();
                workbook.SaveAs(stream);
                return stream.ToArray();
            }
        }

        #region private methods
        private void ReplicateSections(List<DocumentSection> sections, List<OpenXmlElement> elements)
        {
            logger.MethodEnter();
            try
            {
                if (sections is null || sections.Count.Equals(0) || elements is null || elements.Count.Equals(0))
                    return;

                // Regex per estrarre il nome e il numero (es. pippo_*1*)
                Regex regex = new(@"(\w+)_\*(\d+)\*");

                // Dizionario per raggruppare i numeri per nome sezione
                var replicaSet = new Dictionary<string, List<int>>();

                foreach (var section in sections)
                {
                    var match = regex.Match(section.SectionName);
                    if (match.Success)
                    {
                        string name = match.Groups[1].Value;        // Nome (es. pippo, pluto)
                        int number = int.Parse(match.Groups[2].Value); // Numero (es. 1, 2, 3)

                        if (!replicaSet.TryGetValue(name, out var numbers))
                        {
                            numbers = [];
                            replicaSet[name] = numbers;
                        }
                        numbers.Add(number);
                    }
                }

                // Elaborazione delle sezioni da replicare
                foreach (var (sectionName, numbers) in replicaSet)
                {
                    var elementsToReplicate = new List<OpenXmlElement>();
                    bool isCollecting = false;

                    string startMarker = $"[#Section_{sectionName}_*#*]";
                    string endMarker = $"[#EndSection_{sectionName}_*#*]";

                    foreach (var element in elements)
                    {
                        string elementText = element is Text textElement
                            ? textElement.Text.Trim()
                            : element.InnerText.Trim();

                        if (!isCollecting)
                        {
                            if (!string.IsNullOrEmpty(elementText) && elementText.Equals(startMarker, StringComparison.Ordinal))
                            {
                                isCollecting = true;
                                elementsToReplicate.Add(element);
                            }
                        }
                        else
                        {
                            elementsToReplicate.Add(element);
                            if (!string.IsNullOrEmpty(elementText) && elementText.Equals(endMarker, StringComparison.Ordinal))
                            {
                                break; // Usciamo dal ciclo degli elementi una volta trovato il tag di chiusura
                            }
                        }
                    }

                    // TODO: Qui puoi implementare la logica per replicare gli elementi in base ai 'numbers' trovati
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }


        /// <summary>
        /// Controlla che almeno un elemento definito nella lista sia presente nel documento,
        /// gli elementi presenti nelle sezioni non trovate verranno eliminati.
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="sections"></param>
        /// <returns></returns>
        private Dictionary<DocumentSection, List<OpenXmlElement>> CheckSections(WordprocessingDocument wordDoc, List<DocumentSection> sections)
        {
            logger.MethodEnter();
            try
            {
                Dictionary<DocumentSection, List<OpenXmlElement>> sectionMap = new Dictionary<DocumentSection, List<OpenXmlElement>>();
                if (sections?.Count <= 0)
                {
                    return [];
                }

                bool sectionFound = false;
                //header
                List<OpenXmlElement> documentElementList = new List<OpenXmlElement>();
                if (wordDoc.MainDocumentPart?.HeaderParts is not null)
                    foreach (HeaderPart h in wordDoc.MainDocumentPart.HeaderParts)
                    {
                        if (h.RootElement?.Descendants<OpenXmlElement>().ToList() is not null)
                        {
                            documentElementList.AddRange(h.RootElement.Descendants<OpenXmlElement>().ToList());
                        }

                    }

                //body
                if (wordDoc.MainDocumentPart?.Document?.Body?.Descendants<OpenXmlElement>().ToList() is not null)
                    documentElementList.AddRange(wordDoc.MainDocumentPart.Document.Body.Descendants<OpenXmlElement>().ToList());

                //footer
                if (wordDoc.MainDocumentPart?.FooterParts is not null)
                    foreach (FooterPart f in wordDoc.MainDocumentPart.FooterParts)
                    {
                        if (f.RootElement?.Descendants<OpenXmlElement>().ToList() is not null)
                        {
                            documentElementList.AddRange(f.RootElement.Descendants<OpenXmlElement>().ToList());
                        }

                    }
                //trovo le sezioni all'interno del documento che non sono presenti nella lista delle sezioni definite per la sostituzione
                //e creo una DocumentSection con IsVisible false.
                foreach (OpenXmlElement element in documentElementList)
                {
                    string elementText = string.Empty;
                    if(element is Text textelement)
                    {
                        elementText = textelement.Text.Trim();
                    }
                    else
                    {
                        elementText = element.InnerText.Trim();
                    }
                    if (string.IsNullOrEmpty(elementText))
                        continue;
                    if (!elementText.StartsWith("[#Section_"))
                        continue;
                    if (!elementText.EndsWith("]"))
                        continue;
                    var sect = sections?.FirstOrDefault(x => elementText.Equals($"[#Section_{x.SectionName}]"));
                    if (sect is null)
                    {
                        DocumentSection docsect = new DocumentSection()
                        {
                            SectionName = elementText.Replace("[#Section_", "").Replace("]", ""),
                            IsVisible = false
                        };

                        if (sections?.Where(x => x.SectionName.Equals(docsect.SectionName)).ToList().Count <= 0)
                        {
                            sections.Add(docsect);
                        }
                    }
                }

                foreach (DocumentSection section in sections)
                {
                    OpenXmlElement? start_section_element = documentElementList?.FirstOrDefault(x => !string.IsNullOrEmpty(x.InnerText) && x.InnerText.Trim().Equals($"[#Section_{section.SectionName}]"));


                    if (start_section_element is null)
                        continue;

                    int indxstart = documentElementList?.IndexOf(start_section_element) ?? -1;

                    OpenXmlElement? end_section_element = documentElementList?.FirstOrDefault(x => !string.IsNullOrEmpty(x.InnerText) && x.InnerText.Trim().Equals($"[#EndSection_{section.SectionName}]"));
                    if (end_section_element is null)
                        continue;
                    int indxend = documentElementList?.IndexOf(end_section_element) ?? -1;
                    if (indxstart.Equals(-1) || indxend.Equals(-1))
                    {
                        continue;
                    }
                    
                    if (!section.IsVisible)
                    {

                        for (int i = indxstart; i <= indxend; i++)
                        {
                            if(documentElementList[i].Parent is not null)
                            {
                                documentElementList[i].Remove();
                            }
                            
                        }
                    }
                    else
                    {
                        List<OpenXmlElement> sectionelementlist = new List<OpenXmlElement>();
                       for (int i = indxstart; i <= indxend; i++)
                        {
                            if (i.Equals(indxstart) || i.Equals(indxend))
                            {
                                documentElementList[i].RemoveAllChildren();
                                if (documentElementList[i].Parent is not null)
                                {
                                    documentElementList[i].Remove();
                                }
                            }
                            else
                            {
                                sectionelementlist.Add(documentElementList[i]);
                            }
                        }
                        if (sectionMap.TryGetValue(section, out List<OpenXmlElement> elements))
                        {
                            logger.Error("section already defined!");
                        }
                        else
                        {
                            sectionMap.Add(section, sectionelementlist);
                        }


                    }
                }
                return sectionMap;
            }
            catch(Exception ex)
            {
                logger.Error(ex.Message, ex);
                return [];
            }
        }
        
        private void ReplacePlaceHolders(WordprocessingDocument wordDoc, DocumentSection section, List<OpenXmlElement> documentSectionElements)
        {
            logger.MethodEnter();
            try
            {
                if(section?.PlaceHolder?._imagePlaceHoldersDimension is not null)
                {
                    _imagedimension = section?.PlaceHolder?._imagePlaceHoldersDimension;
                }
                if (section.PlaceHolder?.SinglePlaceHolders is not null && section.PlaceHolder?.SinglePlaceHolders.Count > 0)
                {
                    ReplaceSinglePlaceHolders(wordDoc, documentSectionElements, section.PlaceHolder?.SinglePlaceHolders);
                }
                if (section.PlaceHolder?.TablePlaceHolders is not null && section.PlaceHolder?.TablePlaceHolders.Count > 0)
                {
                    ReplaceTablePlaceHolders(wordDoc, documentSectionElements, section.PlaceHolder?.TablePlaceHolders);
                }

            }
            catch(Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }
        /// <summary>
        /// Sostituisce gli elementi openxmlelement con i valori presenti nel dizionario
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="documentSectionElements"></param>
        /// <param name="singlePlaceHolders"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void ReplaceSinglePlaceHolders(WordprocessingDocument wordDoc, List<OpenXmlElement> documentSectionElements, Dictionary<string, object>? singlePlaceHolders)
        {
            logger.MethodEnter();
            try
            {
                HashSet<Text> documenTextList = new HashSet<Text>();
                foreach (OpenXmlElement element in documentSectionElements)
                {
                    if (element is Text textelement)
                    {
                        documenTextList.Add(textelement);
                    }
                    else
                    {
                        List<Text> textelements = element.Descendants<Text>().ToList();
                        if (textelements is not null && textelements.Count > 0)
                        {
                            foreach (Text item in textelements)
                            {
                                documenTextList.Add(item);
                            }
                            
                        }
                    }

                    
                }
                for (int placeHolderIndex = 0; placeHolderIndex < singlePlaceHolders.Count; placeHolderIndex++)
                {
                    // trovo gli elementi text che contengono la chiave
                    HashSet<Text> textelementsToReplace = new HashSet<Text>();
                    bool insertIntoList = false;
                    foreach(Text textelement in documenTextList)
                    {
                        string text = textelement.Text;
                        if (text.Contains("%"))
                        {
                            insertIntoList = true;
                        }
                        
                        if (insertIntoList)
                        {
                            textelementsToReplace.Add(textelement);
                        }
                        if (text.Contains("$") && insertIntoList)
                        {
                            textelementsToReplace.Add(textelement);
                            break;
                        }
                    }
                    string chiave = "%" + singlePlaceHolders.ElementAt(placeHolderIndex).Key + "$";
                    string fullText = string.Join("", textelementsToReplace.Select(t => t.Text)); // Unisce tutto il testo

                    if(fullText.Contains(chiave))
                    {
                        if (singlePlaceHolders.ElementAt(placeHolderIndex).Key.Contains("_IMAGE"))
                        {
                            ImageDimension dimension = null;
                            if (_imagedimension.ContainsKey(singlePlaceHolders.ElementAt(placeHolderIndex).Key))
                            {
                                dimension = _imagedimension[singlePlaceHolders.ElementAt(placeHolderIndex).Key];
                            }
                            if (singlePlaceHolders.ElementAt(placeHolderIndex).Value is string s && string.IsNullOrEmpty(s))
                            {
                                ReplaceTextPlaceHolder(textelementsToReplace.ToList(), singlePlaceHolders.ElementAt(placeHolderIndex).Value, chiave);
                                continue;
                            }
                            ReplaceImagePlaceHolder(wordDoc, textelementsToReplace.ToList(), singlePlaceHolders.ElementAt(placeHolderIndex).Value, dimension);
                        }
                        else if (singlePlaceHolders.ElementAt(placeHolderIndex).Key.Contains("_HTML"))
                        {
                            ReplaceHtmlPlaceHolder(wordDoc, textelementsToReplace.ToList(), singlePlaceHolders.ElementAt(placeHolderIndex).Value);
                        }
                        else
                        {
                            ReplaceTextPlaceHolder(textelementsToReplace.ToList(), singlePlaceHolders.ElementAt(placeHolderIndex).Value, chiave);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }

        private void ReplaceTablePlaceHolders(WordprocessingDocument wordDoc, List<OpenXmlElement> documentSectionElements, List<System.Data.DataTable>? tablePlaceHolders)
        {
            logger.MethodEnter();
            try
            {
                HashSet<Table> documenTableList = new HashSet<Table>();
                foreach (OpenXmlElement element in documentSectionElements)
                {
                    if (element is Table table)
                    {
                        documenTableList.Add(table);
                    }

                    List<Table> tables = element.Descendants<Table>().ToList();
                    if (tables is not null && tables.Count > 0)
                    {
                        foreach (Table t in tables)
                            documenTableList.Add(t);
                    }


                }

                foreach (Table table in documenTableList)
                {
                    //verifico che le colonne combaciano con le colonne definite nel tableplaceholder
                    List<TableRow> tableRowList = table.Elements<TableRow>().ToList();
                    foreach (TableRow row in tableRowList)
                    {
                        //lista dei placeholder trovati nella riga
                        List<string> placeholderList = new List<string>();
                        //lista degli indici colonna dei placeholder immagine
                        List<int> imageindexcell = new List<int>();
                        //lista degli indici colonna dei placeholder html
                        List<int> htmlindexcell = new List<int>();
                        //pattern da ricercare nel testo della cella
                        //La regex @"%\w+\$" cerca una stringa che inizia con %, seguita da una o più lettere/numeri/sottolineature (\w+), e terminante con $.
                        string pattern = @"%\w+\$";
                        bool checkPlaceHolder = false;
                        List<TableCell> tableCellList = row.Elements<TableCell>().ToList();
                        int cell_index = 0;
                        foreach (TableCell cell in tableCellList)
                        {
                            List<string> cellcontentlist = cell.Where(x => Regex.IsMatch(x.InnerText, pattern))?.Select(x => x.InnerText).ToList();
                            string cellcontent = string.Empty;
                            if(cellcontentlist is not null)
                            {
                                cellcontent = string.Join("", cellcontentlist);
                            }
                            if (cellcontent.Contains("%") && cellcontent.Contains("$") &&
                                cellcontent.LastIndexOf("%").Equals(0) &&
                                cellcontent.IndexOf("$").Equals(cellcontent.Length - 1))
                            {
                                if (cellcontent.Contains("_IMAGE"))
                                {
                                    imageindexcell.Add(cell_index);
                                }
                                if (cellcontent.Contains("_HTML"))
                                {
                                    htmlindexcell.Add(cell_index);
                                }
                                Regex reg = new Regex("[%$]");
                                placeholderList.Add(reg.Replace(cellcontent, ""));
                                checkPlaceHolder = true;
                            }
                            cell_index++;
                        }

                        if (checkPlaceHolder && tablePlaceHolders is not null)
                        {
                            foreach (DataTable data_table in tablePlaceHolders)
                            {
                                List<string> columnNames = data_table.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToList();
                                //verifico che le liste di placeholder e columnName siano uguali
                                if (columnNames.All(placeholderList.Contains) && columnNames.Count.Equals(placeholderList.Count))
                                {
                                    //sostituisco i placeholder per la prima riga
                                    foreach (string colname in placeholderList)
                                    {
                                        object value = string.Empty;
                                        if (data_table.Rows.Count > 0)
                                            value = data_table.Rows[0][colname];
                                        string chiave = "%" + colname + "$";

                                        TableCell cell = tableCellList.FirstOrDefault(x => string.Join("",x.Descendants<Text>().Select(y => y.Text)).Contains(chiave));
                                        TableCellProperties? cell_props = cell?.TableCellProperties;
                                        
                                        List<Text> placeholdList = cell?.Descendants<Text>().ToList();
                                        if (colname.Contains("_IMAGE"))
                                        {
                                            ImageDimension dimension = null;
                                            if (_imagedimension.ContainsKey(colname))
                                            {
                                                dimension = _imagedimension[colname];
                                            }
                                            if (value is string s && string.IsNullOrEmpty(s))
                                            {
                                                ReplaceTextPlaceHolder(placeholdList, value, chiave);
                                                continue;
                                            }
                                            ReplaceImagePlaceHolder(wordDoc, placeholdList, value, dimension);
                                        }
                                        else if (colname.Contains("_HTML"))
                                        {
                                            ReplaceHtmlPlaceHolder(wordDoc, placeholdList, value);
                                        }
                                        else
                                        {
                                            ReplaceTextPlaceHolder( placeholdList, value, chiave);
                                        }
                                    }

                                    //inserisco le righe successive
                                    for (int i = 1; i < data_table.Rows.Count; i++)
                                    {
                                        TableRow newtablerow = new TableRow();
                                        //copio le proprietà della riga
                                        TableRowProperties row_properties = row.GetFirstChild<TableRowProperties>();
                                        if(row_properties is not null)
                                        {
                                            newtablerow.AppendChild<TableRowProperties>((TableRowProperties)row_properties.CloneNode(true));
                                        }
                                        for (int j = 0; j < placeholderList.Count; j++)
                                        {
                                            TableCell existing_cell = tableCellList[j].CloneNode(true) as TableCell;
                                            TableCellProperties cell_properties = existing_cell.GetFirstChild<TableCellProperties>();
                                            Paragraph existing_paragraph = existing_cell.GetFirstChild<Paragraph>();
                                            Run existing_run = existing_paragraph.GetFirstChild<Run>();
                                            //copio le proprietà del paragrafo
                                            ParagraphProperties paragraph_properties = existing_paragraph.GetFirstChild<ParagraphProperties>();
                                            Paragraph new_paragraph = new Paragraph();
                                            new_paragraph.AppendChild<ParagraphProperties>((ParagraphProperties)paragraph_properties.CloneNode(true));
                                            //copio le proprietà del run
                                            RunProperties run_properties = existing_run.GetFirstChild<RunProperties>();
                                            Run new_run = new Run();
                                            new_run.AppendChild<RunProperties>((RunProperties)run_properties.CloneNode(true));
                                            new_paragraph.Append(new_run);
                                            TableCell new_cell = new TableCell();
                                            new_run.AppendChild<TableCellProperties>((TableCellProperties)cell_properties.CloneNode(true));
                                            new_cell.Append(new_paragraph);
                                            if (imageindexcell.Contains(j))
                                            {
                                                ImageDimension dimension = null;
                                                if (_imagedimension.ContainsKey(placeholderList[j]))
                                                {
                                                    dimension = _imagedimension[placeholderList[j]];
                                                }
                                                InsertImageCellToTableRow(wordDoc, data_table.Rows[i][placeholderList[j]], new_cell, dimension);
                                            }
                                            else if (htmlindexcell.Contains(j))
                                            {
                                                InsertHtmlCellToTableRow(wordDoc, data_table.Rows[i][placeholderList[j]], new_cell);
                                            }
                                            else
                                            {
                                                InsertTextCellToTableRow(newtablerow, data_table.Rows[i][placeholderList[j]], new_cell);
                                            }
                                        }
                                        table.Append(newtablerow);
                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch(Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }
        /// <summary>
        /// Crea una TableCell con testo html, che viene inserita nella TableRow
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="newtablerow"></param>
        /// <param name="html"></param>
        private void InsertHtmlCellToTableRow(WordprocessingDocument wordDoc, object html, TableCell cell = null)
        {
            logger.MethodEnter();
            try
            {
                if (html is string htmlContent)
                {
                    // Converte HTML in elementi OpenXML
                    List<OpenXmlElement> convertedElements = ConvertHtmlToOpenXml(wordDoc.MainDocumentPart, htmlContent);
                    if(cell is null)
                        cell = new TableCell();
                    
                    Paragraph paragraph = new Paragraph();
                    paragraph.InsertAfter(convertedElements.First(), new Run());
                    foreach (var element in convertedElements.Skip(1))
                    {
                        paragraph.InsertAfter(element, convertedElements.First());
                    }
                    cell.Append(paragraph);

                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }

        /// <summary>
        /// Crea una TableCell con immagine, che viene inserita nella TableRow
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="newtablerow"></param>
        /// <param name="image"></param>
        /// <param name="tableCell"></param>
        /// <param name="dimension"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void InsertImageCellToTableRow(WordprocessingDocument wordDoc, object image,TableCell cell = null, ImageDimension? dimension = null)
        {
            logger.MethodEnter();
            try
            {
                //utilizzo di skiasharp per render il componente cross-platform
                SKBitmap skBitmap = null;
                if (image is string)// caso file path
                {
                    if (File.Exists(image.ToString()))
                    {
                        string fileimagePath = image.ToString();
                        skBitmap = LoadBitmapFromFile(fileimagePath);
                    }
                    else
                    {
                        //base64
                        try
                        {
                            if (!string.IsNullOrEmpty(image.ToString()))
                            {
                                byte[] buffer = System.Convert.FromBase64String(image.ToString());

                                skBitmap = SKBitmap.Decode(new MemoryStream(buffer));
                            }
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message, e);
                        }
                    }
                }
                if (image is SKBitmap b)// caso bitmap
                {
                    skBitmap = b;
                }
                if (skBitmap is null)
                    return;
                if (dimension is not null)
                {
                    skBitmap = skBitmap.Resize(new SKImageInfo(dimension.Width, dimension.Height), new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
                }
                // Converti lo SKBitmap in un MemoryStream (in formato PNG)
                using var skData = skBitmap.Encode(SKEncodedImageFormat.Png, 100);
                using var imageStream = new MemoryStream();
                skData.SaveTo(imageStream);
                imageStream.Position = 0;

                // Aggiungi l'immagine al documento come ImagePart
                ImagePart imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Png);
                imagePart.FeedData(imageStream);
                string relationshipId = wordDoc.MainDocumentPart.GetIdOfPart(imagePart);
                if(cell is null)
                     cell = new TableCell();
                
                cell.Append(new Paragraph(new Run(CreateImageElement(relationshipId, skBitmap.Width, skBitmap.Height))));
            }
            catch(Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }

        /// <summary>
        /// Crea una TableCell con testo, che viene inserita nella TableRow 
        /// </summary>
        /// <param name="newtablerow"></param>
        /// <param name="text"></param>
        private void InsertTextCellToTableRow(TableRow newtablerow, object text, TableCell cell = null)
        {
            logger.MethodEnter();
            try
            {
                if (text is string value)
                {
                    if(cell is null)
                    {

                        cell = new TableCell();
                        cell.Append(new Paragraph(new Run()));
                    }
                    Run run = cell.Descendants<Run>().FirstOrDefault();
                    if(run is null)
                    {
                        return;
                    }
                    if (value.Contains("\n") || value.Contains("<w:br />"))
                    {
                        string[] splitted = value.Split(new string[] { "\n", "<w:br />" }, StringSplitOptions.None);
                        int i = 0;
                        
                        foreach (string itm in splitted)
                        {
                            run.Append(new Text(itm));
                            if (i < splitted.Length - 1)
                            {

                                run.Append(new Break());
                            }
                            i++;
                        }
                    }
                    else
                    {
                        run.Append(new Text(value));
                    }
                    newtablerow.Append(cell);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }

        /// <summary>
        /// Sostituisce gli elementi Text con una stringa HTML
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="textElements"></param>
        /// <param name="html"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void ReplaceHtmlPlaceHolder(WordprocessingDocument wordDoc, List<Text> textElements, object html)
        {
            logger.MethodEnter();
            try
            {
                if(html is string htmlContent)
                {
                    // Converte HTML in elementi OpenXML
                    List<OpenXmlElement> convertedElements = ConvertHtmlToOpenXml(wordDoc.MainDocumentPart, htmlContent);

                    foreach (var textElement in textElements)
                    {
                        Run run = textElement.Parent as Run;
                        if (run is null) continue;

                        // Rimuove il testo del placeholder
                        run.RemoveAllChildren<Text>();

                        // Sostituisce con il contenuto HTML convertito
                        run.Parent.InsertAfter(convertedElements.First(), run);
                        foreach (var element in convertedElements.Skip(1))
                        {
                            run.Parent.InsertAfter(element, convertedElements.First());
                        }
                    }
                }

            }catch(Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }
        /// <summary>
        /// Converte il contenuto HTML in OpenXmlElement (Run, Paragraph, ecc.).
        /// </summary>
        /// <param name="mainDocumentPart">La parte principale del documento Word.</param>
        /// <param name="htmlContent">Il contenuto HTML da convertire.</param>
        /// <returns>Lista di elementi OpenXmlElement.</returns>
        private List<OpenXmlElement> ConvertHtmlToOpenXml(MainDocumentPart? mainDocumentPart, string htmlContent)
        {
            HtmlConverter converter = new HtmlConverter(mainDocumentPart);
            return converter.Parse(htmlContent).Cast<OpenXmlElement>().ToList();
        }

        /// <summary>
        /// Sostituisce il contenuto dell'elemento Text con l'immagine
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <param name="textElements"></param>
        /// <param name="image"></param>
        /// <param name="dimension"></param>
        /// <exception cref="NotImplementedException"></exception>
        private void ReplaceImagePlaceHolder(WordprocessingDocument wordDoc, List<Text> textElements, object image, ImageDimension? dimension)
        {
            logger.MethodEnter();
            try
            {
                //utilizzo di skiasharp per render il componente cross-platform
                SKBitmap skBitmap = null;
                if (image is string)// caso file path
                {
                    if (File.Exists(image.ToString()))
                    {
                        string fileimagePath = image.ToString();
                        skBitmap = LoadBitmapFromFile(fileimagePath);
                    }
                    else
                    {
                        //base64
                        try
                        {
                            if (!string.IsNullOrEmpty(image.ToString()))
                            {
                                byte[] buffer = System.Convert.FromBase64String(image.ToString());

                                skBitmap = SKBitmap.Decode(new MemoryStream(buffer));
                            }
                        }
                        catch (Exception e)
                        {
                            logger.Error(e.Message, e);
                        }
                    }
                }
                if (image is SKBitmap b)// caso bitmap
                {
                    skBitmap = b;
                }
                if (skBitmap is null)
                    return;
                if(dimension is not null)
                    skBitmap = skBitmap.Resize(new SKImageInfo(dimension.Width, dimension.Height), new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.Linear));
            
                // Converti lo SKBitmap in un MemoryStream (in formato PNG)
                using var skData = skBitmap.Encode(SKEncodedImageFormat.Png, 100);
                using var imageStream = new MemoryStream();
                skData.SaveTo(imageStream);
                imageStream.Position = 0;

                // Aggiungi l'immagine al documento come ImagePart
                ImagePart imagePart = wordDoc.MainDocumentPart.AddImagePart(ImagePartType.Png);
                imagePart.FeedData(imageStream);
                string relationshipId = wordDoc.MainDocumentPart.GetIdOfPart(imagePart);

                Run run = textElements.First().Parent as Run;
                run?.RemoveChild(run.Descendants<Text>().FirstOrDefault());
                run?.Append(CreateImageElement(relationshipId, skBitmap.Width, skBitmap.Height));
            }
            catch(Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }

        /// <summary>
        /// Sostituisce il contenuto dell'elemento Text
        /// </summary>
        /// <param name="textElements"></param>
        /// <param name="value"></param>
        /// <param name="chiave"></param>
        private void ReplaceTextPlaceHolder(List<Text> textElements, object value, string chiave)
        {
            logger.MethodEnter();
            try
            {
                string objstring = string.Empty;

                if (value is not null)
                    objstring = value.ToString();

                string fullText = string.Join("", textElements.Select(t => t.Text));
                if (fullText.Contains(chiave))
                {
                    // Cancella i vecchi <Text> esistenti
                    foreach (var textElement in textElements)
                    {
                        textElement.Text = string.Empty;
                    }
                    if (objstring.Contains("\n") || objstring.Contains("<w:br />"))
                    {
                        string[] splitted = objstring.Split(new string[] { "\n", "<w:br />" }, StringSplitOptions.None);
                        Run elem = textElements.First().Parent as Run;
                        if (elem is not null)
                        {
                            // Rimuove tutti i vecchi <Text> dentro il primo <Run>
                            elem.RemoveAllChildren<Text>();
                            int i = 0;
                            foreach (string itm in splitted)
                            {
                                Text txt = new Text(itm);
                                txt.Space = SpaceProcessingModeValues.Preserve;
                                elem.AppendChild(txt);
                                if (i < splitted.Length - 1)
                                {
                                    elem.Append(new Break());
                                }
                                i++;
                            }
                        }
                        

                    }
                    else
                    {
                        // Sostituisci il testo
                        fullText = fullText.Replace(chiave, objstring);
                        // Imposta il nuovo testo solo nel primo <Text> per mantenere la formattazione
                        textElements.First().Text = fullText;
                    }
                }
                    

            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex);
            }
        }
        

        private Drawing CreateImageElement(string relationshipId, long width, long height)
        {
            return new Drawing(
                new Inline(
                    new Extent() { Cx = width * 9525, Cy = height * 9525 },
                    new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DocProperties() { Id = 1, Name = "Picture" },
                    new NonVisualGraphicFrameDrawingProperties(new DW.GraphicFrameLocks() { NoChangeAspect = true }),
                    new DW.Graphic(
                        new DW.GraphicData(
                            new Picture(
                                new DW.NonVisualPictureProperties(
                                    new DW.NonVisualDrawingProperties() { Id = 0, Name = "New Image" },
                                    new DW.NonVisualPictureDrawingProperties()
                                ),
                                new DW.BlipFill(
                                    new DW.Blip() { Embed = relationshipId },
                                    new DW.Stretch(new DW.FillRectangle())
                                ),
                                new DW.ShapeProperties(
                                    new DW.Transform2D(
                                        new DW.Offset() { X = 0, Y = 0 },
                                        new DW.Extents() { Cx = width * 9525, Cy = height * 9525 }
                                    ),
                                    new DW.PresetGeometry(new DW.AdjustValueList()) { Preset = DW.ShapeTypeValues.Rectangle }
                                )
                            )
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                )
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );
        }
        private SKBitmap LoadBitmapFromFile(string filePath)
        {
            logger.MethodEnter();
            try
            {
                if(!File.Exists(filePath))
                {
                    return null;
                }
                // Apri un FileStream sul file specificato
                using FileStream stream = File.OpenRead(filePath);

                // Decodifica l'immagine e restituisci l'SKBitmap
                SKBitmap bitmap = SKBitmap.Decode(stream);
                return bitmap;
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message, ex );
                return null;
            }
            
        }
        
        #endregion
    }
    
}
