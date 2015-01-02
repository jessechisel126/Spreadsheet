/*
 * Programmer: Jesse Chisholm 11278684
 * Program: Excell Program Part 4 (HW10)
 * Date Started: 9/26/14
 * Date Completed: 11/21/14
 * Hours Worked: 50+
 * Collaboration: Light collab with Ben Tatham, Tim Vierow, and Zach Allen, and often referenced MSDN.
 * 
 * Program Description: This program is a simple unfinished version of Excell using Winforms and C#.
 * 
 * File Description: This file contains the Workbook class.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace SpreadsheetEngine
{
    public class Workbook
    {
        #region Fields

        private List<Spreadsheet> _sheets = new List<Spreadsheet>();
        private int _activeSheetIndex = 0;
        public UndoRedoSystem UndoRedo = new UndoRedoSystem();
        public event PropertyChangedEventHandler WorkbookSheetChanged;

        #endregion

        #region Properties

        /// <summary>
        /// Gets the active sheet from the book.
        /// </summary>
        /// <returns>The sheet at the active sheet index.</returns>
        public Spreadsheet ActiveSheet
        {
            get
            {
                return _sheets[_activeSheetIndex];
            }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new empty workbook with one sheet of default size 10x10.
        /// </summary>
        public Workbook()
            : this(10, 10)
        {
        }

        /// <summary>
        /// Creates a workbook with one spreadsheet of specified size.
        /// </summary>
        /// <param name="rows">Number of rows in the spreadsheet.</param>
        /// <param name="cols">Number of columns in the spreadsheet.</param>
        public Workbook(int rows, int cols)
        {
            AddSheet(new Spreadsheet(rows, cols));
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add a sheet to the workbook.
        /// </summary>
        /// <param name="sheet">The sheet to add.</param>
        public void AddSheet(Spreadsheet sheet)
        {
            _sheets.Add(sheet);
            sheet.SpreadsheetCellChanged += OnSpreadsheetCellChanged;
        }

        /// <summary>
        /// Sets the active sheet for the book.
        /// </summary>
        /// <param name="index">The new value of the active sheet index.</param>
        /* FOR FUTURE USE
        public void SetActiveSheet(int index)
        {
            _activeSheetIndex = index;
        }
         */

        /// <summary>
        /// Saves workbook data as XML to file using streams.
        /// </summary>
        /// <param name="stream">The file stream used to save.</param>
        /// <returns>Whether or not saving was successful.</returns>
        public bool Save(Stream stream)
        {
            // Set up the writer.
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.UTF8;
            settings.NewLineChars = "\r\n";
            settings.NewLineOnAttributes = false;
            settings.Indent = true;

            // Create the XmlWriter to write to the stream with our settings.
            XmlWriter writer = XmlWriter.Create(stream, settings);
            if (writer == null)
            {
                return false;
            }

            writer.WriteStartElement("Workbook");

            foreach (Spreadsheet sheet in _sheets)
            {
                sheet.Save(writer);
            }

            writer.WriteEndElement();
            writer.Close();

            return true;
        }

        /// <summary>
        /// Loads workbook data from XML in file using streams.
        /// </summary>
        /// <param name="stream">The file stream used to save.</param>
        /// <returns>Whether or not loading was successful.</returns>
        public bool Load(Stream stream)
        {
            XDocument document = null;

            try
            {
                document = XDocument.Load(stream);
            }
            catch (Exception)
            {
                return false;
            }

            if (document == null)
            {
                return false;
            }

            // Clear the existing data in the spreadsheet before loading.
            _sheets[0].Clear();

            // Load each sheet.
            XElement root = document.Root;
            foreach (XElement child in root.Elements("Spreadsheet"))
            {
                _sheets[0].Load(child);
            }

            // Clear the undo and redo stacks.
            UndoRedo.Clear();

            return true;
        }

        /// <summary>
        /// The SpreadsheetChanged event handler.
        /// </summary>
        /// <param name="sender">The object that fired the SpreadsheetChanged event.</param>
        /// <param name="e">The event arguments for the Property Changed event.</param>
        public void OnSpreadsheetCellChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Value")
            {
                WorkbookSheetChanged(sender, e);
            }
            else if (e.PropertyName == "BackColor")
            {
                WorkbookSheetChanged(sender, e);
            }
        }

        #endregion
    }
}