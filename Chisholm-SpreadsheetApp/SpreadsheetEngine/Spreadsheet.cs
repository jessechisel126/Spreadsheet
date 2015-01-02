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
 * File Description: This file contains the Spreadsheet class, used in the Workbook class.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Xml;
using System.Xml.Linq;
using ExpressionEngine;

namespace SpreadsheetEngine
{
    public class Spreadsheet
    {
        #region Classes

        //An Instantiable Cell definition.
        private class InstanceCell : Cell
        {
            public InstanceCell(int row, int col)
                : base(row, col)
            {
            }

            public void SetValue(string value)
            {
                _value = value;
            }
        }


        #endregion

        #region Fields

        private Cell[,] _cells;
        private Dictionary<string, HashSet<string>> _dependencies;
        public event PropertyChangedEventHandler SpreadsheetCellChanged;

        #endregion

        #region Properties

        public int RowCount
        {
            get { return _cells.GetLength(0); }
        }

        public int ColumnCount
        {
            get { return _cells.GetLength(1); }
        }

        #endregion

        #region Constructors

        /// <summary>
        /// Constructs a spreadsheet object of designated size.
        /// </summary>
        /// <param name="rows">Number of rows in spreadsheet.</param>
        /// <param name="cols">Number of columns in spreadsheet.</param>
        public Spreadsheet(int rows, int cols)
        {
            _cells = new Cell[rows, cols];

            _dependencies = new Dictionary<string, HashSet<string>>();

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    Cell currCell = new InstanceCell(i, j);
                    currCell.PropertyChanged += OnPropertyChanged;
                    _cells[i, j] = currCell;
                }
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Determines whether there is a circular reference.
        /// </summary>
        /// <param name="startCellName">The starting cell to search from. 
        /// This does not change on recursive calls.</param>
        /// <param name="currCellName">The current cell we are checking for a circular reference.</param>
        /// <returns>Whether there is a circular reference.</returns>
        public bool HasCircularReference(string startCellName, string currCellName)
        {
            // If in our current traversal we've run into our 
            // starting cell again, that's a circular reference.
            if (startCellName == currCellName)
            {
                return true;
            }

            // If the dependency entry doesn't exist, there's no circular reference.
            if (!_dependencies.ContainsKey(currCellName))
            {
                return false;
            }

            // Recursively check for circular references in all dependent cells.
            foreach (string dependent in _dependencies[currCellName])
            {
                if (HasCircularReference(startCellName, dependent))
                {
                    return true;
                }
            }
            
            return false;
        }

        /// <summary>
        /// Saves the spreadsheet to xml.
        /// </summary>
        /// <param name="writer">The XmlWriter we are using for writing xml.</param>
        public void Save(XmlWriter writer)
        {
            writer.WriteStartElement("Spreadsheet");

            var cellsToWrite =
                from Cell cell in _cells
                where !cell.HasDefaults
                select cell;

            foreach (Cell cell in cellsToWrite)
            {
                writer.WriteStartElement("Cell");
                writer.WriteAttributeString("Name", cell.Name);

                writer.WriteElementString("Text", cell.Text);
                writer.WriteElementString("BackColor", cell.BackColor.ToString());

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        }

        /// <summary>
        /// Loads a spreadsheet from an xml element.
        /// </summary>
        /// <param name="spreadsheetElement">The Xml element corresponding to the sreadsheet you want to Load.</param>
        public void Load(XElement spreadsheetElement)
        {
            if ("Spreadsheet" != spreadsheetElement.Name)
            {
                return;
            }

            foreach (XElement child in spreadsheetElement.Elements("Cell"))
            {
                Cell cell = GetCell(child.Attribute("Name").Value);

                // Only edit existing cells.
                if (cell == null) { continue; }

                // Load and set text.
                var textElement = child.Element("Text");
                if (textElement != null)
                {
                    cell.Text = textElement.Value;
                }

                //Load and set background color.
                var bgElement = child.Element("BackColor");
                if (bgElement != null)
                {
                    cell.BackColor = int.Parse(bgElement.Value);
                }
            }
        }

        /// <summary>
        /// Clears the spreadsheet.
        /// </summary>
        public void Clear()
        {
            for (int i = 0; i < RowCount; i++)
            {
                for (int j = 0; j < ColumnCount; j++)
                {
                    if (!_cells[i, j].HasDefaults)
                    {
                        _cells[i, j].Clear();
                    }
                }
            }
        }

        /// <summary>
        /// Uses a Cell's text to determine its value.
        /// </summary>
        /// <param name="location">The string location of the Cell. Ex: "A1" evaluates the Cell at (0,0).</param>
        private void EvaluateCell(string location)
        {
            EvaluateCell(GetCell(location));
        }

        /// <summary>
        /// Uses a Cell's text to determine its value.
        /// </summary>
        /// <param name="row">The row of the Cell to evaluate.</param>
        /// <param name="col">The column of the Cell to evaluate.</param>
        private void EvaluateCell(int row, int col)
        {
            EvaluateCell(GetCell(row, col));
        }

        /// <summary>
        /// Uses a Cell's text to determine its value.
        /// </summary>
        /// <param name="cell">The Cell to evaluate.</param>
        private void EvaluateCell(Cell cell)
        {
            // Cast as InstanceCell since we'll need SetValue.
            InstanceCell iCell = cell as InstanceCell;

            if (string.IsNullOrEmpty(iCell.Text))
            {
                // Set value to empty and fire value change event.
                iCell.SetValue("");
                SpreadsheetCellChanged(cell, new PropertyChangedEventArgs("Value"));
            }
            else if (iCell.Text[0] == '=' && iCell.Text.Length > 1)
            {
                bool error = false;

                // Clip off the '=' for our expression string.
                string expString = iCell.Text.Substring(1);

                // Create expression from the expression string.
                Expression exp = new Expression(expString);

                // Get the variables used by the expression.
                string[] variables = exp.GetAllVariables();
                
                // Setting each variable in the expression.
                foreach (string variableName in variables)
                {
                    // Check if there's a self reference from this variable.
                    if (variableName == iCell.Name)
                    {
                        PrintErrorToCell(iCell, variableName, "SELFREF");
                        error = true;
                        break;
                    }

                    // Check if the cell represented by the variable exists.
                    if (GetCell(variableName) == null)
                    {
                        PrintErrorToCell(iCell, variableName, "NAME");
                        error = true;
                        break;
                    }

                    // Set the variable in the expression.
                    SetExpressionVariable(exp, variableName);

                    // Check if there's a circular reference from this variable.
                    if (HasCircularReference(variableName, iCell.Name))
                    {
                        PrintErrorToCell(iCell, variableName, "CIRCREF");
                        error = true;
                        break;
                    }
                }

                // Return if an error occurred.
                if (error) return;

                // Set the value to the evaluation of the expression and fire value change event.
                iCell.SetValue(exp.Evaluate().ToString());
                SpreadsheetCellChanged(cell, new PropertyChangedEventArgs("Value"));
            }
            else
            {
                // Just set value to text and fire value change event.
                iCell.SetValue(iCell.Text);
                SpreadsheetCellChanged(cell, new PropertyChangedEventArgs("Value"));
            }

            // Recursively evaluate all dependent cells
            if (_dependencies.ContainsKey(iCell.Name))
            {
                foreach (string name in _dependencies[iCell.Name])
                {
                    EvaluateCell(name);
                }
            }
        }

        /// <summary>
        /// Sets a variable in an expression based on the spreadsheet.
        /// </summary>
        /// <param name="exp">The expression we are manipulating.</param>
        /// <param name="variableName">The cell for the variable we are setting in the expression.</param>
        private void SetExpressionVariable(Expression exp, string variableName)
        {
            Cell variableCell = GetCell(variableName);
            double value;
            if (string.IsNullOrEmpty(variableCell.Value))
            {
                //...set empty cells to 0.
                exp.SetVariable(variableCell.Name, 0);
            }
            else if (!double.TryParse(variableCell.Value, out value))
            {
                //...set non-value cells to 0.
                exp.SetVariable(variableName, 0);
            }
            else
            {
                //...just set normally otherwise.
                exp.SetVariable(variableName, value);
            }
        }

        /// <summary>
        /// Logs dependencies in _dependencies member.
        /// </summary>
        /// <param name="cellName">A cell's name.</param>
        /// <param name="variablesUsed">All variables referenced in the cell.</param>
        private void LogDependencies(string cellName, string[] variablesUsed)
        {
            // Log dependencies from this expression.
            foreach (string variableName in variablesUsed)
            {
                if (!_dependencies.ContainsKey(variableName))
                {
                    // Build dictionary entry for this variable name.
                    _dependencies[variableName] = new HashSet<string>();
                }

                // Add this cell name to dependencies for this variable name.
                _dependencies[variableName].Add(cellName);
            }
        }

        /// <summary>
        /// Removes all dependencies involving the cell name as a value.
        /// </summary>
        /// <param name="cellName">The cell name we are removing.</param>
        private void RemoveDependencies(string cellName)
        {
            /*
            foreach (string key in _dependencies.Keys)
            {
                if (_dependencies[key].Contains(cellName))
                {
                    _dependencies[key].Remove(cellName);
                }
            }
             */

            // Log all keys in which this cell name shows up as an associated value.
            List<string> keysToModify = new List<string>();

            foreach (string key in _dependencies.Keys)
            {
                if (_dependencies[key].Contains(cellName))
                {
                    keysToModify.Add(key);
                }
            }

            // Remove this cell's name from the list at each key we are modifying.
            foreach (string key in keysToModify)
            {
                HashSet<string> set = _dependencies[key];

                if (set.Contains(cellName))
                {
                    set.Remove(cellName);
                }
            }
        }

        /// <summary>
        /// Tell the world of crimes against computerkind.
        /// </summary>
        /// <param name="reporter">The nosey rat.</param>
        /// <param name="criminal">The guilty party.</param>
        /// <param name="crime">The heinous deed.</param>
        private void PrintErrorToCell(InstanceCell reporter, string criminal, string crime)
        {
            reporter.SetValue("#" + crime + ": Cell " + criminal);
            SpreadsheetCellChanged(reporter as Cell, new PropertyChangedEventArgs("Value"));
        }

        /// <summary>
        /// Gets a Cell at the given row and columns indices.
        /// </summary>
        /// <param name="row">The row of the Cell to get.</param>
        /// <param name="col">The column of the Cell to get.</param>
        /// <returns>The Cell at the given row and column indices.</returns>
        public Cell GetCell(int row, int col)
        {
            return _cells[row, col];
        }

        /// <summary>
        /// Gets a Cell at the given location string.
        /// </summary>
        /// <param name="location">The string location of the Cell. Ex: "A1", "D47", etc.</param>
        /// <returns>The Cell at the given location if the location is valid, null otherwise.</returns>
        public Cell GetCell(string location)
        {
            char letter = location[0];
            if (!Char.IsLetter(letter))
            {
                // If first character is not a letter, return null.
                return null;
            }

            Int16 number;
            if (!Int16.TryParse(location.Substring(1), out number))
            {
                // If the rest of the string is not a number, return null.
                return null;
            }

            Cell result;
            try
            {
                result = GetCell(number - 1, letter - 'A');
            }
            catch (Exception)
            {
                // If the given location does not exist in the spreadsheet, return null.
                return null;
            }

            // If we get here, the location was valid, and we can return the resulting Cell.
            return result;
        }

        /// <summary>
        /// The PropertyChanged event handler.
        /// </summary>
        /// <param name="sender">The object which fired the PropertyChanged event.</param>
        /// <param name="e">The event arguments for the PropertyChanged event.</param>
        public void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Text")
            {
                InstanceCell iCell = sender as InstanceCell;

                // Remove all dependencies involving this cell name as a value.
                RemoveDependencies(iCell.Name);

                // If the cell contains an expression...
                if (iCell.Text != "" && iCell.Text[0] == '=' && iCell.Text.Length > 1)
                {
                    //...clip off the '=' for our expression string and create expression.
                    Expression exp = new Expression(iCell.Text.Substring(1));

                    // Log dependencies from this expression.
                    LogDependencies(iCell.Name, exp.GetAllVariables());
                }

                // Evaluate the current cell.
                // Note: This will recursively call EvaluateCell on cells that
                // are dependent on the sender's value
                EvaluateCell(sender as Cell);
            }
            else if (e.PropertyName == "BackColor")
            {
                SpreadsheetCellChanged(sender, new PropertyChangedEventArgs("BackColor"));
            }
        }

        #endregion
    }
}