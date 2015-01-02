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
 * File Description: This file contains the RestoreTextCmd class.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine.Undos
{
    // A restore text command.
    public class RestoreTextCmd : IUndoRedoCmd
    {
        private string _cellText, _cellName;

        public RestoreTextCmd(string cellText, string cellName)
        {
            _cellText = cellText;
            _cellName = cellName;
        }

        public IUndoRedoCmd Execute(Spreadsheet sheet)
        {
            Cell cell = sheet.GetCell(_cellName);
            string oldText = cell.Text;
            cell.Text = _cellText;
            return new RestoreTextCmd(oldText, _cellName);
        }
    }
}