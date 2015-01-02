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
 * File Description: This file contains Undo/Redo related class definitions.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetEngine
{
    // An undo/redo command interface.
    public interface IUndoRedoCmd
    {
        IUndoRedoCmd Execute(Spreadsheet sheet);
    }

    // A collection of undo/redo commands.
    public class UndoRedoCollection
    {
        // Description of the overall action and a list of the sub-actions.
        public string Description;
        private IUndoRedoCmd[] _cmds;

        public UndoRedoCollection()
        {
        }

        public UndoRedoCollection(IUndoRedoCmd[] cmds, string description)
        {
            _cmds = cmds;
            Description = description;
        }

        public UndoRedoCollection(List<IUndoRedoCmd> cmds, string description)
        {
            _cmds = cmds.ToArray();
            Description = description;
        }

        /// <summary>
        /// Executes each action in the collection.
        /// </summary>
        /// <param name="sheet">The sheet to work on.</param>
        /// <returns>A list of opposite actions.</returns>
        public UndoRedoCollection Execute(Spreadsheet sheet)
        {
            List<IUndoRedoCmd> cmdList = new List<IUndoRedoCmd>();

            foreach (IUndoRedoCmd cmd in _cmds)
            {
                cmdList.Add(cmd.Execute(sheet));
            }

            return new UndoRedoCollection(cmdList.ToArray(), this.Description);
        }
    }

    // A complete undo/redo system.
    public class UndoRedoSystem
    {
        private Stack<UndoRedoCollection> _undos = new Stack<UndoRedoCollection>();
        private Stack<UndoRedoCollection> _redos = new Stack<UndoRedoCollection>();

        //Can you undo? CAN YOU REALLY!?
        public bool CanUndo
        {
            get { return _undos.Count != 0; }
        }

        //Can you redo? CAN YOU REALLY!?
        public bool CanRedo
        {
            get { return _redos.Count != 0; }
        }

        //The description of the next available undo.
        public string UndoDescription
        {
            get
            {
                if (CanUndo) return _undos.Peek().Description;
                return "";
            }
        }

        //The description of the next available redo.
        public string RedoDescription
        {
            get
            {
                if (CanRedo) return _redos.Peek().Description;
                return "";
            }
        }

        /// <summary>
        /// Pushes a call onto the undo stack.
        /// </summary>
        /// <param name="undos">The call to be pushed.</param>
        public void AddUndos(UndoRedoCollection undos)
        {
            _undos.Push(undos);
            _redos.Clear();
        }

        /// <summary>
        /// Performs an undo.
        /// </summary>
        /// <param name="sheet">The sheet to perform the undo on.</param>
        public void Undo(Spreadsheet sheet)
        {
            UndoRedoCollection actions = _undos.Pop();
            _redos.Push(actions.Execute(sheet));
        }

        /// <summary>
        /// Performs a redo.
        /// </summary>
        /// <param name="sheet">The sheet to perform the redo on.</param>
        public void Redo(Spreadsheet sheet)
        {
            UndoRedoCollection actions = _redos.Pop();
            _undos.Push(actions.Execute(sheet));
        }

        /// <summary>
        /// Clears the undo and redo stacks.
        /// </summary>
        public void Clear()
        {
            _undos.Clear();
            _redos.Clear();
        }
    }
}
