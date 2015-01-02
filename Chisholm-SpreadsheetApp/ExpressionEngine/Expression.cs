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
 * File Description: This file contains the Expression class.
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionEngine
{
    public class Expression
    {
        #region Classes

        // Our base node class which each node type inherits from.
        private abstract class Node
        {
            public Node Left, Right;
        }

        // Node representing an operation in an expression tree.
        private class OpNode : Node
        {
            public char Operator;

            public OpNode()
            {
            }

            public OpNode(char op)
            {
                Operator = op;
            }
        }

        // Node representing a variable in an expression tree.
        private class VarNode : Node
        {
            public string Name;

            public VarNode()
            {
            }

            public VarNode(string name)
            {
                Name = name;
            }
        }

        // Node representing a constant in an expression tree.
        private class ConstNode : Node
        {
            public double Value;

            public ConstNode()
            {
            }

            public ConstNode(double num)
            {
                Value = num;
            }
        }

        #endregion

        #region Fields
        
        private Node _root;
        private string _expString;
        private Dictionary<string, double> _vars;
        public readonly static char[] SupportedOps = { '+', '-', '*', '/', '^' };

        #endregion

        #region Properties

        public string ExpString
        {
            get { return _expString; }
            set 
            {
                // On set, clear variable dictionary and recompile tree as well as setting the string.
                _expString = value;
                _vars.Clear();
                _root = Compile(_expString);
            }
        }

        #endregion

        #region Constructors

        public Expression()
        {
            _vars = new Dictionary<string, double>();
        }

        public Expression(string expString)
        {
            string result = "";
            if (expString.Contains(' '))
            {
                foreach (char letter in expString)
                {
                    if (letter != ' ')
                    {
                        result += letter;
                    }
                }
            }
            else
            {
                result = expString;
            }

            // Not using ExpString's set because it would unnecessarily clear the dictionary.
            _expString = result;
            _vars = new Dictionary<string, double>();
            _root = Compile(_expString);
        }

        #endregion

        #region Methods

        /// <summary>
        /// Compiles an expression string.
        /// </summary>
        /// <param name="s">The expression string to be compiled.</param>
        /// <returns>A subtree (as a Node) representing the expression.</returns>
        private Node Compile(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return null;
            }

            // Start looking for a right parenthesis if a left parenthesis is found at the start of the string.
            if (s[0] == '(')
            {
                // We will need to keep track of how many left and right parentheses we encounter.
                int counter = 0;

                // Search through string...
                for (int i = 0; i < s.Length; i++)
                {
                    // If we find another left parenthesis, count it positively,
                    // If we find any right parenthesis, count it negatively.
                    if (s[i] == '(')
                    {
                        counter++;
                    }
                    else if (s[i] == ')')
                    {
                        counter--;

                        // If we hit a count of 0, we've reached the correct 
                        // right parenthesis matching the first left parenthesis.
                        if (counter == 0)
                        {
                            // If not at the end of the string, break (since compilation will resume as normal).
                            // If at the end of the string, we need to compile everything but the first and last parentheses.
                            if (s.Length - 1 != i)
                            {
                                break;
                            }
                            else
                            {
                                return Compile(s.Substring(1, s.Length - 2));
                            }
                        }
                    }
                }
            }

            // Loop through operators going from low precedence to high, with 
            // commutative operators favored over other operators at the same precedence.
            char[] ops = Expression.SupportedOps;
            foreach (char op in ops)
            {
                // Compile the expression based on the current operation.
                // Only return subtree if non-null.
                Node n = Compile(s, op);
                if (n != null) return n;
            }

            // If we get here, then the subtree was null, meaning that the expression
            // is a leaf node (so is either a ConstNode or a VarNode). Return it as such.
            double num;
            if (double.TryParse(s, out num))
            {
                return new ConstNode(num);
            }
            else
            {
                // Initialize the variable in the dictionary when found.
                _vars[s] = 0;
                return new VarNode(s);
            }
        }

        /// <summary>
        /// Recursively compile an expression based on an operation.
        /// </summary>
        /// <param name="exp">The expression string to be compiled.</param>
        /// <param name="op">The operation that compilation is based on.</param>
        /// <returns> A subtree (as a Node) representing the expression. </returns>
        private Node Compile(string exp, char op)
        {
            int pcounter = 0;
            bool hasTerminated = false;

            // Our loop variable (default to end of exp for left associative).
            int i = exp.Length - 1;

            // Whether the operation is right associative (default to left associative).
            bool isRightAssociative = false;
            
            // If our op is exponent, it is right associative and we will start at the beginning of exp.
            if (op == '^')
            {
                isRightAssociative = true;
                i = 0;
            }

            // While the loop hasn't been terminated...
            while (!hasTerminated)
            {
                // Count left parentheses positively if left associative, negatively if right associative.
                // Count right parentheses negatively if left associative, positively if right associative.
                if (exp[i] == '(')
                {
                    if (isRightAssociative) pcounter--;
                    else pcounter++;
                }
                else if (exp[i] == ')')
                {
                    if (isRightAssociative) pcounter++;
                    else pcounter--;
                }

                //If we've reached the current operation and it is not in parentheses...
                if (pcounter == 0 && exp[i] == op)
                {
                    // Create and return a subtree with the current op (as an OpNode) being the root, and the 
                    // left and right expressions (as their own compiled subtrees) being the Left and Right children.
                    OpNode on = new OpNode(op);
                    on.Left = Compile(exp.Substring(0, i));
                    on.Right = Compile(exp.Substring(i + 1));
                    return on;
                }

                // Determine whether to terminate and which direction to increment based on associativity.
                if (isRightAssociative)
                {
                    if (i == exp.Length - 1) hasTerminated = true;
                    i++;
                }
                else
                {
                    if (i == 0) hasTerminated = true;
                    i--;
                }
            }

            // If the expression has unbalanced parentheses, throw an exception,
            // else return null, as nothing can currently be done to this expression.
            if (pcounter != 0)
            {
                throw new Exception();
            }

            return null;
        }

        /// <summary>
        /// Evaluates a node.
        /// </summary>
        /// <param name="n">The node to be evaluated.</param>
        /// <returns>The evaluated value of the node.</returns>
        private double Eval(Node n)
        {
            ConstNode cn = n as ConstNode;
            if (cn != null)
            {
                return cn.Value;
            }

            VarNode vn = n as VarNode;
            if (vn != null)
            {
                return _vars[vn.Name];
            }

            // If OpNode, recursively evaluate Left and Right subtrees and perform operation on them.
            OpNode on = n as OpNode;
            if (on != null)
            {
                switch (on.Operator)
                {
                    case '+':
                        return Eval(on.Left) + Eval(on.Right);
                    case '-':
                        return Eval(on.Left) - Eval(on.Right);
                    case '*':
                        return Eval(on.Left) * Eval(on.Right);
                    case '/':
                        return Eval(on.Left) / Eval(on.Right);
                    case '^':
                        return Math.Pow(Eval(on.Left), Eval(on.Right));
                }
            }

            // If we get here, either the node type is not supported, or the operation is not supported.
            throw new NotSupportedException();
        }

        /// <summary>
        /// Evaluates the expression. Assumes expression has already been properly compiled.
        /// </summary>
        /// <returns>The evaluated value of the expression.</returns>
        public double Evaluate()
        {
            return Eval(_root);
        }

        /// <summary>
        /// Sets a variable's value for the expression.
        /// </summary>
        /// <param name="name">The name of the variable.</param>
        /// <param name="value">The value of the variable.</param>
        public void SetVariable(string name, double value)
        {
            _vars[name] = value;
        }
        
        /// <summary>
        /// Returns all variable names in this expression.
        /// </summary>
        /// <returns>A string array containing all variable names in this expression.</returns>
        public string[] GetAllVariables()
        {
            return _vars.Keys.ToArray();
        }

        #endregion
    }
}