using EnvDTE80;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArnoUtils2
{
    public class Mapping
    {
        private Action<DTE2> _action;
        private string _shortcut;
        private string _commandName;
        private string _undoName;

        public string ShortCut
        {
            get { return _shortcut; }
        }

        /// <summary>
        /// Gets Action
        /// </summary>
        public Action<DTE2> Action
        {
            get { return _action; }
        }

        /// <summary>
        /// Gets CommandName
        /// </summary>
        public string CommandName
        {
            get { return _commandName; }
        }

        /// <summary>
        /// Gets UndoName
        /// </summary>
        public string UndoName
        {
            get { return _undoName; }
        }

        /// <summary>
        /// Gets Shortcut
        /// </summary>
        public string Shortcut
        {
            get { return _shortcut; }
        }

        public Mapping(string commandName, Action<DTE2> action, string shortcut, string undoName)
        {
            _action = action;
            _shortcut = shortcut;
            _commandName = commandName;
            _undoName = undoName;
        }

        public static Mapping Map(string commandName, Action<DTE2> action, string shortcut, string undoName)
        {
            return new Mapping(commandName, action, shortcut, undoName);
        }
    }
}
