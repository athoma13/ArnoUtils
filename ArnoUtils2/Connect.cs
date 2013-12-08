using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

namespace ArnoUtils2
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
        private readonly List<Mapping> _mappings = new List<Mapping>();


		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
            _mappings.Add(Mapping.Map("CreateGetterAndSetters", Utilities.CreateGetterAndSetters, "global::ctrl+g,ctrl+g", "Create getters and setters"));
            _mappings.Add(Mapping.Map("CreateGettersOnly", Utilities.CreateGettersOnly, "global::ctrl+g,ctrl+o", "Create readonly getters"));
            _mappings.Add(Mapping.Map("CreateGetterAndSettersINotifiable", Utilities.CreateGetterAndSettersINotifiable, "global::ctrl+g,ctrl+i", "Create Notifyable getters and setters"));
            _mappings.Add(Mapping.Map("GuardNull", Utilities.GuardNull, "global::ctrl+g,ctrl+n", "Guard null"));
            _mappings.Add(Mapping.Map("ConstructorMacro", Utilities.ConstructorMacro, "global::ctrl+g,ctrl+c", "Constructor Logic"));
            _mappings.Add(Mapping.Map("MakeNewProjectItem", Utilities.MakeNewProjectItem, "global::ctrl+g,ctrl+m", "Make New Project Item"));
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;

            //Only if the setting up the add in for the first time will the commands be added.
            if (!CommandExists("MakeNewProjectItem"))
            {
                ClearCommands();

                var statusEnabled = (int)(vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled);
                foreach (var mapping in _mappings)
                {
                    var command = _applicationObject.Commands.AddNamedCommand(_addInInstance, mapping.CommandName, mapping.CommandName, mapping.CommandName, false, 0, null, statusEnabled);
                    command.Bindings = mapping.ShortCut;
                }
            }
		}

        private bool CommandExists(string commandName)
        {
            var fullCommandName = GetFullCommandName(commandName);
            var result = false;
            try
            {
                result = _applicationObject.Commands.Item(fullCommandName) != null;
            }
            catch
            {
            }

            return result;
        }

        private string GetCommandNamePrefix()
        {
            return GetType().FullName + ".";
        }

        private string GetFullCommandName(string commandName)
        {
            return GetCommandNamePrefix() + commandName;
        }


        private void ClearCommands()
        {
            var prefix = GetCommandNamePrefix();

            var commands = _applicationObject.Commands.OfType<Command>().Where(c => c.Name.Contains(prefix)).ToArray();
            foreach (var cmd in commands)
            {
                cmd.Delete();
            }


        }


		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}

        void IDTCommandTarget.Exec(string commandName,
           vsCommandExecOption executeOption,
           ref object varIn, ref object varOut, ref bool handled)
        {
            int lastDotIndex = commandName.LastIndexOf(".");
            if (lastDotIndex >= 0) commandName = commandName.Substring(lastDotIndex + 1);
            var mapping = _mappings.FirstOrDefault(m => m.CommandName == commandName);
            if (mapping != null)
            {
                _applicationObject.UndoContext.Open(mapping.UndoName, true);
                try
                {
                    mapping.Action(_applicationObject);
                }
                catch (Exception ex)
                {
                    _applicationObject.UndoContext.SetAborted();
                    MessageBox.Show(string.Format("There was an error executing command {0}. Message was: {1}", mapping.CommandName, ex.Message), this.GetType().FullName + " Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    handled = true;
                    if (_applicationObject.UndoContext.IsOpen)
                    {
                        _applicationObject.UndoContext.Close();
                    }
                }
            }
        }

        void IDTCommandTarget.QueryStatus(string CmdName, vsCommandStatusTextWanted NeededText, ref vsCommandStatus StatusOption, ref object CommandText)
        {
            StatusOption = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
        }

		
		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}