using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace SmartCrossRef
{
    public partial class ThisAddIn
    {
        // Dictionary tracking which task pane belongs to which Word window
        private Dictionary<Word.Window, CustomTaskPane> taskPaneDictionary;

        // Track the global toggle state from the Ribbon button
        public bool IsTaskPaneGlobalEnabled { get; set; } = true; // Default to true if you want it open on start
                                                                  // Inside ThisAddIn class structure:
        private Microsoft.Office.Tools.CustomTaskPane _myCustomTaskPane;

        public Microsoft.Office.Tools.CustomTaskPane GetTaskPaneInstance()
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;
                if (activeWindow != null && taskPaneDictionary.ContainsKey(activeWindow))
                {
                    return taskPaneDictionary[activeWindow];
                }
            }
            catch
            {
                // Fallback for transient window initialization states
            }
            return null;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneDictionary = new Dictionary<Word.Window, CustomTaskPane>();

            // Hook up window and document tracking events
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentOpen += Application_DocumentOpen;
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument += Application_NewDocument;

            // Clean up when a window/document closes
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;

            // Selection changed event
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            // Initialize the pane for the very first window that loads on startup
            if (this.Application.Windows.Count > 0)
            {
                GetOrCreateTaskPane(this.Application.ActiveWindow);
            }
        }

        // Central factory method to ensure every window gets its own pane instance safely
        public CustomTaskPane GetOrCreateTaskPane(Word.Window window)
        {
            if (window == null) return null;

            // If it already exists for this window, just return it
            if (taskPaneDictionary.ContainsKey(window))
            {
                return taskPaneDictionary[window];
            }

            try
            {
                // Create a completely fresh instance of your Host and WPF controls for THIS window
                var hostControl = new CrossRefPaneHostControl();

                CustomTaskPane newPane = this.CustomTaskPanes.Add(hostControl, "SmartCrossRef", window);
                newPane.Width = 320;

                // Sync the UI visibility to match your Ribbon toggle state
                newPane.Visible = IsTaskPaneGlobalEnabled;

                // Run initial operations
                hostControl.UpdateTheme();
                hostControl.RefreshContextualData();

                // Store it in our tracking dictionary
                taskPaneDictionary.Add(window, newPane);
                return newPane;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null; // Window is in a transient closing/opening state
            }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            CustomTaskPane pane = GetOrCreateTaskPane(Wn);
            if (pane != null)
            {
                pane.Visible = IsTaskPaneGlobalEnabled;

                // Refresh the data and theme for the active window layout
                var host = pane.Control as CrossRefPaneHostControl;
                host?.UpdateTheme();
                host?.RefreshContextualData();
            }
        }

        private void Application_DocumentOpen(Word.Document Doc)
        {
            GetOrCreateTaskPane(Doc.ActiveWindow);
        }

        private void Application_NewDocument(Word.Document Doc)
        {
            GetOrCreateTaskPane(Doc.ActiveWindow);
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            // Dynamically locate the task pane tied specifically to the window the user is currently typing in
            if (taskPaneDictionary.TryGetValue(Sel.Parent as Word.Window ?? this.Application.ActiveWindow, out CustomTaskPane pane))
            {
                var host = pane.Control as CrossRefPaneHostControl;
                host?.RefreshContextualData();
            }
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            // Prevent memory leaks: Remove destroyed panes from our registry cache
            List<Word.Window> windowsToRemove = new List<Word.Window>();
            foreach (var kvp in taskPaneDictionary)
            {
                try
                {
                    // If accessing the window fails, it means Word closed it
                    var test = kvp.Key.Caption;
                }
                catch
                {
                    windowsToRemove.Add(kvp.Key);
                }
            }

            foreach (var window in windowsToRemove)
            {
                if (taskPaneDictionary.TryGetValue(window, out CustomTaskPane pane))
                {
                    this.CustomTaskPanes.Remove(pane);
                    taskPaneDictionary.Remove(window);
                }
            }
        }

        // Synchronize all windows when clicking the toggle button
        public void SyncAllPanesVisibility(bool visible)
        {
            IsTaskPaneGlobalEnabled = visible;
            foreach (var pane in taskPaneDictionary.Values)
            {
                try
                {
                    pane.Visible = visible;
                }
                catch { /* Handle background window edge cases */ }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
