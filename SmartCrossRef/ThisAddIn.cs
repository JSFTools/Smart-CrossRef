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
        private Dictionary<Word.Window, CustomTaskPane> taskPaneDictionary;
        private readonly object _lockObject = new object(); // Prevents multi-window thread race conditions

        public bool IsTaskPaneGlobalEnabled { get; set; } = true;

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

            // CLEANED: WindowActivate natively handles new files and opened files automatically.
            // Removing NewDocument and DocumentOpen prevents the double-creation race condition.
            this.Application.WindowActivate += Application_WindowActivate;

            // Replaced DocumentBeforeClose with WindowDeactivate for exact window-level death tracking
            this.Application.WindowDeactivate += Application_WindowDeactivate;

            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            if (this.Application.Windows.Count > 0)
            {
                GetOrCreateTaskPane(this.Application.ActiveWindow);
            }
        }

        public CustomTaskPane GetOrCreateTaskPane(Word.Window window)
        {
            if (window == null) return null;

            // Thread-safe wrapper ensures two window events cannot build a pane simultaneously
            lock (_lockObject)
            {
                if (taskPaneDictionary.ContainsKey(window))
                {
                    return taskPaneDictionary[window];
                }

                try
                {
                    var hostControl = new CrossRefPaneHostControl();
                    CustomTaskPane newPane = this.CustomTaskPanes.Add(hostControl, "SmartCrossRef", window);
                    newPane.Width = 320;
                    newPane.Visible = IsTaskPaneGlobalEnabled;

                    hostControl.UpdateTheme();
                    hostControl.RefreshContextualData();

                    taskPaneDictionary.Add(window, newPane);
                    return newPane;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return null; // Window is in a transient state
                }
            }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            CustomTaskPane pane = GetOrCreateTaskPane(Wn);
            if (pane != null)
            {
                pane.Visible = IsTaskPaneGlobalEnabled;

                var host = pane.Control as CrossRefPaneHostControl;
                host?.UpdateTheme();
                host?.RefreshContextualData();
            }
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                Word.Window activeWin = Sel.Parent as Word.Window ?? this.Application.ActiveWindow;
                if (activeWin != null && taskPaneDictionary.TryGetValue(activeWin, out CustomTaskPane pane))
                {
                    var host = pane.Control as CrossRefPaneHostControl;
                    host?.RefreshContextualData();
                }
            }
            catch
            {
                /* Prevent transient selection errors while dragging selections across layouts */
            }
        }

        // FIX: Tracks exact window dismissal cleanly without leaking ghost elements
        private void Application_WindowDeactivate(Word.Document Doc, Word.Window Wn)
        {
            lock (_lockObject)
            {
                List<Word.Window> deadWindows = new List<Word.Window>();

                // Scan for the deactivated window or any missing window hooks
                foreach (var kvp in taskPaneDictionary)
                {
                    try
                    {
                        // Check if the specific window handle match occurs
                        if (kvp.Key == Wn || kvp.Key.Caption == null)
                        {
                            deadWindows.Add(kvp.Key);
                        }
                    }
                    catch
                    {
                        // Access failed = Window was structurally destroyed by the user
                        deadWindows.Add(kvp.Key);
                    }
                }

                foreach (var window in deadWindows)
                {
                    if (taskPaneDictionary.TryGetValue(window, out CustomTaskPane pane))
                    {
                        try
                        {
                            this.CustomTaskPanes.Remove(pane);
                        }
                        catch { /* Already removed by Word internally */ }

                        taskPaneDictionary.Remove(window);
                    }
                }
            }
        }

        public void SyncAllPanesVisibility(bool visible)
        {
            IsTaskPaneGlobalEnabled = visible;
            lock (_lockObject)
            {
                foreach (var pane in taskPaneDictionary.Values)
                {
                    try
                    {
                        pane.Visible = visible;
                    }
                    catch { /* Handle closed background window edge cases */ }
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}