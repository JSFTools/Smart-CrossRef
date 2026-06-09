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
        private Dictionary<int, CustomTaskPane> taskPaneDictionary;
        private readonly object _lockObject = new object(); // Prevents multi-window thread race conditions

        public bool IsTaskPaneGlobalEnabled { get; set; } = true;

        public Microsoft.Office.Tools.CustomTaskPane GetTaskPaneInstance()
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;
                if (activeWindow != null)
                {
                    // Extract the persistent OS window handle
                    int hwnd = activeWindow.Hwnd;

                    // Look up the task pane using the integer key
                    if (taskPaneDictionary.ContainsKey(hwnd))
                    {
                        return taskPaneDictionary[hwnd];
                    }
                }
            }
            catch
            {
                // Fallback for transient window initialization states (e.g., when Word is shutting down)
            }
            return null;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneDictionary = new Dictionary<int, CustomTaskPane>();

            // CLEANED: WindowActivate natively handles new files and opened files automatically.
            // Removing NewDocument and DocumentOpen prevents the double-creation race condition.
            this.Application.WindowActivate += Application_WindowActivate;

            // Replaced DocumentBeforeClose with WindowDeactivate for exact window-level death tracking
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;

            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            if (this.Application.Windows.Count > 0)
            {
                GetOrCreateTaskPane(this.Application.ActiveWindow);
            }
        }

        public CustomTaskPane GetOrCreateTaskPane(Word.Window window)
        {
            if (window == null) return null;

            lock (_lockObject)
            {
                // 1. Capture the unchangeable operating system window handle
                int hwnd = window.Hwnd;

                // 2. Look up the pane using the handle. If it exists, return it instantly!
                // This will now successfully find your collapsed pane when focus returns.
                if (taskPaneDictionary.ContainsKey(hwnd))
                {
                    return taskPaneDictionary[hwnd];
                }

                try
                {
                    var hostControl = new CrossRefPaneHostControl();
                    CustomTaskPane newPane = this.CustomTaskPanes.Add(hostControl, "SmartCrossRef", window);
                    newPane.Width = 320;

                    // 3. This line ONLY fires the very first time the document loads up
                    newPane.Visible = true;

                    hostControl.UpdateTheme();

                    if (newPane.Visible)
                    {
                        hostControl.RefreshContextualData();
                    }

                    // 4. Save using the window handle integer
                    taskPaneDictionary.Add(hwnd, newPane);
                    return newPane;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    return null;
                }
            }
        }

        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            // 1. Fetch or create the pane for this window
            CustomTaskPane pane = GetOrCreateTaskPane(Wn);
            if (pane != null)
            {
                var host = pane.Control as CrossRefPaneHostControl;
                if (host != null)
                {
                    // 2. Always keep the styling up to date
                    host.UpdateTheme();

                    // 3. NATIVE RESPECT: Check Word's current pane state.
                    // If the user collapsed it, pane.Visible will be false, so we skip the scan.
                    // If the user left it open, pane.Visible is true, so we refresh the data.
                    if (pane.Visible)
                    {
                        host.RefreshContextualData();
                    }
                }

                // 4. Keep your Ribbon toggle button matching the true state of this window
                UpdateRibbonButtonState(pane.Visible);
            }
        }

        private void UpdateRibbonButtonState(bool isVisible)
        {
            try
            {
                // Replace 'Ribbon1' and 'MyToggleButton' with your actual Ribbon and Control IDs
                if (Globals.Ribbons.Ribbon1 != null && Globals.Ribbons.Ribbon1.btnTogglePane != null)
                {
                    Globals.Ribbons.Ribbon1.btnTogglePane.Checked = isVisible;
                }
            }
            catch
            {
                // Handle edge cases where the Ribbon UI hasn't fully drawn yet
            }
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                Word.Window activeWin = Sel.Parent as Word.Window ?? this.Application.ActiveWindow;
                if (activeWin != null && taskPaneDictionary.TryGetValue(activeWin.Hwnd, out CustomTaskPane pane))
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

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            lock (_lockObject)
            {
                if (Doc == null) return;

                List<int> handlesToRemove = new List<int>();

                // 1. Find all task panes in our dictionary that belong to the closing document
                foreach (CustomTaskPane pane in this.CustomTaskPanes)
                {
                    try
                    {
                        Word.Window paneWindow = pane.Window as Word.Window;

                        // If the pane's window belongs to the document being closed, queue it up
                        if (pane.Window != null && paneWindow.Document == Doc)
                        {
                            // Find the matching handle in our dictionary
                            var match = taskPaneDictionary.FirstOrDefault(kvp => kvp.Value == pane);
                            if (match.Key != 0)
                            {
                                handlesToRemove.Add(match.Key);
                            }
                        }
                    }
                    catch
                    {
                        // Handle transient states during teardown
                    }
                }

                // 2. Safely purge only the dead window references from our tracking dictionary
                foreach (int hwnd in handlesToRemove)
                {
                    if (taskPaneDictionary.TryGetValue(hwnd, out CustomTaskPane pane))
                    {
                        try
                        {
                            this.CustomTaskPanes.Remove(pane);
                        }
                        catch { /* Word may have already disposed of the structural UI */ }

                        taskPaneDictionary.Remove(hwnd);
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