using Microsoft.Office.Tools.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Windows.LogicalTreeHelper;
using Word = Microsoft.Office.Interop.Word;

namespace SmartCrossRef
{

    public class CrossRefTargetItem
    {
        public string DisplayText { get; set; } // E.g., "Heading 1: Introduction"
        public string Category { get; set; }    // "Heading", "Table", "Figure", "Footnote"
        public Word.Range WordRange { get; set; }    // Reference to jump to if clicked

        public override string ToString()
        {
            return $"[{Category}] {DisplayText}";
        }
    }

    public enum ScanScope
    {
        NearbyOnly,
        FullDocument
    }

    /// <summary>
    /// Interaction logic for winCrossRefPane.xaml
    /// </summary>
    public partial class winCrossRefPane : UserControl
    {
        private bool _isPaneCollapsed = false;
        private int _originalPaneWidth = 320; // Default sizing baseline

        // Master list tracking everything scanned during the last background document sweep
        private List<CrossRefTargetItem> _allDocumentItems = new List<CrossRefTargetItem>();

        public winCrossRefPane()
        {
            InitializeComponent();
        }

        private void BtnCollapseToggle_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Office.Tools.CustomTaskPane taskPane = Globals.ThisAddIn.GetTaskPaneInstance();

                if (!_isPaneCollapsed)
                {
                    // --- ACTION: COLLAPSE ---
                    if (taskPane != null) _originalPaneWidth = taskPane.Width;

                    // 1. Hide the primary lower controls workspace completely
                    MainContentGrid.Visibility = Visibility.Collapsed;
                    HeaderTitle.Visibility = Visibility.Collapsed;
                    btnAbout.Visibility = Visibility.Collapsed;

                    // 2. Adjust content values & flip orientation variables
                    BtnCollapseToggle.Content = "▼ Smart CrossRef";
                    ButtonRotation.Angle = 90;

                    // Give it fresh breathing margins for its tall, vertical state
                    BtnCollapseToggle.Padding = new Thickness(16, 6, 16, 6);
                    BtnCollapseToggle.Margin = new Thickness(0, 15, 0, 15);

                    // 3. Move the button alignment handles to center fill the newly opened layout space
                    Grid.SetColumn(BtnCollapseToggle, 0);
                    Grid.SetColumnSpan(BtnCollapseToggle, 2);

                    // FORCE THE CONTAINERS TO EXPAND VERTICALLY
                    HeaderBorder.VerticalAlignment = VerticalAlignment.Stretch;
                    HeaderGrid.Height = 160; // Explicitly map a tall visual track box for the button to live in

                    if (taskPane != null)
                    {
                        taskPane.Width = 32;
                    }

                    _isPaneCollapsed = true;
                }
                else
                {
                    // --- ACTION: EXPAND ---
                    // 1. Re-reveal structural workspace objects
                    MainContentGrid.Visibility = Visibility.Visible;
                    HeaderTitle.Visibility = Visibility.Visible;
                    btnAbout.Visibility = Visibility.Visible;

                    // 2. Return text vectors to zero degree alignment reading baseline
                    BtnCollapseToggle.Content = "▶ Collapse Pane";
                    ButtonRotation.Angle = 0;
                    BtnCollapseToggle.Padding = new Thickness(8, 3, 8, 3);
                    BtnCollapseToggle.Margin = new Thickness(0);

                    // 3. Reset layout tracking positions back to defaults
                    Grid.SetColumnSpan(BtnCollapseToggle, 1);
                    Grid.SetColumn(BtnCollapseToggle, 1);

                    // CLEAR THE HEIGHT OVERRIDES
                    HeaderGrid.Height = double.NaN; // Setting to NaN tells WPF to revert to standard "Auto" sizing behavior

                    if (taskPane != null)
                    {
                        taskPane.Width = _originalPaneWidth > 45 ? _originalPaneWidth : 320;
                    }

                    _isPaneCollapsed = false;
                }
            }
            catch (Exception)
            {
                MainContentGrid.Visibility = _isPaneCollapsed ? Visibility.Visible : Visibility.Collapsed;
                _isPaneCollapsed = !_isPaneCollapsed;
            }
        }

        private void BtnAbout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Instantiates and shows our newly designed dialog safely centered over Word
                winAbout aboutWindow = new winAbout();
                aboutWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to open About page: {ex.Message}");
            }
        }
        public void PopulateContextualItems(ScanScope scope)
        {
            try
            {
                if (_isPaneCollapsed)
                    return;

                // DEFENSIVE GUARD: Ensure Word actually has a live, accessible document open
                if (Globals.ThisAddIn.Application.Documents.Count == 0)
                {
                    // No documents are open in Word yet (e.g., startup phase). Safe exit!
                    return;
                }

                Word.Document doc = null;
                try
                {
                    doc = Globals.ThisAddIn.Application.ActiveDocument;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Word is in a transient state where a document exists but isn't active yet
                    return;
                }

                if (doc == null) return;

                // ADDITIONAL GUARD: Make sure your ListBoxes are fully loaded in the WPF visual tree
                // (Prevents null errors if WPF hasn't finished painting the controls on startup)
                if (ListBoxNearbyObjects == null || ListBoxAllObjects == null || ListBoxCategory == null)
                {
                    return;
                }

                // 1. DYNAMICALLY ASSIGN THE PAGE BOUNDARIES BASED ON SCOPE
                int startPage;
                int endPage;

                if (scope == ScanScope.NearbyOnly)
                {
                    // Fetch total page count safely for boundary checks
                    int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                    int currentCursorPage = (int)Globals.ThisAddIn.Application.Selection.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                    // Your existing 3-page localized radius calculation logic
                    startPage = Math.Max(1, currentCursorPage - 1);
                    endPage = Math.Min(totalPages, currentCursorPage + 1);
                }
                else
                {
                    // FULL DOCUMENT: Span from the first page straight to the end
                    startPage = 1;
                    endPage = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                }

                // 2. Get the target text range spanning across the three pages
                Word.Range searchRange = GetRangeForPages(doc, startPage, endPage);
                if (searchRange == null) return;

                List<CrossRefTargetItem> foundItems = new List<CrossRefTargetItem>();

                // 3a. Scan for Headings inside this range
                foreach (Word.Paragraph para in searchRange.Paragraphs)
                {
                    Word.Style style = para.get_Style() as Word.Style;
                    if (style != null && style.NameLocal.StartsWith("Heading"))
                    {
                        string text = para.Range.Text.Trim('\r', '\n', ' ', '\t');
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            // 1. Extract the automated list/heading number prefix (e.g., "1.1", "Appendix A")
                            string headingNumber = para.Range.ListFormat.ListString;

                            // 2. Combine them cleanly if a number exists, otherwise fallback to just the text
                            string fullDisplayText = !string.IsNullOrWhiteSpace(headingNumber)
                                ? $"{headingNumber} {text}"
                                : text;

                            foundItems.Add(new CrossRefTargetItem
                            {
                                DisplayText = fullDisplayText,
                                Category = "Heading",
                                WordRange = para.Range
                            });
                        }
                    }
                }

                // 4. Scan for objects
                // Word fields (like Seq fields used in captions) are the most reliable way to find them
                foreach (Word.Field field in searchRange.Fields)
                {
                    if (field.Type == Word.WdFieldType.wdFieldSequence)
                    {
                        string fieldCode = field?.Code?.Text?.ToUpper();
                        Word.Paragraph parentPara = field.Result.Paragraphs[1];
                        string fullCaptionText = parentPara.Range.Text.Trim('\r', '\n', ' ', '\t');

                        if (fieldCode != null)
                        {
                            if (fieldCode.Contains("TABLE"))
                            {
                                foundItems.Add(new CrossRefTargetItem
                                {
                                    DisplayText = fullCaptionText,
                                    Category = "Table",
                                    WordRange = parentPara.Range
                                });
                            }
                            else if (fieldCode.Contains("FIGURE"))
                            {
                                foundItems.Add(new CrossRefTargetItem
                                {
                                    DisplayText = fullCaptionText,
                                    Category = "Figure",
                                    WordRange = parentPara.Range
                                });
                            }
                            else if (fieldCode.Contains("EQUATION"))
                            {
                                foundItems.Add(new CrossRefTargetItem
                                {
                                    DisplayText = fullCaptionText,
                                    Category = "Equation",
                                    WordRange = parentPara.Range
                                });
                            }
                        }
                    }
                }

                // 5. Scan for Footnotes, EndNotes, and Bookmarks
                foreach (Word.Footnote footnote in searchRange.Footnotes)
                {
                    string footnoteText = footnote.Range.Text?.Trim('\r', '\n', ' ', '\t');
                    foundItems.Add(new CrossRefTargetItem
                    {
                        DisplayText = footnoteText,
                        Category = "Footnote",
                        WordRange = footnote.Range
                    });
                }

                foreach (Word.Endnote endnote in searchRange.Endnotes)
                {
                    string endnoteText = endnote.Range.Text?.Trim('\r', '\n', ' ', '\t');
                    foundItems.Add(new CrossRefTargetItem
                    {
                        DisplayText = endnoteText,
                        Category = "Endnote",
                        WordRange = endnote.Range
                    });
                }

                foreach (Word.Bookmark bookmark in searchRange.Bookmarks)
                {
                    // avoid zero length bookmarks as there is nothing to insert into the text
                    if (bookmark.Range.Text.Length > 0)
                    {
                        string bookmarkText = bookmark.Range.Text?.Trim('\r', '\n', ' ', '\t');
                        foundItems.Add(new CrossRefTargetItem
                        {
                            DisplayText = bookmarkText,
                            Category = "Bookmark",
                            WordRange = bookmark.Range
                        });
                    }
                }

                // 6. Remove duplicates probably due to tracked changes
                // Group by the display text and select the first object instance from each unique group
                foundItems = System.Linq.Enumerable.ToList(
                    System.Linq.Enumerable.Select(
                        System.Linq.Enumerable.GroupBy(foundItems, item => new { item.DisplayText, item.Category }),
                        group => System.Linq.Enumerable.First(group)
                    )
                );

                // -------------------------------------------------------------
                // 7. ROUTING: Direct the results to the correct UI ListBox container
                // -------------------------------------------------------------
                if (scope == ScanScope.NearbyOnly)
                {
                    ListBoxNearbyObjects.Items.Clear();
                    foreach (var item in foundItems)
                    {
                        ListBoxNearbyObjects.Items.Add(item);
                    }
                }
                else
                {
                    ListBoxAllObjects.Items.Clear();

                    // 1. Get the newly selected item text safely
                    if (ListBoxCategory.SelectedItem is ListBoxItem selectedBoxItem)
                    {
                        string selectedCategory = selectedBoxItem.Content.ToString();
                        List<CrossRefTargetItem> filteredItems;

                        if (selectedCategory == "All")
                            filteredItems = foundItems;
                        else
                            // 3. Filter your master list to find objects where the category matches
                            filteredItems = System.Linq.Enumerable.ToList(
                                System.Linq.Enumerable.Where(foundItems, item =>
                                    string.Equals(item.Category, selectedCategory, StringComparison.OrdinalIgnoreCase)
                                )
                            );

                        // 4. Bind or pipe the results right directly into ListBox3
                        foreach (var item in filteredItems)
                        {
                            // If you use basic strings, add item.DisplayText. 
                            // If you use custom objects, add the object directly (relying on its .ToString() override)
                            ListBoxAllObjects.Items.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error scanning document: {ex.Message}");
            }
        }


        // Helper to generate a Word Range spanning from startPage to endPage
        private Word.Range GetRangeForPages(Word.Document doc, int startPage, int endPage)
        {
            Word.Range startRange = null;
            Word.Range endRange = null;

            try
            {
                // Go to the beginning of the start page
                object whatPage = Word.WdGoToItem.wdGoToPage;
                object whichPage = Word.WdGoToDirection.wdGoToAbsolute;
                object countStart = startPage;

                startRange = doc.GoTo(ref whatPage, ref whichPage, ref countStart);

                // Go to the end of the target end page
                object countEnd = endPage + 1;
                endRange = doc.GoTo(ref whatPage, ref whichPage, ref countEnd);
            }
            catch
            {
                // If it fails to find endPage + 1 (e.g. end of document), default to document end
            }

            int startPos = startRange?.Start ?? 0;
            int endPos = endRange?.Start ?? doc.Content.End;

            // Return safe merged range bounds
            return doc.Range(startPos, endPos);
        }

        // Call this from VSTO to explicitly pass Word's current theme colors
        public void ApplyOfficeTheme(Color backgroundColor, Color textColor, Color borderHorizontal)
        {
            var wordText = new SolidColorBrush(textColor);
            var wordBorder = new SolidColorBrush(borderHorizontal);
            var wordBackground = new SolidColorBrush(backgroundColor);

            // 1. Set the root control background
            this.Background = wordBackground;

            // 2. Set global font properties on the root control wrapper
            // Word's task pane UI natively leverages Segoe UI at 11.5pt / 12px sizes
            this.FontFamily = new FontFamily("Segoe UI");
            this.FontSize = 12;

            // 3. Walk the layout tree and paint everything matching our target types
            ApplyThemeToLogicalTree(this, wordBackground, wordText, wordBorder);
        }

        public static void ApplyThemeToLogicalTree(DependencyObject target, Brush bgBrush, Brush textBrush, Brush borderBrush)
        {
            if (target == null) return;

            // 1. STRIKE A SAFE BALANCE: Update colors based on the clean Logical Control Type
            if (target is RadioButton || target is CheckBox)
            {
                var toggleControl = (System.Windows.Controls.Control)target;
                toggleControl.Foreground = textBrush;
                toggleControl.IsEnabled = true;
                // Do not touch backgrounds/borders to keep the native dot/check crisp
            }
            else if (target is System.Windows.Controls.ListBox listBox)
            {
                listBox.Background = bgBrush;
                listBox.Foreground = textBrush;
                listBox.BorderBrush = borderBrush;
            }
            else if (target is System.Windows.Controls.GroupBox groupBox)
            {
                groupBox.Foreground = textBrush;
                groupBox.BorderBrush = borderBrush;
                groupBox.Background = System.Windows.Media.Brushes.Transparent;
            }
            else if (target is System.Windows.Controls.Grid grid)
            {
                // Explicitly paint the background of your grids
                grid.Background = bgBrush;
            }
            else if (target is System.Windows.Controls.TextBlock textBlock)
            {
                textBlock.Foreground = textBrush;
            }
            else if (target is System.Windows.Controls.Label label)
            {
                label.Foreground = textBrush;
            }
            else if (target is System.Windows.Controls.Button btn)
            {
                if (btn.Name != "BtnCollapseToggle")
                {
                    btn.Background = bgBrush;
                    btn.Foreground = textBrush;
                    btn.BorderBrush = borderBrush;
                }
            }
            else if (target is System.Windows.Controls.Border border)
            {
                border.BorderBrush = borderBrush;
                if (border.Name != "HeaderBorder")
                {
                    border.Background = bgBrush;
                }
            }

            // 2. SAFE SCANNING: Use LogicalTreeHelper to bypass hidden Win32 control wrappers
            System.Collections.IEnumerable children = LogicalTreeHelper.GetChildren(target);
            foreach (object child in children)
            {
                if (child is DependencyObject childDependencyObject)
                {
                    ApplyThemeToLogicalTree(childDependencyObject, bgBrush, textBrush, borderBrush);
                }
            }
        }

        private void ListBoxNearby_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 1. Ensure a valid item was double-clicked
            CrossRefTargetItem selectedItem = ListBoxNearbyObjects.SelectedItem as CrossRefTargetItem;
            if (selectedItem == null)
                return;

            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app.ActiveDocument;
                Word.Selection currentSelection = app.Selection;

                // 2. Determine Word's Reference Type
                object referenceType = ConvertCategoryToReferenceType(selectedItem.Category);
                // Clean the display text string up to ensure safe matching bounds
                string cleanedTargetName = selectedItem.DisplayText?
                    .Replace("\r", "")
                    .Replace("\n", "")
                    .Replace("\u001e", "-") // Replaces Word's non-breaking hyphen with a standard hyphen
                    .Replace("\u001f", "-") // Replaces Word's optional hyphen (just in case)
                    .Trim();

                // 3. Retrieve Word's hidden matching list of cross-reference items for this specific type
                // This returns a 1-based string array exactly matching Word's "Cross-reference" dialog box
                System.Array wordItems = doc.GetCrossReferenceItems(ref referenceType) as System.Array;

                object referenceItemIndex = null;

                // 4. Find the matching 1-based index position from the array
                for (int i = 1; i <= wordItems.Length; i++)
                {
                    string currentWordItem = wordItems.GetValue(i).ToString().Replace("\r", "").Replace("\n", "").Trim();

                    // This ensures a match ignoring the heading number
                    if (currentWordItem.EndsWith(cleanedTargetName, StringComparison.OrdinalIgnoreCase))
                    {
                        referenceItemIndex = i;
                        break;
                    }
                }

                // 5. Fallback handler if a perfect string layout match isn't located
                if (referenceItemIndex == null)
                {
                    MessageBox.Show("Word could not map the item text to its internal document collection index list.",
                                    "Mapping Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 6. Check Hyperlink configuration setting
                object insertAsHyperlink = CheckBoxInsertAsHyperlink.IsChecked ?? true;

                // Mandatory optional parameters for Interop compilation safety
                object separateNumbers = false;
                object tableAttributes = false;

                // 7. Execute the Cross Reference Insertion right at the cursor position
                // CRITICAL EXCEPTION: Handle Heading Number + Text combined insertion
                if (RadioInsertNumberCaption.IsChecked == true && selectedItem.Category == "Heading")
                {
                    // First Call: Insert the full heading number (e.g., "1.1.3")
                    Word.WdReferenceKind numberKind = Word.WdReferenceKind.wdNumberFullContext;
                    currentSelection.InsertCrossReference(
                        ref referenceType, numberKind, ref referenceItemIndex,
                        ref insertAsHyperlink, ref separateNumbers, ref tableAttributes
                    );

                    // Move the cursor past the newly inserted field and type a space separator
                    currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    currentSelection.TypeText(" ");

                    // Second Call: Insert the heading text content (e.g., "Introduction")
                    Word.WdReferenceKind textKind = Word.WdReferenceKind.wdContentText;
                    currentSelection.InsertCrossReference(
                        ref referenceType, textKind, ref referenceItemIndex,
                        ref insertAsHyperlink, ref separateNumbers, ref tableAttributes
                    );
                }
                else
                {
                    // Standard Single Call for all other configurations (and non-heading categories)
                    Word.WdReferenceKind referenceKind = DetermineReferenceKind(selectedItem.Category);

                    currentSelection.InsertCrossReference(
                        ref referenceType, referenceKind, ref referenceItemIndex,
                        ref insertAsHyperlink, ref separateNumbers, ref tableAttributes
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not insert cross-reference: {ex.Message}",
                                "Insertion Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // Converts our custom Category string to Word's explicit Reference Type Enum/String parameter
        private object ConvertCategoryToReferenceType(string category)
        {
            switch (category)
            {
                case "Heading":
                    return Word.WdReferenceType.wdRefTypeHeading;
                case "Table":
                    return "Table";     // Handled as string custom caption match
                case "Figure":
                    return "Figure";    // Handled as string custom caption match
                case "Equation":
                    return "Equation";  // Handled as string custom caption match
                case "Footnote":
                    return Word.WdReferenceType.wdRefTypeFootnote;
                case "Endnote":
                    return Word.WdReferenceType.wdRefTypeEndnote;
                default:
                    return Word.WdReferenceType.wdRefTypeHeading;
            }
        }

        // Evaluates user UI selections to dictate exactly what string content is written into the document
        private Word.WdReferenceKind DetermineReferenceKind(string category)
        {
            // CASE 1: User wants JUST the text description (Excludes numbers and labels entirely)
            if (RadioInsertCaption.IsChecked == true)
            {
                switch (category)
                {
                    case "Heading":
                        return Word.WdReferenceKind.wdContentText;           // Inserts: "Introduction"
                    case "Footnote":
                        return Word.WdReferenceKind.wdFootnoteNumber;        // Fallback
                    case "Endnote":
                        return Word.WdReferenceKind.wdEndnoteNumber;          // Fallback
                    default:
                        return Word.WdReferenceKind.wdOnlyCaptionText;        // Inserts: "My Sample Data Layout"
                }
            }

            // CASE 2: User wants BOTH the number/label AND the descriptive text together
            else if (RadioInsertNumberCaption.IsChecked == true)
            {
                switch (category)
                {
                    case "Heading":
                        return Word.WdReferenceKind.wdNumberFullContext;          // Inserts: "1.1 Introduction"
                    case "Footnote":
                        return Word.WdReferenceKind.wdFootnoteNumber;
                    case "Endnote":
                        return Word.WdReferenceKind.wdEndnoteNumber;
                    default:
                        return Word.WdReferenceKind.wdEntireCaption;           // Inserts: "Figure 1: Flowchart Diagram"
                }
            }

            // CASE 3: Default (RadioInsertNumber is Checked) -> User wants ONLY numbers/labels
            else
            {
                switch (category)
                {
                    case "Heading":
                        // FIXED: Using wdNumberFullContext to cleanly print "1.1.3" instead of an imaginary enum!
                        return Word.WdReferenceKind.wdNumberFullContext;
                    case "Footnote":
                        return Word.WdReferenceKind.wdFootnoteNumber;
                    case "Endnote":
                        return Word.WdReferenceKind.wdEndnoteNumber;
                    default:
                        return Word.WdReferenceKind.wdOnlyLabelAndNumber;     // Inserts: "Table 1" or "Figure 2"
                }
            }
        }

        private void ListBoxCategory_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            try
            {
                // 1. Get the newly selected item text safely
                PopulateContextualItems(scope: ScanScope.FullDocument);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Category switch error: {ex.Message}");
            }
        }

        private void ListBoxCategory_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Get the newly selected item text safely
                PopulateContextualItems(scope: ScanScope.FullDocument);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Category switch error: {ex.Message}");
            }
        }
    }
}
